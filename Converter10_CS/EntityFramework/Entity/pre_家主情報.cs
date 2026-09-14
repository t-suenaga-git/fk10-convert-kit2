namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_家主情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string ow_no { get; set; }

        public string kojinhojin_flg { get; set; }

        public string ow_name { get; set; }

        public string ow_kana { get; set; }

        public string keisyo { get; set; }

        public string post_code { get; set; }

        public string addr1 { get; set; }

        public string addr2 { get; set; }

        public string toukiaddr1 { get; set; }

        public string toukiaddr2 { get; set; }

        public string tel1 { get; set; }

        public string tel2 { get; set; }

        public string fax { get; set; }

        public string mobiletel1 { get; set; }

        public string mobiletel2 { get; set; }

        public string mail { get; set; }

        public string mobilemail { get; set; }

        public string birthday { get; set; }

        public string nensyu { get; set; }

        public string biko_kihon { get; set; }

        public string gender { get; set; }

        public string honseki { get; set; }

        public string kinmu_name { get; set; }

        public string kinmu_kana { get; set; }

        public string kinmu_postcode { get; set; }

        public string kinmu_addr1 { get; set; }

        public string kinmu_addr2 { get; set; }

        public string kinmu_tel1 { get; set; }

        public string kinmu_tel2 { get; set; }

        public string kinmu_fax { get; set; }

        public string kinmu_gyosyu { get; set; }

        public string kinmu_busyo { get; set; }

        public string kinmu_nyuryokuym { get; set; }

        public string kinmu_nyusyaym { get; set; }

        public string url { get; set; }

        public string gyosyu { get; set; }

        public string daihyo_name { get; set; }

        public string daihyo_kana { get; set; }

        public string daihyo_yakusyoku { get; set; }

        public string daihyo_atenaflg { get; set; }

        public string tanto_name { get; set; }

        public string tanto_kana { get; set; }

        public string tanto_busyo { get; set; }

        public string tanto_atenaflg { get; set; }

        public string sihonkin { get; set; }

        public string jugyosu { get; set; }

        public string nyuryoku_ym { get; set; }

        public string torihikisaki { get; set; }

        public string renraku_name { get; set; }

        public string renraku_kana { get; set; }

        public string renraku_keisyo { get; set; }

        public string renraku_postcode { get; set; }

        public string renraku_addr1 { get; set; }

        public string renraku_addr2 { get; set; }

        public string renraku_tel1 { get; set; }

        public string renraku_tel2 { get; set; }

        public string renraku_fax { get; set; }

        public string renraku_mobiletel1 { get; set; }

        public string renraku_mobiletel2 { get; set; }

        public string renraku_aidagara { get; set; }

        public string biko_renraku { get; set; }

        public string sofu_kbn { get; set; }

        public string sofu_name { get; set; }

        public string sofu_kana { get; set; }

        public string sofu_keisyo { get; set; }

        public string sofu_postcode { get; set; }

        public string sofu_addr1 { get; set; }

        public string sofu_addr2 { get; set; }

        public string sofu_tel1 { get; set; }

        public string sofu_tel2 { get; set; }

        public string sofu_fax { get; set; }

        public string biko_sofu { get; set; }

        public string memo1 { get; set; }

        public string memo2 { get; set; }

        public string memo3 { get; set; }

        public string memo4 { get; set; }

        public string memo5 { get; set; }

        public string memo6 { get; set; }

        public string memo7 { get; set; }

        public string memo8 { get; set; }

        public string memo9 { get; set; }

        public string memo10 { get; set; }
    }
}
