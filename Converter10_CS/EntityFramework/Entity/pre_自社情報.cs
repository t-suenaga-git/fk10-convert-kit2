namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_自社情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string jisya_no { get; set; }

        public string jisya_name { get; set; }

        public string jisya_kana { get; set; }

        public string jisya_name2 { get; set; }

        public string jisya_name3 { get; set; }

        public string daihyo_yakusyoku { get; set; }

        public string daihyo_name { get; set; }

        public string post_code { get; set; }

        public string addr1 { get; set; }

        public string addr2 { get; set; }

        public string tel1 { get; set; }

        public string tel2 { get; set; }

        public string mobiletel1 { get; set; }

        public string mobiletel2 { get; set; }

        public string fax { get; set; }

        public string mail { get; set; }

        public string mobilemail { get; set; }

        public string url { get; set; }

        public string biko_kihon { get; set; }

        public string menkyo_bango { get; set; }

        public string menkyo_ymd { get; set; }

        public string syunin_bango { get; set; }

        public string syunin_name { get; set; }

        public string kanrikyokai_bango { get; set; }

        public string keieikanrisi_bango { get; set; }

        public string keieikanrisi_name { get; set; }

        public string kanrigy_bango { get; set; }

        public string kanyu_dantai1 { get; set; }

        public string kanyu_dantai2 { get; set; }

        public string kanyu_dantai3 { get; set; }

        public string memo1 { get; set; }

        public string memo2 { get; set; }

        public string memo3 { get; set; }

        public string memo4 { get; set; }

        public string memo5 { get; set; }
    }
}
