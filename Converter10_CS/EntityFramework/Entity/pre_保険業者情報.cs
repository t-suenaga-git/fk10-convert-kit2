namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_保険業者情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string gy_hokenno { get; set; }

        public string gy_hokenname { get; set; }

        public string gy_hokenkana { get; set; }

        public string post_code { get; set; }

        public string addr1 { get; set; }

        public string addr2 { get; set; }

        public string tel1 { get; set; }

        public string tel2 { get; set; }

        public string Fax { get; set; }

        public string mobiletel1 { get; set; }

        public string mobiletel2 { get; set; }

        public string tanto_busyoyakusyoku { get; set; }

        public string tanto_name { get; set; }

        public string tanto_kana { get; set; }

        public string mail { get; set; }

        public string mobilemail { get; set; }

        public string url { get; set; }

        public string biko_kihon { get; set; }

        public string kinyu_no { get; set; }

        public string kinyu_tenno { get; set; }

        public string koza_syubetu { get; set; }

        public string koza_bango { get; set; }

        public string koza_meigi { get; set; }

        public string koza_meigikana { get; set; }

        public string memo1 { get; set; }

        public string memo2 { get; set; }

        public string memo3 { get; set; }

        public string memo4 { get; set; }

        public string memo5 { get; set; }
    }
}
