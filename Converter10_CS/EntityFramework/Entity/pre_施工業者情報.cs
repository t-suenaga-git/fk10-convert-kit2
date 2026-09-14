namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_施工業者情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string gy_sekono { get; set; }

        public string gy_sekoname { get; set; }

        public string gy_sekokana { get; set; }

        public string eigyosyo { get; set; }

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
    }
}
