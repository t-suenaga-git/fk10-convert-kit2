namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_ライフライン業者情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string gy_lifelineno { get; set; }

        public string gy_lifelinename { get; set; }

        public string gy_lifelinekana { get; set; }

        public string eigyosyo { get; set; }

        public string post_code { get; set; }

        public string addr1 { get; set; }

        public string addr2 { get; set; }

        public string tel1 { get; set; }

        public string tel2 { get; set; }

        public string Fax { get; set; }

        public string mobiletel1 { get; set; }

        public string mobiletel2 { get; set; }

        public string biko_kihon { get; set; }

        public string lifeline_denkiflg { get; set; }

        public string lifeline_josuidoflg { get; set; }

        public string lifeline_gasflg { get; set; }

        public string lifeline_toyuflg { get; set; }

        public string lifeline_other1flg { get; set; }

        public string lifeline_haisuiflg { get; set; }
    }
}
