namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_自社担当者情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string logonuser_no { get; set; }

        public string logonuser_name { get; set; }

        public string logonuser_kana { get; set; }

        public string tel1 { get; set; }

        public string tel2 { get; set; }

        public string mobiletel1 { get; set; }

        public string fax1 { get; set; }

        public string mailaddress1 { get; set; }

        public string mobilemailaddress1 { get; set; }

        public string takkenmenkyo { get; set; }

        public string biko { get; set; }
    }
}
