namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約入居者情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string nyukyosya_cnt { get; set; }

        public string kys_no { get; set; }

        public string name { get; set; }

        public string kana { get; set; }

        public string gender { get; set; }

        public string aidagara { get; set; }

        public string birthday { get; set; }

        public string kinmusaki { get; set; }

        public string tel { get; set; }

        public string mobiletel { get; set; }

        public string biko { get; set; }
    }
}
