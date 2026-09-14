namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_振込依頼人情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string sgfirai_no { get; set; }

        public string sgfirai_name { get; set; }

        public string sgfirainin_code { get; set; }

        public string sgfirainin_kana { get; set; }

        public string biko { get; set; }

        public string jisya_no { get; set; }

        public string jisya_kozano { get; set; }
    }
}
