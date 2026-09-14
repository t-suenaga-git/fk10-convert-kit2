namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_家賃入金口座情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string yatin_kozano { get; set; }

        public string yatin_kozaname { get; set; }

        public string yatin_kozakana { get; set; }

        public string jisya_no { get; set; }

        public string jisya_kozano { get; set; }

        public string ow_no { get; set; }

        public string ow_kozano { get; set; }

        public string yatin_kozabiko { get; set; }
    }
}
