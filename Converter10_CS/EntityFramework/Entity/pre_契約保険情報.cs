namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約保険情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string hoken_no { get; set; }

        public string gy_hokenno { get; set; }

        public string ky_ymd { get; set; }

        public string hoken_gak { get; set; }

        public string biko { get; set; }

        public string syoken_bango { get; set; }
    }
}
