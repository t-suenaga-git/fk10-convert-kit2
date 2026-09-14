namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_部屋所有者情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string kasi1_ow_no { get; set; }

        public string syo1_ow_no { get; set; }

        public string kasi2_ow_no { get; set; }

        public string syo2_ow_no { get; set; }

        public string syo_startymd { get; set; }

        public string syo_endymd { get; set; }
    }
}
