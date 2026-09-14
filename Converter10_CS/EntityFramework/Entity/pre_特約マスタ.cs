namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_特約マスタ
    {
        [Key]
        public int ROW_ID { get; set; }
        public string tokuyaku_grpno { get; set; }
        public string tokuyaku_no { get; set; }
        public string tokuyaku_title { get; set; }
        public string tokuyaku_template { get; set; }
    }
}
