namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_エリアマスタ
    {
        [Key]
        public int ROW_ID { get; set; }

        public string area_no { get; set; }

        public string area_name { get; set; }
    }
}
