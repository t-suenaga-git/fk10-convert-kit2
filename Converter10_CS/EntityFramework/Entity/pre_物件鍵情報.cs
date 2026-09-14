namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件鍵情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string kagi_no { get; set; }

        public string honsu { get; set; }

        public string biko { get; set; }

        public string hokan { get; set; }

        public string gyshare { get; set; }
    }
}
