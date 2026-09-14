namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_バス交通マスタ
    {
        [Key]
        public int ROW_ID { get; set; }
        public string buskotu_no { get; set; }
        public string buskotu_name { get; set; }
    }
}
