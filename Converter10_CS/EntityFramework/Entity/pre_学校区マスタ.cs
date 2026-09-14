namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_学校区マスタ
    {
        [Key]
        public int ROW_ID { get; set; }
        public string ken_no { get; set; }
        public string si_no { get; set; }
        public string no { get; set; }
        public string add_cyo { get; set; }
        public string add_cyome { get; set; }
        public string add_banti { get; set; }
        public string syogaku_name { get; set; }
        public string cyugaku_name { get; set; }
    }
}
