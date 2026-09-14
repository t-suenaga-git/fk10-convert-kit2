namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約車情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string car_cnt { get; set; }

        public string carmaker { get; set; }

        public string carname { get; set; }

        public string carcolor { get; set; }

        public string carnumber { get; set; }

        public string biko { get; set; }

        public string car_kukaku { get; set; }
    }
}
