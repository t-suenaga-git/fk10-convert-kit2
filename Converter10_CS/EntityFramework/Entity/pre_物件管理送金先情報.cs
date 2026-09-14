namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件管理送金先情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string sosaki_recno { get; set; }

        public string so_ow_no { get; set; }

        public string so_ow_kozano { get; set; }

        public string so_rit { get; set; }

        public string so_fixsogak { get; set; }
    }
}
