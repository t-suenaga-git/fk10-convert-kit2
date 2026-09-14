namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件管理入金項目情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string taisyokbn { get; set; }

        public string nkin_sortorder { get; set; }

        public string nkin_name { get; set; }

        public string sokin_rit { get; set; }

        public string kanrigak_rit { get; set; }

        public string hosyo_flg { get; set; }

        public string so_no { get; set; }
    }
}
