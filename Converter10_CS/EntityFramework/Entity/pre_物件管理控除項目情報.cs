namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件管理控除項目情報
    {
        [Key]
        public int ROW_ID { get; set; }    
        public string bk_no { get; set; }
        public string hy_no { get; set; }
        public string taisyokbn { get; set; }
        public string nkin_sortorder { get; set; }
        public string nkin_name { get; set; }
        public string sosai_flg { get; set; }
        public string kojo_gak { get; set; }
        public string kojo_gakzeikbn { get; set; }
        public string tateazu_flg { get; set; }
        public string kojo_zeigak { get; set; }
    }
}
