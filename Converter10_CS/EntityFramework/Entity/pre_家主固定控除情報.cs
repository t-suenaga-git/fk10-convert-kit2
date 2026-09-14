namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_家主固定控除情報
    {
        [Key]
        public int ROW_ID { get; set; }
        public string bk_no { get; set; }
        public string nkin_sortorder { get; set; }
        public string nkin_name { get; set; }
        public string hy_no { get; set; }
        public string sosai_flg { get; set; }
        public string gak { get; set; }
        public string zei_kbn { get; set; }
        public string zei_gak { get; set; }
        public string nkbn_yotei { get; set; }
        public string yatin_kozano { get; set; }
        public string gtstart_ym { get; set; }
        public string gtend_ym { get; set; }
        public string tateazu_flg { get; set; }
        public string sq_mmkbn { get; set; }
        public string sq_interval { get; set; }
        public string biko { get; set; }
    }
}
