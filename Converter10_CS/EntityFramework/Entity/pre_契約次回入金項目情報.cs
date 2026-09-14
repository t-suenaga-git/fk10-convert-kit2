namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約次回入金項目情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string tuki_kbn { get; set; }

        public string nkin_name { get; set; }

        public string nkin_recno { get; set; }

        public string nkin_sortorder { get; set; }

        public string nkin_kbn { get; set; }

        public string sq_gak { get; set; }

        public string sq_zeikbn { get; set; }

        public string sq_zeigak { get; set; }

        public string calc_kbn { get; set; }

        public string calc_nkinno { get; set; }

        public string calc_monthcnt { get; set; }

        public string sqsaki_no { get; set; }

        public string nkbn_yotei { get; set; }

        public string sq_mmkbn { get; set; }

        public string sqstart_ymd { get; set; }

        public string sq_interval { get; set; }

        public string biko { get; set; }
    }
}
