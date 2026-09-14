namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_口座振替情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string fkae_no { get; set; }

        public string fkae_name { get; set; }

        public string fkae_kana { get; set; }

        public string fkae_fb_fkomiraino { get; set; }

        public string fkae_fb_fkomiraikana { get; set; }

        public string kamei_no { get; set; }

        public string hikiotosibi { get; set; }

        public string tesu_gak { get; set; }

        public string biko { get; set; }

        public string jisya_no { get; set; }

        public string fkae_fb_nkinukekozano { get; set; }
    }
}
