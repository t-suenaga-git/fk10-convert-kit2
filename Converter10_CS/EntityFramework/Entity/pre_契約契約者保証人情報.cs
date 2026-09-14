namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約契約者保証人情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string kys1_no { get; set; }

        public string nyukyo1_flg { get; set; }

        public string kys2_no { get; set; }

        public string nyukyo2_flg { get; set; }

        public string kys3_no { get; set; }

        public string nyukyo3_flg { get; set; }

        public string kys_hosyo_no { get; set; }

        public string hosyo1_no { get; set; }

        public string hosyo2_no { get; set; }
    }
}
