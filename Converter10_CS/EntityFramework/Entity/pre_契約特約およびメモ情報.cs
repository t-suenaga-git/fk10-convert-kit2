namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約特約およびメモ情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string genjo_tokuyaku { get; set; }

        public string syuzen_tokuyaku { get; set; }

        public string sonota_tokuyaku { get; set; }

        public string memo1 { get; set; }

        public string memo2 { get; set; }

        public string memo3 { get; set; }

        public string memo4 { get; set; }

        public string memo5 { get; set; }

        public string memo6 { get; set; }

        public string memo7 { get; set; }

        public string memo8 { get; set; }

        public string memo9 { get; set; }

        public string memo10 { get; set; }
    }
}
