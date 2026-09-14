namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_家主口座情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string ow_no { get; set; }

        public string ow_kozano { get; set; }

        public string kinyu_no { get; set; }

        public string kinyu_tenno { get; set; }

        public string koza_syubetu { get; set; }

        public string koza_bango { get; set; }

        public string koza_meigi { get; set; }

        public string koza_meigikana { get; set; }

        public string yucyokoza_kigo1 { get; set; }

        public string yucyokoza_kigo2 { get; set; }

        public string yucyokoza_bango { get; set; }

        public string biko_koza { get; set; }

        public string biko_furikomi { get; set; }

        public string biko_sgfirai { get; set; }
    }
}
