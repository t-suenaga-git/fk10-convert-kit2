namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約者口座情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string kys_no { get; set; }

        public string kys_kozano { get; set; }

        public string kinyu_no { get; set; }

        public string kinyu_tenno { get; set; }

        public string koza_syubetu { get; set; }

        public string koza_bango { get; set; }

        public string koza_meigi { get; set; }

        public string koza_meigikana { get; set; }

        public string yucyokoza_kigo1 { get; set; }

        public string yucyokoza_kigo2 { get; set; }

        public string yucyokoza_bango { get; set; }

        public string fkae_no { get; set; }

        public string fkae_tesugak { get; set; }

        public string fkae_kysbango { get; set; }

        public string biko_koza { get; set; }

        public string koza_yucyoflg { get; set; }
    }
}
