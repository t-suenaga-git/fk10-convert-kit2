namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_保険種類マスタ
    {
        [Key]
        public int ROW_ID { get; set; }

        public string hoken_ruino { get; set; }

        public string hoken_ruiname { get; set; }

        public string hoken_ruikana { get; set; }

        public string biko_kihon { get; set; }
    }
}
