namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_部屋鍵タイトルマスタ
    {
        [Key]
        public int ROW_ID { get; set; }

        public string kagi_no { get; set; }

        public string kagi_name { get; set; }
    }
}
