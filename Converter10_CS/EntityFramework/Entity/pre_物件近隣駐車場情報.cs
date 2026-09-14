namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件近隣駐車場情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string parking_no { get; set; }

        public string naiyo { get; set; }

        public string kyori { get; set; }

        public string gak { get; set; }

        public string zeikbn { get; set; }
    }
}
