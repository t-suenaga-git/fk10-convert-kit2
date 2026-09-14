namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_部屋共通セールスポイント情報
    {
        [Key]
        public int ROW_ID { get; set; }
        public string bk_no { get; set; }
        public string hy_no { get; set; }
        public string parts_no { get; set; }
        public string salespoint_parts { get; set; }
    }
}
