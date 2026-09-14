namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件交通情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string ensen_cnt { get; set; }

        public string setting_kotutype { get; set; }

        public string ensen_no { get; set; }

        public string eki_no { get; set; }

        public string firsttrain_flg { get; set; }

        public string kyori { get; set; }

        public string toho_min { get; set; }

        public string bus_companytoeki { get; set; }

        public string bus_stationtoeki { get; set; }

        public string bus_min { get; set; }

        public string bus_kyoritoeki { get; set; }

        public string bus_tohomintoeki { get; set; }

        public string car_kyori { get; set; }

        public string car_min { get; set; }

        public string bus_company { get; set; }

        public string bus_station { get; set; }

        public string bus_kyori { get; set; }

        public string bus_tohomin { get; set; }
    }
}
