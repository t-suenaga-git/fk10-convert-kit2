namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約者保証人情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string kys_no { get; set; }

        public string hosyo_no { get; set; }

        public string hosyo_name { get; set; }

        public string hosyo_kana { get; set; }

        public string keisyo { get; set; }

        public string post_code { get; set; }

        public string addr1 { get; set; }

        public string addr2 { get; set; }

        public string tel1 { get; set; }

        public string tel2 { get; set; }

        public string fax { get; set; }

        public string mobiletel1 { get; set; }

        public string mobiletel2 { get; set; }

        public string birthday { get; set; }

        public string nensyu { get; set; }

        public string biko_hosyo { get; set; }

        public string aidagara { get; set; }

        public string kinmu_name { get; set; }

        public string kinmu_kana { get; set; }

        public string kinmu_postcode { get; set; }

        public string kinmu_addr1 { get; set; }

        public string kinmu_addr2 { get; set; }

        public string kinmu_tel1 { get; set; }

        public string kinmu_tel2 { get; set; }

        public string kinmu_fax { get; set; }

        public string kinmu_gyosyu { get; set; }

        public string kinmu_busyo { get; set; }

        public string kinmu_nyuryokuym { get; set; }

        public string biko_kinmu { get; set; }
    }
}
