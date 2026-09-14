namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_契約情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string status { get; set; }

        public string status_ymd { get; set; }

        public string cyukai_gy_fudono { get; set; }

        public string syunin_no { get; set; }

        public string kaiyaku_months { get; set; }

        public string cyukai_tantoname { get; set; }

        public string kystart_ymd { get; set; }

        public string kyend_ymd { get; set; }

        public string ky_bango { get; set; }

        public string ky_ymd { get; set; }

        public string kanri_gy_fudono { get; set; }

        public string gy_hosyono { get; set; }

        public string hosyo_naiyo { get; set; }

        public string kokyaku_bango { get; set; }

        public string kyrui_no { get; set; }

        public string siyo_mokuteki { get; set; }

        public string kosin_umu { get; set; }

        public string yatin_kbn { get; set; }

        public string fkae_startym { get; set; }

        public string yatin_kozakbn { get; set; }

        public string yatin_kozano { get; set; }

        public string maitukiyatin_kozano { get; set; }

        public string biko { get; set; }

        public string nextky_startymd { get; set; }

        public string nextky_endymd { get; set; }

        public string soyotei_ymd { get; set; }

        public string kanritesu_flg { get; set; }

        public string kanritesu_gak { get; set; }

        public string nextnkinset_kbn { get; set; }

        public string biko2 { get; set; }

        public string hoken_biko { get; set; }

        public string nkinsime_ymd { get; set; }

        public string yatin_jisansaki { get; set; }

        public string yokugetu_uketoriumu { get; set; }

        public string yokugetu_uketorimonth { get; set; }

        public string ky_logonuser_no { get; set; }

        public string sq_logonuser_no { get; set; }
    }
}
