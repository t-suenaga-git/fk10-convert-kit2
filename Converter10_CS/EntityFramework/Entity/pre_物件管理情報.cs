namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件管理情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string kanri_keitaikbn { get; set; }

        public string ikkatu_kanriflg { get; set; }

        public string cyukai_kyflg { get; set; }

        public string cyukai_koflg { get; set; }

        public string cyukai_kaiflg { get; set; }

        public string soymd1_gaitoukbn { get; set; }

        public string soymd1_simekbn { get; set; }

        public string soymd1_sokintukikbn { get; set; }

        public string soymd1_sokinsimekbn { get; set; }

        public string ikkatu_kbn { get; set; }

        public string ikkatu_bkgak { get; set; }

        public string kanritesu_kbn { get; set; }

        public string kanritesu_cyosyukbn { get; set; }

        public string kanri_reigaiky { get; set; }

        public string kanri_reigaikai { get; set; }

        public string hyteigaku_kbn { get; set; }

        public string hyteigaku_kanrigak { get; set; }

        public string bkteigaku_kanrigak { get; set; }

        public string kanri_taxflg { get; set; }

        public string biko_basic { get; set; }

        public string kanritesu_zeiumukbn { get; set; }

        public string no_soruleflg { get; set; }

        public string cyukai_zuijisokin { get; set; }
    }
}
