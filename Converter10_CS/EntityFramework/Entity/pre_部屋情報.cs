namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_部屋情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string hy_no { get; set; }

        public string hy_ruinokbn { get; set; }

        public string madori { get; set; }

        public string madori_biko { get; set; }

        public string men_senyujitu { get; set; }

        public string men_yuka { get; set; }

        public string men_senyutouki { get; set; }

        public string men_toki { get; set; }

        public string men_balcony { get; set; }

        public string kaiyaku_months { get; set; }

        public string syozai_kaisu1 { get; set; }

        public string syozai_tikaflg1 { get; set; }

        public string mukikbn { get; set; }

        public string balcony_mukikbn { get; set; }

        public string nyukyo_jokyokbn { get; set; }

        public string nyukyo_jokyomemo { get; set; }

        public string nyukyo_syunflg { get; set; }

        public string nyukyo_ym { get; set; }

        public string nyukyo_syunkbn { get; set; }

        public string nyukyo_joken { get; set; }

        public string kakunin_ymd { get; set; }

        public string bosyu_startflg { get; set; }

        public string torihiki_taiyokbn { get; set; }

        public string torihiki_kyakutuke { get; set; }

        public string moto_gy_fudono { get; set; }

        public string biko { get; set; }

        public string hosyo_gyno { get; set; }

        public string hosyo_naiyo { get; set; }

        public string kyrui_nokbn { get; set; }

        public string keiyaku_kikan { get; set; }

        public string keiyaku_kijitu { get; set; }

        public string parking_biko { get; set; }

        public string parking_bikebiko { get; set; }

        public string parking_cyurinbiko { get; set; }

        public string shared_salespoint { get; set; }

        public string hosyo_kbn { get; set; }

        public string jisya_no { get; set; }

        public string jisya_tanto { get; set; }

        public string nextnkin_kosindefault { get; set; }

        public string toki_ymd { get; set; }

        public string siyo_mokuteki { get; set; }

        public string nyukyo_jikikbn { get; set; }

        public string torihiki_jisyakbn { get; set; }

        public string parking_gakkbn { get; set; }

        public string parking_gak { get; set; }

        public string parking_gakzei { get; set; }

        public string parking_tintaisu { get; set; }

        public string genjyo_naiyo { get; set; }

        public string nyukyo_naiyo { get; set; }

        public string etc_naiyo { get; set; }

        public string memo1 { get; set; }

        public string memo2 { get; set; }

        public string memo3 { get; set; }

        public string memo4 { get; set; }

        public string memo5 { get; set; }
    }
}
