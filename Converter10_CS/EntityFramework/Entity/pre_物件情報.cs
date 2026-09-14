namespace Converter10.EntityFramework.Entity
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class pre_物件情報
    {
        [Key]
        public int ROW_ID { get; set; }

        public string bk_no { get; set; }

        public string bk_name { get; set; }

        public string bk_kana { get; set; }

        public string tatemono_sikibetu { get; set; }

        public string post_code { get; set; }

        public string addr_kenno { get; set; }

        public string addr_sino { get; set; }

        public string addr_cyo { get; set; }

        public string addr_cyome { get; set; }

        public string addr_banti { get; set; }

        public string addr_etc { get; set; }

        public string bk_ruinokbn { get; set; }

        public string gps_wgsido { get; set; }

        public string gps_wgskeido { get; set; }

        public string syunko_ymd { get; set; }

        public string kaidate { get; set; }

        public string tika { get; set; }

        public string elevator_flg { get; set; }

        public string elevator_number { get; set; }

        public string kozo_nokbn { get; set; }

        public string moto_gy_fudono { get; set; }

        public string biko_kihon { get; set; }

        public string kosu_total { get; set; }

        public string men_nobeyuka { get; set; }

        public string men_sikiti { get; set; }

        public string men_parking { get; set; }

        public string men_yukatoki { get; set; }

        public string men_sikititoki { get; set; }

        public string men_kentiku { get; set; }

        public string men_kentikutoki { get; set; }

        public string yatin_jisansaki { get; set; }

        public string yatin_kozano { get; set; }

        public string yatin_kykozaflg { get; set; }

        public string yatin_kykozano { get; set; }

        public string kozo_taikakbn { get; set; }

        public string gy_sekono { get; set; }

        public string gy_hosyuno { get; set; }

        public string kanri_hosiki { get; set; }

        public string kanri_gyname { get; set; }

        public string kanri_gytanto { get; set; }

        public string kanri_gytel { get; set; }

        public string kanrinin_gyomukeitai { get; set; }

        public string kanrinin_tel { get; set; }

        public string jisya_no { get; set; }

        public string kotu_sonota1 { get; set; }

        public string kotu_sonota1kyori { get; set; }

        public string kotu_sonota2 { get; set; }

        public string kotu_sonota2kyori { get; set; }

        public string denki_gy_lifeno { get; set; }

        public string water_gy_lifeno { get; set; }

        public string gas_gy_lifeno { get; set; }

        public string haisui_gy_lifeno { get; set; }

        public string toyu_gy_lifeno { get; set; }

        public string lifeline1_gy_lineno { get; set; }

        public string lifeline2_gy_lineno { get; set; }

        public string lifeline3_gy_lineno { get; set; }

        public string kensingyomu_umu { get; set; }

        public string parking_kanriflg { get; set; }

        public string parking_car { get; set; }

        public string parking_bike { get; set; }

        public string parking_bicycle { get; set; }

        public string parking_bicyclefreeflg { get; set; }

        public string bk_logonuser_no { get; set; }

        public string kanri_gyfax { get; set; }

        public string homeelevator_flg { get; set; }

        public string kanrinin_name { get; set; }

        public string parking_caraki { get; set; }

        public string parking_bikeaki { get; set; }

        public string parking_bicycleaki { get; set; }

        public string syogaku_name { get; set; }

        public string syogaku_kyori { get; set; }

        public string cyugaku_name { get; set; }

        public string cyugaku_kyori { get; set; }

        public string bk_area_no { get; set; }

        public string emergencyelevator_flg { get; set; }

        public string memo1 { get; set; }

        public string memo2 { get; set; }

        public string memo3 { get; set; }

        public string memo4 { get; set; }

        public string memo5 { get; set; }
    }
}
