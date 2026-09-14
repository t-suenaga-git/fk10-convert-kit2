
namespace Converter10.Njc.Model
{

    #region クレーム基本情報モデル

    public class Claimdata_Model
    {

        public string Vari_Claim_no { get; set; }
        public string Vari_Status { get; set; }
        public string Vari_Title { get; set; }
        public string Vari_Uke_ymd { get; set; }
        public string Vari_Uke_logonuser_no { get; set; }
        public string Vari_Uke_name { get; set; }
        public string Vari_Uke_tel1 { get; set; }
        public string Vari_Uke_tel2 { get; set; }
        public string Vari_Renraku_timestart { get; set; }
        public string Vari_Renraku_timeend { get; set; }
        public string Vari_Kinkyu_kbn { get; set; }
        public string Vari_Taio_limitymd { get; set; }
        public string Vari_Kasyo_ruino { get; set; }
        public string Vari_Claim_ruino { get; set; }
        public string Vari_Uke_report { get; set; }
        public string Vari_Taio_logonuser_no { get; set; }
        public string Vari_Emailbiko { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Reform_guid { get; set; }
        public string Vari_Taiosaki_kbn { get; set; }
        public string Vari_Taiosaki_gy_no { get; set; }
        public string Vari_Taiosaki_tantoname { get; set; }
        public string Vari_Taiosaki_tel { get; set; }
        public string Vari_Taio_finishymd { get; set; }
        public string Vari_Taio_report { get; set; }
        public string Vari_Hutan1_kbn { get; set; }
        public string Vari_Hutan1_gaketc { get; set; }
        public string Vari_Hutan2_kbn { get; set; }
        public string Vari_Hutan2_gaketc { get; set; }
        public string Vari_Hutan3_kbn { get; set; }
        public string Vari_Hutan3_gaketc { get; set; }
        public string Vari_Hutanbiko { get; set; }
        public string Vari_Jisya_no { get; set; }
        public string Vari_Jisya_name { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Uke_hm { get; set; }
        public string Vari_Taio_finishhm { get; set; }
        public string Vari_Uke_namesjis { get; set; }
        public string Vari_Taiosaki_tantonamesjis { get; set; }
        public string Vari_Ow_no { get; set; }

    }

    #endregion

    #region クレーム対応履歴情報モデル

    public class Claim_taio_Model
    {

        public string Vari_Claim_no { get; set; }
        public string Vari_Taio_no { get; set; }
        public string Vari_Taio_ymd { get; set; }
        public string Vari_Taio_logonuser_no { get; set; }
        public string Vari_Taio_report { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Taio_hm { get; set; }

    }

    #endregion

    #region クレーム関連ファイル情報モデル

    public class Claimdata_relfile_Model
    {

        public string Vari_Claim_no { get; set; }
        public string Vari_File_no { get; set; }
        public string Vari_Fullpath { get; set; }
        public string Vari_Addtime { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}