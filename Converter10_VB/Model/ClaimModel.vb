Namespace Njc.Model

#Region "クレーム基本情報モデル"

    Public Class Claimdata_Model

        Public Property Vari_Claim_no As String
        Public Property Vari_Status As String
        Public Property Vari_Title As String
        Public Property Vari_Uke_ymd As String
        Public Property Vari_Uke_logonuser_no As String
        Public Property Vari_Uke_name As String
        Public Property Vari_Uke_tel1 As String
        Public Property Vari_Uke_tel2 As String
        Public Property Vari_Renraku_timestart As String
        Public Property Vari_Renraku_timeend As String
        Public Property Vari_Kinkyu_kbn As String
        Public Property Vari_Taio_limitymd As String
        Public Property Vari_Kasyo_ruino As String
        Public Property Vari_Claim_ruino As String
        Public Property Vari_Uke_report As String
        Public Property Vari_Taio_logonuser_no As String
        Public Property Vari_Emailbiko As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Reform_guid As String
        Public Property Vari_Taiosaki_kbn As String
        Public Property Vari_Taiosaki_gy_no As String
        Public Property Vari_Taiosaki_tantoname As String
        Public Property Vari_Taiosaki_tel As String
        Public Property Vari_Taio_finishymd As String
        Public Property Vari_Taio_report As String
        Public Property Vari_Hutan1_kbn As String
        Public Property Vari_Hutan1_gaketc As String
        Public Property Vari_Hutan2_kbn As String
        Public Property Vari_Hutan2_gaketc As String
        Public Property Vari_Hutan3_kbn As String
        Public Property Vari_Hutan3_gaketc As String
        Public Property Vari_Hutanbiko As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Jisya_name As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Uke_hm As String
        Public Property Vari_Taio_finishhm As String
        Public Property Vari_Uke_namesjis As String
        Public Property Vari_Taiosaki_tantonamesjis As String
        Public Property Vari_Ow_no As String

    End Class

#End Region

#Region "クレーム対応履歴情報モデル"

    Public Class Claim_taio_Model

        Public Property Vari_Claim_no As String
        Public Property Vari_Taio_no As String
        Public Property Vari_Taio_ymd As String
        Public Property Vari_Taio_logonuser_no As String
        Public Property Vari_Taio_report As String
        Public Property Vari_History As String
        Public Property Vari_Taio_hm As String

    End Class

#End Region

#Region "クレーム関連ファイル情報モデル"

    Public Class Claimdata_relfile_Model

        Public Property Vari_Claim_no As String
        Public Property Vari_File_no As String
        Public Property Vari_Fullpath As String
        Public Property Vari_Addtime As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
