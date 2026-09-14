Namespace Njc.Model

#Region "修繕基本情報モデル"

    Public Class Szendata_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Hy_useflg As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Ky_useflg As String
        Public Property Vari_Ky_guid As String
        Public Property Vari_Szen_name As String
        Public Property Vari_Szen_ukeymd As String
        Public Property Vari_Szen_endymd As String
        Public Property Vari_Tanto_no As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Koji_yoteistartymd As String
        Public Property Vari_Koji_yoteiendymd As String
        Public Property Vari_Koji_basyo As String
        Public Property Vari_Koji_gaiyo As String
        Public Property Vari_Koji_kagi As String
        Public Property Vari_Koji_tatiai As String
        Public Property Vari_Biko As String
        Public Property Vari_Kys_futanrit As String
        Public Property Vari_Ow_futanrit As String
        Public Property Vari_Rowid As String
        Public Property Vari_History As String
        Public Property Vari_Jisya_futanrit As String
        Public Property Vari_Kys_sqflg As String
        Public Property Vari_Kys_no As String
        Public Property Vari_Kys_gtym As String
        Public Property Vari_Kys_sqsimeymd As String
        Public Property Vari_Kys_nkbn As String
        Public Property Vari_Kys_sorit As String
        Public Property Vari_Kys_yatinkozano As String
        Public Property Vari_Kys_biko As String
        Public Property Vari_Ow_sqflg As String
        Public Property Vari_Ow_no As String
        Public Property Vari_Ow_kaisyukbn As String
        Public Property Vari_Ow_kojoymd As String
        Public Property Vari_Ow_kojogtym As String
        Public Property Vari_Ow_sosaki_soruleguid As String
        Public Property Vari_Ow_sosaki_sorule_no As String
        Public Property Vari_Ow_sokozano As String
        Public Property Vari_Ow_sqymd As String
        Public Property Vari_Ow_sqgtym As String
        Public Property Vari_Ow_sqsqsimeymd As String
        Public Property Vari_Ow_sqnkbn As String
        Public Property Vari_Ow_sqyatinkozano As String
        Public Property Vari_Ow_sqsakino As String
        Public Property Vari_Ow_biko As String

    End Class

#End Region

#Region "修繕見積情報モデル"

    Public Class Szendata_szen_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Szen_mituno As String
        Public Property Vari_Mitu_title As String
        Public Property Vari_Mitu_bango As String
        Public Property Vari_Mitu_ymd As String
        Public Property Vari_Seiyaku_flg As String
        Public Property Vari_Seiyaku_ymd As String
        Public Property Vari_Futan_kbn As String
        Public Property Vari_Kys_futanrit As String
        Public Property Vari_Ow_futanrit As String
        Public Property Vari_Jisya_futanrit As String
        Public Property Vari_Zei_kbn As String
        Public Property Vari_Gokei_zeikbn As String
        Public Property Vari_Gokei_zeirit As String
        Public Property Vari_Kys_gokeizeigak As String
        Public Property Vari_Ow_gokeizeigak As String
        Public Property Vari_Jisya_gokeizeigak As String
        Public Property Vari_Hasu_futankbn As String

    End Class

#End Region

#Region "修繕見積詳細情報モデル"

    Public Class Szendata_szenmeisai_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Szen_mituno As String
        Public Property Vari_Szen_meisaino As String
        Public Property Vari_Szen_name As String
        Public Property Vari_Tekiyo As String
        Public Property Vari_Mitu_suryo As String
        Public Property Vari_Mitu_tani As String
        Public Property Vari_Mitu_tanka As String
        Public Property Vari_Mitu_zeikbn As String
        Public Property Vari_Mitu_zeigak As String
        Public Property Vari_Kys_futanrit As String
        Public Property Vari_Ow_futanrit As String
        Public Property Vari_Jisya_futanrit As String
        Public Property Vari_Szen_gyno As String
        Public Property Vari_Jikko_suryo As String
        Public Property Vari_Jikko_tani As String
        Public Property Vari_Jikko_tanka As String
        Public Property Vari_Jikko_zeikbn As String
        Public Property Vari_Jikko_zeigak As String

    End Class

#End Region

#Region "修繕クレーム関連付け情報モデル"

    Public Class Szendata_claim_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Claim_no As String
        Public Property Vari_Claimdata_sortorder As String
        Public Property Vari_Szendata_sortorder As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "修繕関連ファイル情報モデル"

    Public Class Szendata_relfile_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Szen_no As String
        Public Property Vari_File_no As String
        Public Property Vari_Fullpath As String
        Public Property Vari_Addtime As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "修繕メモ情報モデル"

    Public Class Szendata_memo_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
