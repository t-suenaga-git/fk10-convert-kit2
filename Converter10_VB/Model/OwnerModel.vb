Namespace Njc.Model

#Region "家主情報モデル"

    Public Class Owdata_Model

        Public Property Vari_Ow_no As String
        Public Property Vari_Kojinhojin_flg As String
        Public Property Vari_Ow_name As String
        Public Property Vari_Ow_namesjis As String
        Public Property Vari_Ow_kana As String
        Public Property Vari_Keisyo As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr_kenno As String
        Public Property Vari_Addr_sino As String
        Public Property Vari_Addr_cyo As String
        Public Property Vari_Addr_cyome As String
        Public Property Vari_Addr_cyomeptn As String
        Public Property Vari_Addr_banti As String
        Public Property Vari_Addr_etc As String
        Public Property Vari_Toukiaddr1 As String
        Public Property Vari_Toukiaddr2 As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Yusenmail_kbn As String
        Public Property Vari_Birthday As String
        Public Property Vari_Nensyu As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Gender As String
        Public Property Vari_Honseki As String
        Public Property Vari_Kinmu_name As String
        Public Property Vari_Kinmu_namesjis As String
        Public Property Vari_Kinmu_kana As String
        Public Property Vari_Kinmu_postcode As String
        Public Property Vari_Kinmu_addr1 As String
        Public Property Vari_Kinmu_addr2 As String
        Public Property Vari_Kinmu_tel1 As String
        Public Property Vari_Kinmu_tel2 As String
        Public Property Vari_Kinmu_fax As String
        Public Property Vari_Kinmu_gyosyu As String
        Public Property Vari_Kinmu_busyo As String
        Public Property Vari_Kinmu_nyuryokuym As String
        Public Property Vari_Kinmu_nyusyaym As String
        Public Property Vari_Url As String
        Public Property Vari_Gyosyu As String
        Public Property Vari_Daihyo_name As String
        Public Property Vari_Daihyo_namesjis As String
        Public Property Vari_Daihyo_kana As String
        Public Property Vari_Daihyo_yakusyoku As String
        Public Property Vari_Daihyo_atenaflg As String
        Public Property Vari_Tanto_name As String
        Public Property Vari_Tanto_namesjis As String
        Public Property Vari_Tanto_kana As String
        Public Property Vari_Tanto_busyo As String
        Public Property Vari_Tanto_atenaflg As String
        Public Property Vari_Sihonkin As String
        Public Property Vari_Jugyosu As String
        Public Property Vari_Nyuryoku_ym As String
        Public Property Vari_Torihikisaki As String
        Public Property Vari_Renraku_name As String
        Public Property Vari_Renraku_namesjis As String
        Public Property Vari_Renraku_kana As String
        Public Property Vari_Renraku_keisyo As String
        Public Property Vari_Renraku_postcode As String
        Public Property Vari_Renraku_addr1 As String
        Public Property Vari_Renraku_addr2 As String
        Public Property Vari_Renraku_tel1 As String
        Public Property Vari_Renraku_tel2 As String
        Public Property Vari_Renraku_fax As String
        Public Property Vari_Renraku_mobiletel1 As String
        Public Property Vari_Renraku_mobiletel2 As String
        Public Property Vari_Renraku_yusentelkbn As String
        Public Property Vari_Renraku_aidagara As String
        Public Property Vari_Biko_renraku As String
        Public Property Vari_Sofu_kbn As String
        Public Property Vari_Sofu_name As String
        Public Property Vari_Sofu_namesjis As String
        Public Property Vari_Sofu_kana As String
        Public Property Vari_Sofu_keisyo As String
        Public Property Vari_Sofu_postcode As String
        Public Property Vari_Sofu_addr1 As String
        Public Property Vari_Sofu_addr2 As String
        Public Property Vari_Sofu_tel1 As String
        Public Property Vari_Sofu_tel2 As String
        Public Property Vari_Sofu_fax As String
        Public Property Vari_Biko_sofu As String
        Public Property Vari_Event_tantono As String
        Public Property Vari_Useflg As String
        Public Property Vari_Findkeyword As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Kaikei_jouhou As String
        Public Property Vari_Addr1 As String            '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add
        Public Property Vari_Addr2 As String            '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add

    End Class

#End Region

#Region "家主口座情報モデル"

    Public Class Owdata_koza_Model

        Public Property Vari_Ow_no As String
        Public Property Vari_Ow_kozano As String
        Public Property Vari_Kinyu_no As String
        Public Property Vari_Kinyu_tenno As String
        Public Property Vari_Koza_syubetu As String
        Public Property Vari_Koza_bango As String
        Public Property Vari_Koza_meigi As String
        Public Property Vari_Koza_meigikana As String
        Public Property Vari_Yucyokoza_kigo1 As String
        Public Property Vari_Yucyokoza_kigo2 As String
        Public Property Vari_Yucyokoza_bango As String
        Public Property Vari_Biko_koza As String
        Public Property Vari_Biko_furikomi As String
        Public Property Vari_Sgfirai_kbn As String
        Public Property Vari_Sgfirai_no As String
        Public Property Vari_Sgfirai_tesufutankbn As String
        Public Property Vari_Sgfirai_tesukeisankbn As String
        Public Property Vari_Sgfirai_tesukotei1gak As String
        Public Property Vari_Sgfirai_tesukotei2gak As String
        Public Property Vari_Biko_sgfirai As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Koza_printkbn As String

    End Class

#End Region

#Region "家主イベント情報モデル"

    Public Class Owdata_event_Model

        Public Property Vari_Ow_no As String
        Public Property Vari_Event_kbn As String
        Public Property Vari_Event_cnt As String
        Public Property Vari_Event_ymd As String
        Public Property Vari_Event_data As String

    End Class

#End Region

#Region "家主メモ情報モデル"

    Public Class Owdata_memo_Model

        Public Property Vari_Ow_no As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
