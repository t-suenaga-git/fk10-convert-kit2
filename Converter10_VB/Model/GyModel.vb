Namespace Njc.Model

#Region "仲介・管理業者マスタモデル"

    Public Class Gydata_fudo_Model

        Public Property Vari_Gy_fudono As String
        Public Property Vari_Gy_fudoname As String
        Public Property Vari_Gy_fudonamesjis As String
        Public Property Vari_Gy_fudokana As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr_kenno As String
        Public Property Vari_Addr_sino As String
        Public Property Vari_Addr_cyo As String
        Public Property Vari_Addr_cyome As String
        Public Property Vari_Addr_cyomeptn As String
        Public Property Vari_Addr_banti As String
        Public Property Vari_Addr_etc As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Daihyo_yakusyoku As String
        Public Property Vari_Daihyo_name As String
        Public Property Vari_Tanto_busyoyakusyoku As String
        Public Property Vari_Tanto_name As String
        Public Property Vari_Tanto_namesjis As String
        Public Property Vari_Tanto_kana As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Url As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Yusenmail_kbn As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Sofu_hakkofurikomi As String
        Public Property Vari_Menkyo_bango As String
        Public Property Vari_Menkyo_ymd As String
        Public Property Vari_Syunin_bango As String
        Public Property Vari_Syunin_name As String
        Public Property Vari_Kanrikyokai_bango As String
        Public Property Vari_Keieikanrisi_bango As String
        Public Property Vari_Keieikanrisi_name As String
        Public Property Vari_Kanrigy_bango As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Daihyo_namesjis As String
        Public Property Vari_Syunin_namesjis As String
        Public Property Vari_Tanto_autoinputflg As String
        Public Property Vari_Keieikanrisi_namesjis As String

    End Class

#End Region

#Region "仲介・管理業者マスタ口座モデル"

    Public Class Gydata_fudokoza_Model

        Public Property Vari_Gy_fudono As String
        Public Property Vari_Gy_fudokozano As String
        Public Property Vari_Kinyu_no As String
        Public Property Vari_Kinyu_tenno As String
        Public Property Vari_Koza_syubetu As String
        Public Property Vari_Koza_bango As String
        Public Property Vari_Koza_meigi As String
        Public Property Vari_Koza_meigikana As String
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
        Public Property Vari_Yucyokoza_kigo1 As String
        Public Property Vari_Yucyokoza_kigo2 As String
        Public Property Vari_Yucyokoza_bango As String

    End Class

#End Region

#Region "仲介・管理業者マスタメモモデル"

    Public Class Gydata_fudomemo_Model

        Public Property Vari_Gy_fudono As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "保険業者マスタモデル"

    Public Class Gydata_hoken_Model

        Public Property Vari_Gy_hokenno As String
        Public Property Vari_Gy_hokenname As String
        Public Property Vari_Gy_hokennamesjis As String
        Public Property Vari_Gy_hokenkana As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr1 As String
        Public Property Vari_Addr2 As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Tanto_busyoyakusyoku As String
        Public Property Vari_Tanto_name As String
        Public Property Vari_Tanto_namesjis As String
        Public Property Vari_Tanto_kana As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Url As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Yusenmail_kbn As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String

    End Class

#End Region

#Region "保険業者マスタ口座モデル"

    Public Class Gydata_hokenkoza_Model

        Public Property Vari_Gy_hokenno As String
        Public Property Vari_Gy_hokenkozano As String
        Public Property Vari_Kinyu_no As String
        Public Property Vari_Kinyu_tenno As String
        Public Property Vari_Koza_syubetu As String
        Public Property Vari_Koza_bango As String
        Public Property Vari_Koza_meigi As String
        Public Property Vari_Koza_meigikana As String
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

    End Class

#End Region

#Region "保険業者マスタメモモデル"

    Public Class Gydata_hokenmemo_Model

        Public Property Vari_Gy_hokenno As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "家賃保証業者マスタモデル"

    Public Class Gydata_hosyo_Model

        Public Property Vari_Gy_hosyono As String
        Public Property Vari_Gy_hosyoname As String
        Public Property Vari_Gy_hosyonamesjis As String
        Public Property Vari_Gy_hosyokana As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr1 As String
        Public Property Vari_Addr2 As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Tanto_busyoyakusyoku As String
        Public Property Vari_Tanto_name As String
        Public Property Vari_Tanto_namesjis As String
        Public Property Vari_Tanto_kana As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Url As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Yusenmail_kbn As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Rendo_namekbn As String

    End Class

#End Region

#Region "家賃保証業者マスタ口座モデル"

    Public Class Gydata_hosyokoza_Model

        Public Property Vari_Gy_hosyono As String
        Public Property Vari_Gy_hosyokozano As String
        Public Property Vari_Kinyu_no As String
        Public Property Vari_Kinyu_tenno As String
        Public Property Vari_Koza_syubetu As String
        Public Property Vari_Koza_bango As String
        Public Property Vari_Koza_meigi As String
        Public Property Vari_Koza_meigikana As String
        Public Property Vari_Fkae_no As String
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

    End Class

#End Region

#Region "家賃保証業者マスタメモモデル"

    Public Class Gydata_hosyomemo_Model

        Public Property Vari_Gy_hosyono As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "修繕業者マスタモデル"

    Public Class Gydata_szen_Model

        Public Property Vari_Gy_szenno As String
        Public Property Vari_Gy_szenname As String
        Public Property Vari_Gy_szennamesjis As String
        Public Property Vari_Gy_szenkana As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr1 As String
        Public Property Vari_Addr2 As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Daihyo_yakusyoku As String
        Public Property Vari_Daihyo_name As String
        Public Property Vari_Tanto_busyoyakusyoku As String
        Public Property Vari_Tanto_name As String
        Public Property Vari_Tanto_namesjis As String
        Public Property Vari_Tanto_kana As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Url As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Yusenmail_kbn As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Daihyo_namesjis As String

    End Class

#End Region

#Region "修繕業者マスタ口座モデル"

    Public Class Gydata_szenkoza_Model

        Public Property Vari_Gy_szenno As String
        Public Property Vari_Gy_szenkozano As String
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

    End Class

#End Region

#Region "修繕業者マスタメモモデル"

    Public Class Gydata_szenmemo_Model

        Public Property Vari_Gy_szenno As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "ライフラインマスタモデル"

    Public Class Gydata_lifeline_Model

        Public Property Vari_Gy_lifelineno As String
        Public Property Vari_Gy_lifelinename As String
        Public Property Vari_Gy_lifelinenamesjis As String
        Public Property Vari_Gy_lifelinekana As String
        Public Property Vari_Eigyosyo As String
        Public Property Vari_Eigyosyosjis As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr1 As String
        Public Property Vari_Addr2 As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Lifeline_denkiflg As String
        Public Property Vari_Lifeline_josuidoflg As String
        Public Property Vari_Lifeline_gasflg As String
        Public Property Vari_Lifeline_toyuflg As String
        Public Property Vari_Lifeline_other1flg As String
        Public Property Vari_Lifeline_haisuiflg As String

    End Class

#End Region

#Region "施工会社マスタモデル"

    Public Class Gydata_seko_Model

        Public Property Vari_Gy_sekono As String
        Public Property Vari_Gy_sekoname As String
        Public Property Vari_Gy_sekonamesjis As String
        Public Property Vari_Gy_sekokana As String
        Public Property Vari_Eigyosyo As String
        Public Property Vari_Eigyosyosjis As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr1 As String
        Public Property Vari_Addr2 As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Tanto_busyoyakusyoku As String
        Public Property Vari_Tanto_name As String
        Public Property Vari_Tanto_namesjis As String
        Public Property Vari_Tanto_kana As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Url As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Yusenmail_kbn As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String

    End Class

#End Region

#Region "施設保守業者マスタモデル"

    Public Class Gydata_sisetu_Model

        Public Property Vari_Gy_sisetuno As String
        Public Property Vari_Gy_sisetuname As String
        Public Property Vari_Gy_sisetunamesjis As String
        Public Property Vari_Gy_sisetukana As String
        Public Property Vari_Eigyosyo As String
        Public Property Vari_Eigyosyosjis As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr1 As String
        Public Property Vari_Addr2 As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Tanto_busyoyakusyoku As String
        Public Property Vari_Tanto_name As String
        Public Property Vari_Tanto_namesjis As String
        Public Property Vari_Tanto_kana As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Url As String
        Public Property Vari_Yusentel_kbn As String
        Public Property Vari_Yusenmail_kbn As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String

    End Class

#End Region

End Namespace
