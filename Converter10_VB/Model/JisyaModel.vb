Namespace Njc.Model

#Region "自社情報モデル"

    Public Class Jisyadata_Model

        Public Property Vari_Jisya_no As String
        Public Property Vari_Jisya_name As String
        Public Property Vari_Jisya_namesjis As String
        Public Property Vari_Jisya_kana As String
        Public Property Vari_Jisya_name2 As String
        Public Property Vari_Jisya_name2sjis As String
        Public Property Vari_Jisya_name3 As String
        Public Property Vari_Jisya_name3sjis As String
        Public Property Vari_Daihyo_yakusyoku As String
        Public Property Vari_Daihyo_name As String
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
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Mobiletel2 As String
        Public Property Vari_Fax As String
        Public Property Vari_Mail As String
        Public Property Vari_Mobilemail As String
        Public Property Vari_Url As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Menkyo_bango As String
        Public Property Vari_Menkyo_ymd As String
        Public Property Vari_Syunin_bango As String
        Public Property Vari_Syunin_name As String
        Public Property Vari_Kanrikyokai_bango As String
        Public Property Vari_Keieikanrisi_bango As String
        Public Property Vari_Keieikanrisi_name As String
        Public Property Vari_Kanrigy_bango As String
        Public Property Vari_Kanyu_dantai1 As String
        Public Property Vari_Kanyu_dantai2 As String
        Public Property Vari_Kanyu_dantai3 As String
        Public Property Vari_Kanyu_dantai4 As String
        Public Property Vari_Kanyu_dantai5 As String
        Public Property Vari_Kanyu_dantai6 As String
        Public Property Vari_Kanyu_dantai7 As String
        Public Property Vari_Kanyu_dantai8 As String
        Public Property Vari_Kanyu_dantai9 As String
        Public Property Vari_Kanyu_dantai10 As String
        Public Property Vari_History As String
        Public Property Vari_Useflg As String
        Public Property Vari_Daihyo_namesjis As String
        Public Property Vari_Syunin_namesjis As String
        Public Property Vari_Keieikanrisi_namesjis As String

    End Class

#End Region

#Region "自社口座情報モデル"

    Public Class Jisyadata_koza_Model

        Public Property Vari_Jisya_no As String
        Public Property Vari_Jisya_kozano As String
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
        Public Property Vari_History As String
        Public Property Vari_Useflg As String

    End Class

#End Region

#Region "自社担当者情報モデル"    '20160608 自社担当者情報構築

    Public Class Profile_logonuser_Model

        Public Property Vari_Logonuser_no As String
        Public Property Vari_Logonuser_name As String
        Public Property Vari_Logonuser_namesjis As String
        Public Property Vari_Logonuser_kana As String
        Public Property Vari_Logongroup_no As String
        Public Property Vari_Tel1 As String
        Public Property Vari_Tel2 As String
        Public Property Vari_Mobiletel1 As String
        Public Property Vari_Fax1 As String
        Public Property Vari_Mailaddress1 As String
        Public Property Vari_Mobilemailaddress1 As String
        Public Property Vari_Takkenmenkyo As String
        Public Property Vari_Biko As String
        Public Property Vari_Denylogon As String
        Public Property Vari_Noassign As String
        Public Property Vari_Password As String
        Public Property Vari_Behaviourdefine_base64 As String
        Public Property Vari_Colordefine_base64 As String
        Public Property Vari_N3spreadcolordefine_base64 As String
        Public Property Vari_Menudesigndefine_base64 As String
        Public Property Vari_Smtpdefine_base64 As String
        Public Property Vari_History As String
        Public Property Vari_Sendtargetsiteid As String
        Public Property Vari_Windowsuser_name As String
        Public Property Vari_Hysearchdefine_base64 As String
        Public Property Vari_Csvsetting_base64 As String
        Public Property Vari_Dontshowfeedbackagain As String

    End Class

#End Region

#Region "自社メモ情報モデル"

    Public Class Jisyadata_memo_Model

        Public Property Vari_Jisya_no As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "振込依頼人情報モデル"

    Public Class M_fb_sgfirai_Model

        Public Property Vari_Sgfirai_no As String
        Public Property Vari_Sgfirai_name As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Jisya_kozano As String
        Public Property Vari_Sgfirainin_code As String
        Public Property Vari_Sgfirainin_kana As String
        Public Property Vari_Biko As String           '20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg (Vari_Biko_kihon → Vari_Biko)
        Public Property Vari_Sgfzenfmt_no As String
        Public Property Vari_File_sosin As String
        Public Property Vari_Crlf As String
        Public Property Vari_Changewo_flg As String
        Public Property Vari_Biko_data As String
        Public Property Vari_Keisandefault_flg As String
        Public Property Vari_Futancyousei_flg As String
        Public Property Vari_Biko_tesu As String
        Public Property Vari_History As String
        Public Property Vari_Sgfdata_createflg As String
        Public Property Vari_Useflg As String

    End Class

#End Region

#Region "振込手数料情報モデル"

    Public Class M_fb_sgfiraitesu_Model

        Public Property Vari_Sgfirai_no As String
        Public Property Vari_Rec_no As String
        Public Property Vari_From_gak As String
        Public Property Vari_To_gak As String
        Public Property Vari_Doukoudouten_densingak As String
        Public Property Vari_Doukoudouten_bunsyogak As String
        Public Property Vari_Doukoutaten_densingak As String
        Public Property Vari_Doukoutaten_bunsyogak As String
        Public Property Vari_Takou_densingak As String
        Public Property Vari_Takou_bunsyogak As String

    End Class

#End Region

#Region "口座振替情報モデル"

    Public Class M_fb_fkaejyoho_Model

        Public Property Vari_Fkae_no As String
        Public Property Vari_Fkae_name As String
        Public Property Vari_Fkae_kana As String
        Public Property Vari_Servicetype As String
        Public Property Vari_Fkae_fb_fkomiraino As String
        Public Property Vari_Fkae_fb_fkomiraikana As String
        Public Property Vari_Kamei_no As String
        Public Property Vari_Jlease_nyukinkbn As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Fkae_fb_nkinukekozano As String
        Public Property Vari_Hikiotosibi As String
        Public Property Vari_Tesu_gak As String
        Public Property Vari_Biko As String
        Public Property Vari_Fkae_fb_orgfmtno As String
        Public Property Vari_Datasort As String
        Public Property Vari_Keiyakusyano_syuturyoku As String
        Public Property Vari_Crlf As String
        Public Property Vari_Wo_mojihenkan As String
        Public Property Vari_File_sosin As String
        Public Property Vari_File_jyusin As String
        Public Property Vari_Fdsakusei As String
        Public Property Vari_Seikyu_tani As String
        Public Property Vari_Seikyu_taino As String
        Public Property Vari_Seikyu_tukitani As String
        Public Property Vari_Tesu_nyukinkanri As String
        Public Property Vari_Tesu_kurikosi As String
        Public Property Vari_Syogorule As String
        Public Property Vari_Yutyobango As String
        Public Property Vari_History As String
        Public Property Vari_Useflg As String
        Public Property Vari_Multi_flg As String            '20160519 EXEUpdateに伴う修正 口座振替情報 -add
        Public Property Vari_Yutyo_appendcode As String     '20160519 EXEUpdateに伴う修正 口座振替情報 -add
        Public Property Vari_Yutyo_usefkomcode As String    '20160519 EXEUpdateに伴う修正 口座振替情報 -add
        Public Property Vari_Resultfile_notuse As String    '20160519 EXEUpdateに伴う修正 口座振替情報 -add
        Public Property Vari_Saifkae_use As String          '20161012 革命10アップデートに伴う修正 -add

    End Class

#End Region

#Region "入出金取得情報モデル"

    Public Class M_fb_nskinsetting_Model

        Public Property Vari_Ns_no As String
        Public Property Vari_Ns_name As String
        Public Property Vari_Ns_kana As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Orgfmt_no As String
        Public Property Vari_File_jyusin As String
        Public Property Vari_Data_syubetu As String
        Public Property Vari_Crlf As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Useflg As String
        Public Property Vari_Yatin_kozano As String
        Public Property Vari_File_sosin As String
        Public Property Vari_Yokinbunkatusetting As String
        Public Property Vari_Duplicate_check As String

    End Class

#End Region

#Region "家賃入金口座情報モデル"

    Public Class M_yatinkoza_Model

        Public Property Vari_Yatin_kozano As String
        Public Property Vari_Yatin_kozaname As String
        Public Property Vari_Yatin_kozanamesjis As String
        Public Property Vari_Yatin_kozakana As String
        Public Property Vari_Sqsaki_kbn As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Jisya_kozano As String
        Public Property Vari_Ow_no As String
        Public Property Vari_Ow_kozano As String
        Public Property Vari_Spc_useflg As String
        Public Property Vari_Spc_kozasiteikbn As String
        Public Property Vari_Spc_syumokukbn As String
        Public Property Vari_Spc_ninibango As String
        Public Property Vari_Spc_kanyubango As String
        Public Property Vari_Spc_kobetuflg As String
        Public Property Vari_Spc_kobetukbn As String
        Public Property Vari_Spc_kobetuareano As String
        Public Property Vari_Spc_kobetutikuno As String
        Public Property Vari_Spc_kobetutel As String
        Public Property Vari_Spc_ansyobango As String
        Public Property Vari_Spc_servicecode As String
        Public Property Vari_History As String
        Public Property Vari_Useflg As String
        Public Property Vari_Yatin_kozabiko As String
        Public Property Vari_Headertemplate As String
        Public Property Vari_Trallertemplate As String

    End Class

#End Region

#Region "ANSERエリア情報モデル"

    Public Class M_spcarea_Model

        Public Property Vari_Spc_areano As String
        Public Property Vari_Spc_areaname As String
        Public Property Vari_Biko_spcarea As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "ANSERアクセスポイント情報モデル"

    Public Class M_spcaccesspoint_Model

        Public Property Vari_Spc_tikuno As String
        Public Property Vari_Spc_areano As String
        Public Property Vari_Spc_tikuname As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "ANSER接続情報モデル"

    Public Class M_spcsetuzoku_Model

        Public Property Vari_Spc_setuzokuno As String
        Public Property Vari_Spc_name As String
        Public Property Vari_Spc_kaisensyu As String
        Public Property Vari_Spc_devtype As String
        Public Property Vari_Spc_devname As String
        Public Property Vari_Spc_setuzoku_hoho As String
        Public Property Vari_Spc_areano As String
        Public Property Vari_Spc_tikuno As String
        Public Property Vari_Spc_tel As String
        Public Property Vari_Spc_gaisen As String
        Public Property Vari_Spc_retry_kaisu As String
        Public Property Vari_Spc_retry_kankaku As String
        Public Property Vari_Spc_servicecode As String
        Public Property Vari_Spc_biko As String
        Public Property Vari_History As String
        Public Property Vari_Spc_crlf As String

    End Class

#End Region

End Namespace
