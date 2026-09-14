Namespace Njc.Model

#Region "契約基本情報モデル"

    Public Class Kydata_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Ky_no As String
        Public Property Vari_Ky_deleteflg As String
        Public Property Vari_Delete_guid As String
        Public Property Vari_Delete_day As String
        Public Property Vari_Delete_cnt As String
        Public Property Vari_Syokai_kyymd As String
        Public Property Vari_Status As String
        Public Property Vari_Cancelriyu As String
        Public Property Vari_Status_ymd As String
        Public Property Vari_Cyukai_gy_fudono As String
        Public Property Vari_Syunin_logonuser_no As String
        Public Property Vari_Tetuke_gak1 As String
        Public Property Vari_Tetuke_ymd1 As String
        Public Property Vari_Tetuke_biko1 As String
        Public Property Vari_Kaiyaku_flg As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Movefrom_kyguid As String
        Public Property Vari_Moveto_kyguid As String
        Public Property Vari_Svbunrui_no As String
        Public Property Vari_Krbunrui_no As String
        Public Property Vari_Kaiyaku_uketukekbn As String
        Public Property Vari_Kaiyaku_months As String
        Public Property Vari_Kaiyaku_days As String
        Public Property Vari_Kaiyaku_day As String
        Public Property Vari_Cyukai_tantoname As String
        Public Property Vari_Cyukai_tantonamesjis As String

    End Class

#End Region

#Region "契約履歴情報モデル"

    Public Class Kydata_kihon_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Ko_no As String
        Public Property Vari_Henko_no As String
        Public Property Vari_Sqdata_guid As String
        Public Property Vari_Ky_bango As String
        Public Property Vari_Ky_ymd As String
        Public Property Vari_Kystart_ymd As String
        Public Property Vari_Kyend_ymd As String
        Public Property Vari_Henko_ymd As String
        Public Property Vari_Tuti_ymd As String
        Public Property Vari_Print_ymd As String
        Public Property Vari_Kanri_gy_fudono As String
        Public Property Vari_Gy_hosyono As String
        Public Property Vari_Hosyo_naiyo As String
        Public Property Vari_Kokyaku_bango As String
        Public Property Vari_Kyrui_no As String
        Public Property Vari_Siyo_mokuteki As String
        Public Property Vari_Kosin_umu As String
        Public Property Vari_Yatin_kbn As String
        Public Property Vari_Fkae_startym As String
        Public Property Vari_Fkae_willstartflg As String
        Public Property Vari_Yatin_kozakbn As String
        Public Property Vari_Yatin_kozano As String
        Public Property Vari_Maitukiyatin_kozano As String
        Public Property Vari_Yokugetu_uketoriflg As String
        Public Property Vari_Biko As String
        Public Property Vari_Tougetu_sagakuflg As String
        Public Property Vari_Nextky_startymd As String
        Public Property Vari_Nextky_endymd As String
        Public Property Vari_Kosin_hiwariflg As String
        Public Property Vari_Sokojorule_kbn As String
        Public Property Vari_Soyotei_ymdflg As String
        Public Property Vari_Soyotei_ymd As String
        Public Property Vari_Kanritesu_flg As String
        Public Property Vari_Kanritesu_gak As String
        Public Property Vari_Nextnkinset_kbn As String
        Public Property Vari_Sqdata_startymd As String
        Public Property Vari_Sqdata_endymd As String
        Public Property Vari_History As String
        Public Property Vari_Biko2 As String
        Public Property Vari_Hoken_biko As String
        Public Property Vari_Nkinsime_ymd As String
        Public Property Vari_Yatin_jisansaki As String
        Public Property Vari_Hoken_kikan As String
        Public Property Vari_Hoken_gak As String
        Public Property Vari_Confirmky_sekininsya As String
        Public Property Vari_Confirmky_ymd As String
        Public Property Vari_Confirmky_print As String
        Public Property Vari_Confirmkai_sekininsya As String
        Public Property Vari_Confirmkai_ymd As String
        Public Property Vari_Confirmkai_print As String
        Public Property Vari_Hikiuke_name As String
        Public Property Vari_Hikiuke_addr As String
        Public Property Vari_Hikiuke_tel As String
        Public Property Vari_Kohokennkin_ymd As String
        Public Property Vari_Kokanryo_ymd As String
        Public Property Vari_Kokanryotuti_ymd As String
        Public Property Vari_Kotuti_ymd As String
        Public Property Vari_Kosaisoku_ymd As String
        Public Property Vari_Kosyoruiuke_ymd As String
        Public Property Vari_Kokairenraku_umu As String
        Public Property Vari_Kokairenraku_ymd As String
        Public Property Vari_Kokairenraku_logonuser_no As String
        Public Property Vari_Kokairenraku_biko As String
        Public Property Vari_Maitukisq_umu As String
        Public Property Vari_Hikiuke_namesjis As String
        Public Property Vari_Kofkae_ymd As String
        Public Property Vari_Yokugetu_uketorimonth As String
        Public Property Vari_Ky_logonuser_no As String
        Public Property Vari_Sq_logonuser_no As String
        Public Property Vari_Torihiki_tesumoto As String            '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        Public Property Vari_Torihiki_tesukyaku As String           '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        Public Property Vari_Torihiki_tesugak As String             '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        Public Property Vari_Torihiki_tesuzeikbn As String          '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        Public Property Vari_Torihiki_tesuzeigak As String          '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add

    End Class

#End Region

#Region "契約車情報モデル"

    Public Class Kydata_car_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Car_cnt As String
        Public Property Vari_Carmaker As String
        Public Property Vari_Carname As String
        Public Property Vari_Carcolor As String
        Public Property Vari_Carnumber As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Car_kukaku As String

    End Class

#End Region

#Region "契約契約者情報モデル"

    Public Class Kydata_kys_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Kys_cnt As String
        Public Property Vari_Kys_no As String
        Public Property Vari_Nyukyo_flg As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "契約入居者情報モデル"

    Public Class Kydata_nyukyo_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Nyukyosya_cnt As String
        Public Property Vari_Kys_no As String
        Public Property Vari_Name As String
        Public Property Vari_Nameu As String
        Public Property Vari_Kana As String
        Public Property Vari_Gender As String
        Public Property Vari_Aidagara As String
        Public Property Vari_Birthday As String
        Public Property Vari_Kinmusaki As String
        Public Property Vari_Tel As String
        Public Property Vari_Mobiletel As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "契約保証人情報モデル"

    Public Class Kydata_hosyonin_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Hosyonin_cnt As String
        Public Property Vari_Kys_no As String
        Public Property Vari_Hosyonin_no As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "契約保険情報モデル"

    Public Class Kydata_hoken_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Hoken_no As String
        Public Property Vari_Gy_hokenno As String
        Public Property Vari_Ky_ymd As String
        Public Property Vari_Kystart_ymd As String
        Public Property Vari_Kyend_ymd As String
        Public Property Vari_Hoken_gak As String
        Public Property Vari_Mankituti_flg As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Syoken_bango As String

    End Class

#End Region

#Region "契約特約事項情報モデル"

    Public Class Kydata_tokuyaku_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Tokuyaku_grpno As String
        Public Property Vari_Naiyo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "契約メモ情報モデル"

    Public Class Kydata_memo_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "契約入金項目情報モデル"

    Public Class Kydata_nkin_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Tuki_kbn As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkin_recno As String
        Public Property Vari_Nkin_sortorder As String
        Public Property Vari_Nkin_kbn As String
        Public Property Vari_Sq_gak As String
        Public Property Vari_Sq_zeikbn As String
        Public Property Vari_Sq_zeigak As String
        Public Property Vari_Calc_kbn As String
        Public Property Vari_Calc_nkinno As String
        Public Property Vari_Calc_monthcnt As String
        Public Property Vari_Sqsaki_no As String
        Public Property Vari_Nkbn_yotei As String
        Public Property Vari_Sq_mmkbn As String
        Public Property Vari_Frstart_ymd As String
        Public Property Vari_Frend_ymd As String
        Public Property Vari_Frsq_gak As String
        Public Property Vari_Sqstart_ymd As String
        Public Property Vari_Sq_ptn As String
        Public Property Vari_Sq_interval As String
        Public Property Vari_Sq_nen As String
        Public Property Vari_Sq_tuki As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Biko As String
        Public Property Vari_Nkin_guid As String
        Public Property Vari_History As String
        Public Property Vari_Fr_kbn As String

    End Class

#End Region

#Region "契約次回入金項目情報モデル"

    Public Class Kydata_nkin_nx_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Tuki_kbn As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkin_recno As String
        Public Property Vari_Nkin_sortorder As String
        Public Property Vari_Nkin_kbn As String
        Public Property Vari_Sq_gak As String
        Public Property Vari_Sq_zeikbn As String
        Public Property Vari_Sq_zeigak As String
        Public Property Vari_Calc_kbn As String
        Public Property Vari_Calc_nkinno As String
        Public Property Vari_Calc_monthcnt As String
        Public Property Vari_Sqsaki_no As String
        Public Property Vari_Nkbn_yotei As String
        Public Property Vari_Sq_mmkbn As String
        Public Property Vari_Sqstart_ymd As String
        Public Property Vari_Sq_ptn As String
        Public Property Vari_Sq_interval As String
        Public Property Vari_Sq_nen As String
        Public Property Vari_Sq_tuki As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "契約変動費各戸メーター情報モデル"

    Public Class Kydata_hendo_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Hyhendo_guid As String
        Public Property Vari_Hendo_sortorder As String
        Public Property Vari_Useflg As String
        Public Property Vari_Hendo_kbn As String
        Public Property Vari_Meter_name As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Hendorule_no As String
        Public Property Vari_Hendorule_biko As String
        Public Property Vari_Sqsaki_no As String
        Public Property Vari_Nkbn_yotei As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Sq_mmkbn As String

    End Class

#End Region

#Region "契約控除ルール情報モデル"

    Public Class Kydata_kojorule_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Taisyokbn As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkin_sortorder As String
        Public Property Vari_Sosai_flg As String
        Public Property Vari_Calc_kbn As String
        Public Property Vari_Kojo_gak As String
        Public Property Vari_Kojo_gakzeikbn As String
        Public Property Vari_Kojonkin_no As String
        Public Property Vari_Kojo_rit As String
        Public Property Vari_Kojo_ritzeikbn As String
        Public Property Vari_Kojo_ritutizei As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Tateazu_flg As String
        Public Property Vari_Kojo_zeigak As String
        Public Property Vari_Sotaisyo_flg As String
        Public Property Vari_Sotaisyonkin_no As String

    End Class

#End Region

#Region "契約送金ルール情報モデル"

    Public Class Kydata_sorule_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Taisyokbn As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkin_sortorder As String
        Public Property Vari_Sokin_rit As String
        Public Property Vari_Kanrigak_rit As String
        Public Property Vari_Hosyo_flg As String

    End Class

#End Region

#Region "契約解約情報モデル"

    Public Class Kydata_kai_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Kai_ymd As String
        Public Property Vari_Seisan_ymd As String
        Public Property Vari_Seisan_completeymd As String
        Public Property Vari_Uketuke_ymd As String
        Public Property Vari_Riyu As String
        Public Property Vari_Tatiai_ymd As String
        Public Property Vari_Uketuke_logonuser_no As String
        Public Property Vari_Tatiai_logonuser_no As String
        Public Property Vari_Seisan_logonuser_no As String
        Public Property Vari_Szen_logonuser_no As String
        Public Property Vari_Szen_jisyano As String
        Public Property Vari_Szen_endymd As String
        Public Property Vari_Szen_kojiyoteistartymd As String
        Public Property Vari_Szen_kojiyoteiendymd As String
        Public Property Vari_Szen_kojibasyo As String
        Public Property Vari_Szen_kojigaiyo As String
        Public Property Vari_Next_name As String
        Public Property Vari_Next_namesjis As String
        Public Property Vari_Next_postcode As String
        Public Property Vari_Next_addr1 As String
        Public Property Vari_Next_addr2 As String
        Public Property Vari_Next_tel1 As String
        Public Property Vari_Yatin_kozano As String
        Public Property Vari_Biko As String
        Public Property Vari_Biko_tatiai As String
        Public Property Vari_Seisan_henkinymd As String
        Public Property Vari_Seisan_sqymd As String
        Public Property Vari_Kanritesu_flg As String
        Public Property Vari_Kanritesu_gak As String
        Public Property Vari_Rowid As String
        Public Property Vari_History As String
        Public Property Vari_So_yoteiymd As String
        Public Property Vari_Tatiai As String
        Public Property Vari_Tatiai_yoteiymd As String
        Public Property Vari_Bosyu_jokenymd As String
        Public Property Vari_Ow_logonuser_no As String
        Public Property Vari_Bosyujoken As String
        Public Property Vari_Confirmkai_sekininsya As String
        Public Property Vari_Confirmkai_ymd As String
        Public Property Vari_Confirmkai_print As String
        'Public Property Vari_Szen_jisyafutanumuflg As String
        'Public Property Vari_Szen_jisyafutangak As String
        'Public Property Vari_Szen_jisyafutanzeikbn As String
        'Public Property Vari_Szen_jisyafutanzeigak As String
        'Public Property Vari_Szen_bikojisyafutan As String
        'Public Property Vari_Szen_sonotafutanumuflg As String
        'Public Property Vari_Szen_sonotafutangak As String
        'Public Property Vari_Szen_sonotafutanzeikbn As String
        'Public Property Vari_Szen_sonotafutanzeigak As String
        'Public Property Vari_Szen_sonotafutanname As String
        'Public Property Vari_Szen_bikosonotafutan As String
        Public Property Vari_Tatiai_yoteitime As String
        Public Property Vari_Next_keisyo As String
        Public Property Vari_Nkin_kaiyakutukinoprintflg As String      '20160519 EXEUpdateに伴う修正 契約解約情報 -add
        Public Property Vari_Szen_kysno As String                       '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_kysnkbn As String                     '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_kyssorit As String                    '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_kyssogak As String                    '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owno As String                        '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owkaisyukbn As String                 '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owsosakisoruleguid As String          '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owsosakisoruleno As String            '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owsokozano As String                  '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owsqsimeymd As String                 '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_ownkbn As String                      '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owyatinkozano As String               '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owsqsakino As String                  '20160620 EXEUpdateに伴う修正2 -add
        Public Property Vari_Szen_owkojoymd As String                   '20160829 革命10バージョンアップに伴う修正 -add
        Public Property Vari_Szen_owsqymd As String                     '20160829 革命10バージョンアップに伴う修正 -add

    End Class

#End Region


#Region "契約修繕見積情報モデル"

    Public Class Kydata_kaiszen_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
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

#Region "契約修繕見積詳細情報モデル"

    Public Class Kydata_kaiszenmeisai_Model

        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
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

End Namespace
