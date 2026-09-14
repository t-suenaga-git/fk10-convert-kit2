Namespace Njc.Model

#Region "部屋基本情報モデル"

    Public Class Hydata_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_no As String
        Public Property Vari_Hy_gaibuno As String
        Public Property Vari_Hy_deleteflg As String
        Public Property Vari_Delete_guid As String
        Public Property Vari_Delete_day As String
        Public Property Vari_Delete_cnt As String
        Public Property Vari_Hy_ruinokbn As String
        Public Property Vari_Madori_cnt As String
        Public Property Vari_Madori_typekbn As String
        Public Property Vari_Madori_biko As String
        Public Property Vari_Men_senyujitu As String
        Public Property Vari_Men_senyujitutubo As String
        Public Property Vari_Men_yuka As String
        Public Property Vari_Men_yukatubo As String
        Public Property Vari_Men_senyutouki As String
        Public Property Vari_Men_senyutoukitubo As String
        Public Property Vari_Men_toki As String
        Public Property Vari_Men_tokitubo As String
        Public Property Vari_Men_balcony As String
        Public Property Vari_Men_balconytubo As String
        Public Property Vari_Men_tempo As String
        Public Property Vari_Men_tempotubo As String
        Public Property Vari_Men_jutaku As String
        Public Property Vari_Men_jutakutubo As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Delete_cause As String
        Public Property Vari_Kaiyaku_uketukekbn As String
        Public Property Vari_Kaiyaku_months As String
        Public Property Vari_Kaiyaku_days As String
        Public Property Vari_Kaiyaku_day As String
        Public Property Vari_Sort_hy_no As String       '20160519 EXEUpdateに伴う修正 部屋基本情報 -add
        Public Property Vari_Lastupdate As String       '20160519 EXEUpdateに伴う修正 部屋基本情報 -add
        Public Property Vari_Wmp_id As String           '20161012 革命10アップデートに伴う修正 -add

    End Class

#End Region

#Region "部屋詳細情報モデル"

    Public Class Hydata_detail_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Hygyomu_setflg As String
        Public Property Vari_Syozai_kaisu1 As String
        Public Property Vari_Syozai_kaisu2 As String
        Public Property Vari_Syozai_kaisu3 As String
        Public Property Vari_Syozai_tikaflg1 As String
        Public Property Vari_Syozai_tikaflg2 As String
        Public Property Vari_Syozai_tikaflg3 As String
        Public Property Vari_Mukikbn As String
        Public Property Vari_Kadoheya As String
        Public Property Vari_Saikokbn As String
        Public Property Vari_Balcony_mukikbn As String
        Public Property Vari_Nyukyo_jokyokbn As String
        Public Property Vari_Nyukyo_jokyomemo As String
        Public Property Vari_Nyukyo_syunflg As String
        Public Property Vari_Nyukyo_ym As String
        Public Property Vari_Nyukyo_syunkbn As String
        Public Property Vari_Nyukyo_joken As String
        Public Property Vari_Kakunin_ymd As String
        Public Property Vari_Bosyu_startflg As String
        Public Property Vari_Freerent_flg As String
        Public Property Vari_Freerent_month As String
        Public Property Vari_Freerent_detaill As String
        Public Property Vari_Hoken_kbn As String
        Public Property Vari_Hoken_kikan As String
        Public Property Vari_Hoken_gak As String
        Public Property Vari_Hoken_biko As String
        Public Property Vari_Torihiki_taiyokbn As String
        Public Property Vari_Torihiki_kyakutuke As String
        Public Property Vari_Torihiki_tesumoto As String
        Public Property Vari_Torihiki_tesukyaku As String
        Public Property Vari_Torihiki_futankasi As String
        Public Property Vari_Torihiki_futankari As String
        Public Property Vari_Torihiki_kyakutukecomment As String
        Public Property Vari_Torihiki_gykokokukatudokbn As String
        Public Property Vari_Moto_gy_fudono As String
        Public Property Vari_Koukoku_ryoukbn As String
        Public Property Vari_Koukoku_jogengak As String
        Public Property Vari_Koukoku_jokennaiyo As String
        Public Property Vari_Biko As String
        Public Property Vari_Addr_replaceflg As String
        Public Property Vari_Addr_replacecyome As String
        Public Property Vari_Addr_replacecyomeptn As String
        Public Property Vari_Addr_replacebanti As String
        Public Property Vari_Addr_replaceetc As String
        Public Property Vari_Hosyo_gyno As String
        Public Property Vari_Hosyo_naiyo As String
        Public Property Vari_Kyrui_nokbn As String
        Public Property Vari_Keiyaku_kikan As String
        Public Property Vari_Keiyaku_kijitu As String
        Public Property Vari_Parking_biko As String
        Public Property Vari_Parking_bikebiko As String
        Public Property Vari_Parking_cyurinbiko As String
        Public Property Vari_Shared_salespoint As String
        Public Property Vari_History As String
        Public Property Vari_Keisai_bkflg As String
        Public Property Vari_Keisai_hyflg As String
        Public Property Vari_Keisai_bantiflg As String
        Public Property Vari_Keisai_mapflg As String
        Public Property Vari_Hosyo_kbn As String
        Public Property Vari_Koukoku_jogengakkbn As String
        Public Property Vari_Koukoku_jogenrit As String
        Public Property Vari_Koukoku_jogentaxkbn As String
        Public Property Vari_Commonsalespoint_useflg As String
        Public Property Vari_Floors_flg As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Jisya_tanto As String
        Public Property Vari_Confirmky_sekininsya As String
        Public Property Vari_Confirmky_ymd As String
        Public Property Vari_Confirmky_print As String
        Public Property Vari_Confirmkai_sekininsya As String
        Public Property Vari_Confirmkai_ymd As String
        Public Property Vari_Confirmkai_print As String
        Public Property Vari_Nextnkin_kosindefault As String
        Public Property Vari_Toki_ymd As String
        Public Property Vari_Syo_kenriflg As String
        Public Property Vari_Syo_kenrikbn As String
        Public Property Vari_Other_kenriflg As String
        Public Property Vari_Btob_groupkbn As String
        Public Property Vari_Jisyaweb_osusumebk As String
        Public Property Vari_Kaiyaku_ym As String
        Public Property Vari_Taikyo_ym As String
        Public Property Vari_Svbunrui_no As String
        Public Property Vari_Krbunrui_no As String
        Public Property Vari_Siyo_mokuteki As String
        Public Property Vari_Nyukyo_sintikukbn As String
        Public Property Vari_Nyukyo_jikikbn As String
        Public Property Vari_Torihiki_jisyakbn As String
        Public Property Vari_Keiyaku_kikankbn As String

    End Class

#End Region

#Region "部屋所有者情報モデル"

    Public Class Hydata_syo_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Kn_no As String
        Public Property Vari_Sorule_guid As String
        Public Property Vari_Kasi1_ow_no As String
        Public Property Vari_Syo1_ow_no As String
        Public Property Vari_Kasi2_ow_no As String
        Public Property Vari_Syo2_ow_no As String
        Public Property Vari_Syo_startymd As String
        Public Property Vari_Syo_endymd As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋駐車場情報モデル"

    Public Class Hydata_parking_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Parking_kbn As String
        Public Property Vari_Parking_akisu As String
        Public Property Vari_Parking_status As String
        Public Property Vari_Parking_gakkbn As String
        Public Property Vari_Parking_gak As String
        Public Property Vari_Parking_gakzei As String
        Public Property Vari_History As String
        Public Property Vari_Parking_tintaisu As String

    End Class

#End Region

#Region "部屋特約情報モデル"

    Public Class Hydata_tokuyaku_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Tokuyaku_grpno As String
        Public Property Vari_Naiyo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋鍵情報モデル"

    Public Class Hydata_kagi_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Kagi_no As String
        Public Property Vari_Honsu As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Hokan As String
        Public Property Vari_Gyshare As String

    End Class

#End Region

#Region "部屋面積情報モデル"

    Public Class Hydata_othermenseki_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Mensekitype As String
        Public Property Vari_No As String
        Public Property Vari_Mensekikbn As String
        Public Property Vari_Name As String
        Public Property Vari_Meter As String
        Public Property Vari_Tubo As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋間取内訳情報モデル"

    Public Class Hydata_madoriutiwake_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_No As String
        Public Property Vari_Hykbn As String
        Public Property Vari_Jo As String
        Public Property Vari_History As String
        Public Property Vari_Syozaikai As String

    End Class

#End Region

#Region "部屋設備情報モデル"

    Public Class Hydata_setubilst_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Komok_guid As String
        Public Property Vari_Override_kbn As String
        Public Property Vari_Komok_name As String
        Public Property Vari_Disp1name As String
        Public Property Vari_Disp2name As String
        Public Property Vari_Disp3name As String
        Public Property Vari_Disp1iconguid As String
        Public Property Vari_Disp2iconguid As String
        Public Property Vari_Disp3iconguid As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋入金項目情報モデル"

    Public Class Hydata_nkin_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Tuki_kbn As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkin_recno As String
        Public Property Vari_Nkin_kbn As String
        Public Property Vari_Sq_gak As String
        Public Property Vari_Sq_zeikbn As String
        Public Property Vari_Calc_kbn As String
        Public Property Vari_Calc_nkinno As String
        Public Property Vari_Calc_monthcnt As String
        Public Property Vari_Sqsaki_no As String
        Public Property Vari_Sq_mmkbn As String
        Public Property Vari_Sqstart_ymd As String
        Public Property Vari_Sq_ptn As String
        Public Property Vari_Sq_interval As String
        Public Property Vari_Sq_nen As String
        Public Property Vari_Sq_tuki As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋変動費各戸メーター情報モデル"

    Public Class Hydata_hendo_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Hyhendo_guid As String
        Public Property Vari_Rec_no As String
        Public Property Vari_Hendo_kbn As String
        Public Property Vari_Meter_name As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Biko As String
        Public Property Vari_Hendorule_no As String
        Public Property Vari_Hendorule_biko As String
        Public Property Vari_History As String
        Public Property Vari_Sq_mmkbn As String
        Public Property Vari_Useflg As String

    End Class

#End Region

#Region "部屋修繕維持管理連絡先情報モデル"

    Public Class Hydata_szeniji_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Syuzenijikanri_no As String
        Public Property Vari_Syuzenijikanri_kasyo As String
        Public Property Vari_Syuzenijikanri_taisyokbn As String
        Public Property Vari_Syuzenijikanri_gyno As String
        Public Property Vari_Syuzenijikanri_simei As String
        Public Property Vari_Syuzenijikanri_address As String
        Public Property Vari_Syuzenijikanri_tel As String
        Public Property Vari_Syuzenijikanri_jisyano As String
        Public Property Vari_Syuzenijikanri_kasino As String
        Public Property Vari_Syuzenijikanri_syono As String
        Public Property Vari_Syuzenijikanri_simeiu As String

    End Class

#End Region

#Region "部屋メモ情報モデル"

    Public Class Hydata_memo_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

    '汎用ツールのみ

#Region "部屋共通セールスポイント情報モデル"

    Public Class Hydata_commonsalespointparts_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Parts_no As String
        Public Property Vari_Salespoint_parts As String

    End Class

#End Region

#Region "部屋契約解約確認事項情報モデル"

    Public Class Hydata_confirm_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Confirm_kbn As String
        Public Property Vari_Confirm_no As String
        Public Property Vari_Naiyo As String

    End Class

#End Region

#Region "部屋権利情報モデル"

    Public Class Hydata_kenri_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Kenri_no As String
        Public Property Vari_Other_kenrirui As String
        Public Property Vari_Other_kenribiko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋参照ファイル情報モデル"

    Public Class Hydata_relfile_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_File_no As String
        Public Property Vari_Fullpath As String
        Public Property Vari_Addtime As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋原状回復目安単価情報モデル"

    Public Class Hydata_szen_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Szen_name As String
        Public Property Vari_Ryo As String
        Public Property Vari_Unit_name As String
        Public Property Vari_Tanka As String

    End Class

#End Region

    'プログラム内で自動実行

#Region "部屋外部No管理情報モデル"

    Public Class Hydata_gaibuno_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Current_no As String
        Public Property Vari_History As String

    End Class

#End Region

    '他の箇所で移行

#Region "部屋画像情報モデル"

    'Public Class Hydata_gazo_Model

    '    Public Property Vari_Hy_guid As String
    '    Public Property Vari_Hy_gazono As String
    '    Public Property Vari_Hy_gazocomment As String
    '    Public Property Vari_Hy_gazolastupdate As String
    '    Public Property Vari_Hy_gazofilesizekbyte As String
    '    Public Property Vari_Hy_gazowidth As String
    '    Public Property Vari_Hy_gazoheight As String
    '    Public Property Vari_Hy_gazothumbimage As String
    '    Public Property Vari_History As String

    'End Class

#End Region

    '移行対象外

#Region "部屋業務期間情報モデル"

    Public Class Hydata_gyomukikan_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Kn_no As String
        Public Property Vari_Startymd As String
        Public Property Vari_Endymd As String
        Public Property Vari_History As String

    End Class

#End Region

    '不明テーブル
    '20160720 連動情報構築 -del sta
    '#Region "部屋BtoBグループ情報モデル"

    '    Public Class Hydata_btobgroup_Model

    '        Public Property Vari_Hy_guid As String
    '        Public Property Vari_Group_no As String
    '        Public Property Vari_Display_kbn As String
    '        Public Property Vari_History As String

    '    End Class

    '#End Region
    '20160720 連動情報構築 -del end
#Region "部屋KeyValue情報モデル"

    Public Class Hydata_keyvalue_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Data_kbn As String
        Public Property Vari_Item_kbn As String
        Public Property Vari_Value As String

    End Class

#End Region
    '20160720 連動情報構築 -del sta
    '#Region "部屋広告補足情報(作り直し中)モデル"

    '    Public Class Hydata_kokoku_Model

    '        Public Property Vari_Hy_guid As String
    '        Public Property Vari_Site_no As String
    '        Public Property Vari_Item As String

    '    End Class

    '#End Region
    '20160720 連動情報構築 -del end
#Region "部屋hydata_kys_atenaprintモデル"

    Public Class Hydata_kys_atenaprint_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String

    End Class

#End Region

#Region "部屋ポータル最終送信履歴情報モデル"

    Public Class Hydata_lastsend_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Site_no As String
        Public Property Vari_Setting_guid As String
        Public Property Vari_Lastsendtime As String

    End Class

#End Region

#Region "部屋セールスポイント情報モデル"

    Public Class Hydata_salespoint_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Site_no As String
        Public Property Vari_Salespoint As String

    End Class

#End Region
    '20160720 連動情報構築 -del sta
    '#Region "部屋連動サイト毎地図表示情報モデル"

    '    Public Class Hydata_mapdisp_Model

    '        Public Property Vari_Hy_guid As String
    '        Public Property Vari_Site_no As String
    '        Public Property Vari_Site_jisyano As String
    '        Public Property Vari_Map_dispflg As String
    '        Public Property Vari_History As String

    '    End Class

    '#End Region
    '20160720 連動情報構築 -del end
#Region "部屋hydata_panorama情報モデル"

    Public Class Hydata_panorama_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_No As String
        Public Property Vari_Gazokbn As String
        Public Property Vari_Title As String
        Public Property Vari_Lenskbn As String
        Public Property Vari_Uri As String
        Public Property Vari_Gazolastupdate As String
        Public Property Vari_Gazofilesizekbyte As String
        Public Property Vari_Gazowidth As String
        Public Property Vari_Gazoheight As String
        Public Property Vari_Gazothumbimage As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋支店情報モデル"

    Public Class Hydata_siten_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Rec_no As String
        Public Property Vari_Siten_no As String
        Public Property Vari_History As String

    End Class

#End Region
    '20160720 連動情報構築 -del sta
    '#Region "部屋hydata_sosin情報モデル"

    '    Public Class Hydata_sosin_Model

    '        Public Property Vari_Hy_guid As String
    '        Public Property Vari_Site_no As String
    '        Public Property Vari_Site_jisyano As String
    '        Public Property Vari_Select_no As String
    '        Public Property Vari_Komok_guid As String
    '        Public Property Vari_History As String

    '    End Class

    '#End Region
    '20160720 連動情報構築 -del end
#Region "部屋hydata_youtube情報モデル"

    Public Class Hydata_youtube_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_No As String
        Public Property Vari_Uri As String
        Public Property Vari_Comment As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "部屋親メーター関連情報モデル"

    Public Class Hydata_hendoparent_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Bkhendo_guid As String
        Public Property Vari_Kanren_flg As String
        Public Property Vari_Childmeter_name As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
