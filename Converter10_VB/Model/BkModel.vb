Namespace Njc.Model

#Region "物件基本情報モデル"

    Public Class Bkdata_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Bk_no As String
        Public Property Vari_Bk_deleteflg As String
        Public Property Vari_Delete_guid As String
        Public Property Vari_Delete_day As String
        Public Property Vari_Delete_cnt As String
        Public Property Vari_Bk_name As String
        Public Property Vari_Bk_deletename As String
        Public Property Vari_Bk_kana As String
        Public Property Vari_Tatemono_sikibetu As String
        Public Property Vari_Post_code As String
        Public Property Vari_Addr_kenno As String
        Public Property Vari_Addr_sino As String
        Public Property Vari_Addr_cyo As String
        Public Property Vari_Addr_cyome As String
        Public Property Vari_Addr_cyomeptn As String
        Public Property Vari_Addr_banti As String
        Public Property Vari_Addr_etc As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_Bk_namesjis As String
        Public Property Vari_Bk_gaibuno As String
        Public Property Vari_Delete_cause As String
        Public Property Vari_Lastupdate As String       '20160519 EXEUpdateに伴う修正 物件基本情報 -add

    End Class

#End Region

#Region "物件外部No情報"

    Public Class Bkdata_gaibuno_Model

        Public Property Vari_Current_no As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "物件詳細情報モデル"

    Public Class Bkdata_detail_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Bk_ruinokbn As String
        Public Property Vari_Gps_wgsido As String
        Public Property Vari_Gps_wgskeido As String
        Public Property Vari_Gps_tokyoido As String
        Public Property Vari_Gps_tokyokeido As String
        Public Property Vari_Bk_dentetuflg As String
        Public Property Vari_Syunko_ymd As String
        Public Property Vari_Kaidate As String
        Public Property Vari_Tika As String
        Public Property Vari_Elevator_flg As String
        Public Property Vari_Elevator_number As String
        Public Property Vari_Rooftop_flg As String
        Public Property Vari_Kozo_nokbn As String
        Public Property Vari_Kozo_yanekbn As String
        Public Property Vari_Moto_gy_fudono As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Kosu_total As String
        Public Property Vari_Men_nobeyuka As String
        Public Property Vari_Men_nobeyukatubo As String
        Public Property Vari_Men_sikiti As String
        Public Property Vari_Men_sikititubo As String
        Public Property Vari_Men_parking As String
        Public Property Vari_Men_parkingtubo As String
        Public Property Vari_Men_yukatoki As String
        Public Property Vari_Men_yukatokitubo As String
        Public Property Vari_Men_sikititoki As String
        Public Property Vari_Men_sikititokitubo As String
        Public Property Vari_Default_sqtuki As String
        Public Property Vari_Keiyakuyou_sime As String
        Public Property Vari_Yatin_jisansaki As String
        Public Property Vari_Yatin_kozano As String
        Public Property Vari_Yatin_kykozaflg As String
        Public Property Vari_Yatin_kykozano As String
        Public Property Vari_Kozo_taikakbn As String
        Public Property Vari_Gy_sekono As String
        Public Property Vari_Gy_hosyuno As String
        Public Property Vari_Kanri_hosiki As String
        Public Property Vari_Kanri_gyname As String
        Public Property Vari_Kanri_gytanto As String
        Public Property Vari_Kanri_gytel As String
        Public Property Vari_Kanrinin_gyomukeitai As String
        Public Property Vari_Kanrinin_tel As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Kotu_sonota1 As String
        Public Property Vari_Kotu_sonota1kyori As String
        Public Property Vari_Kotu_sonota2 As String
        Public Property Vari_Kotu_sonota2kyori As String
        Public Property Vari_Denki_gy_lifeno As String
        Public Property Vari_Water_gy_lifeno As String
        Public Property Vari_Gas_gy_lifeno As String
        Public Property Vari_Haisui_gy_lifeno As String
        Public Property Vari_Toyu_gy_lifeno As String
        Public Property Vari_Lifeline1_gy_lineno As String
        Public Property Vari_Lifeline2_gy_lineno As String
        Public Property Vari_Lifeline3_gy_lineno As String
        Public Property Vari_Kensingyomu_umu As String
        Public Property Vari_Kensinorder_kbn As String
        Public Property Vari_Parenthendo_useflg As String
        Public Property Vari_Parking_kanriflg As String
        Public Property Vari_Parking_car As String
        Public Property Vari_Parking_bike As String
        Public Property Vari_Parking_bicycle As String
        Public Property Vari_Parking_bicyclefreeflg As String
        Public Property Vari_Kubunsyo As String
        'Public Property Vari_Soymd_basis As String  '20160519 EXEUpdateに伴う修正 物件詳細情報 -del
        Public Property Vari_Isiwata_kirokukbn As String
        Public Property Vari_Isiwata_syokai1flg As String
        Public Property Vari_Isiwata_syokai2flg As String
        Public Property Vari_Isiwata_syokai3flg As String
        Public Property Vari_Isiwata_syokai4flg As String
        Public Property Vari_Isiwata_gy_sekono As String
        Public Property Vari_Isiwata_cyosaymd As String
        Public Property Vari_Isiwata_cyosakikankbn As String
        Public Property Vari_Isiwata_cyosahani As String
        Public Property Vari_Ishiwata_useflg As String
        Public Property Vari_Isiwata_usearea As String
        Public Property Vari_Isiwata_biko As String
        Public Property Vari_Taisin_sindanflg As String
        Public Property Vari_Taisin_syokai1flg As String
        Public Property Vari_Taisin_syokai2flg As String
        Public Property Vari_Taisin_syokai3flg As String
        Public Property Vari_Taisin_syorui1flg As String
        Public Property Vari_Taisin_syorui2flg As String
        Public Property Vari_Taisin_syorui3flg As String
        Public Property Vari_Taisin_biko As String
        Public Property Vari_Horei_dosyatiiki As String
        Public Property Vari_Horei_ruikbn As String
        Public Property Vari_Horei_naiyo As String
        Public Property Vari_Sikiti_riyoruikbn As String
        Public Property Vari_Sikiti_kystartymd As String
        Public Property Vari_Sikiti_kyendymd As String
        Public Property Vari_Sikiti_biko As String
        Public Property Vari_Nyukyoritu_startymd As String
        Public Property Vari_History As String
        Public Property Vari_Hyothergyfudo_flg As String
        Public Property Vari_Gomi_hosoku As String
        Public Property Vari_Toki_ymd As String
        Public Property Vari_Syo_kenriflg As String
        Public Property Vari_Syo_kenrikbn As String
        Public Property Vari_Other_kenriflg As String
        Public Property Vari_Men_kentiku As String
        Public Property Vari_Men_kentikutubo As String
        Public Property Vari_Men_kentikutoki As String
        Public Property Vari_Men_kentikutokitubo As String
        Public Property Vari_Bk_logonuser_no As String
        Public Property Vari_Kanri_gyfax As String
        Public Property Vari_Homeelevator_flg As String
        Public Property Vari_Horei_dosyatokubetutiiki As String
        Public Property Vari_Horei_zoseitakutitiiki As String
        Public Property Vari_Horei_tunamitiiki As String
        Public Property Vari_Horei_dosyatiikibiko As String
        Public Property Vari_Horei_dosyatokubetutiikibiko As String
        Public Property Vari_Horei_zoseitakutitiikibiko As String
        Public Property Vari_Horei_tunamitiikibiko As String
        Public Property Vari_Svbunrui_no As String
        Public Property Vari_Krbunrui_no As String
        Public Property Vari_Kanrinin_name As String
        Public Property Vari_Kozo_other As String
        Public Property Vari_Kadoti_flg As String
        Public Property Vari_Cityplan As String
        Public Property Vari_Yototiki As String
        Public Property Vari_Kanrinin_namesjis As String
        Public Property Vari_Kanri_gytantosjis As String
        Public Property Vari_Kanri_gynamesjis As String
        Public Property Vari_Parking_caraki As String
        Public Property Vari_Parking_bikeaki As String
        Public Property Vari_Parking_bicycleaki As String
        Public Property Vari_Syogaku_name As String
        Public Property Vari_Syogaku_kyori As String
        Public Property Vari_Cyugaku_name As String
        Public Property Vari_Cyugaku_kyori As String
        Public Property Vari_Bk_area_no As String
        Public Property Vari_Emergencyelevator_flg As String   '20160519 EXEUpdateに伴う修正 物件詳細情報 -del

    End Class

#End Region

#Region "物件所有者情報"

    Public Class Bkdata_syo_Model

        Public Property Vari_Bk_guid As String
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

#Region "物件ゴミ情報モデル"

    Public Class Bkdata_dust_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Dust_no As String
        Public Property Vari_Dust_rui As String
        Public Property Vari_Dust_youbi As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "×物件画像情報"

    Public Class Bkdata_gazo_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Bk_gazono As String
        Public Property Vari_Bk_gazocomment As String
        Public Property Vari_Bk_gazolastupdate As String
        Public Property Vari_Bk_gazofilesizekbyte As String
        Public Property Vari_Bk_gazowidth As String
        Public Property Vari_Bk_gazoheight As String
        Public Property Vari_Bk_gazothumbimage As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "物件変動費親メーター情報(汎用ツールのみ)"

    Public Class Bkdata_hendo_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Bkhendo_guid As String
        Public Property Vari_Rec_no As String
        Public Property Vari_Hendo_kbn As String
        Public Property Vari_Meter_name As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Child_meter As String
        Public Property Vari_Biko As String
        Public Property Vari_Useflg As String
        Public Property Vari_History As String
        Public Property Vari_Sq_mmkbn As String
        Public Property Vari_Tani As String
        Public Property Vari_Tanisjis As String

    End Class

#End Region

#Region "物件鍵情報(汎用ツールのみ)"

    Public Class Bkdata_kagi_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Kagi_no As String
        Public Property Vari_Honsu As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Hokan As String
        Public Property Vari_Gyshare As String

    End Class

#End Region

#Region "物件権利情報"

    Public Class Bkdata_kenri_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Kenri_no As String
        Public Property Vari_Other_kenrirui As String
        Public Property Vari_Other_kenribiko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "×物件KeyValueテーブル"

    Public Class Bkdata_keyvalue_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Data_kbn As String
        Public Property Vari_Item_kbn As String
        Public Property Vari_Value As String

    End Class

#End Region

#Region "物件交通情報"

    Public Class Bkdata_kotu_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Ensen_cnt As String
        Public Property Vari_Ensen_no As String
        Public Property Vari_Eki_no As String
        Public Property Vari_Firsttrain_flg As String
        Public Property Vari_Kyori As String
        Public Property Vari_Toho_min As String
        Public Property Vari_Car_min As String
        Public Property Vari_Bus_min As String
        Public Property Vari_Bus_company As String
        Public Property Vari_Bus_station As String
        Public Property Vari_Bus_tohomin As String
        Public Property Vari_Bus_kyori As String
        Public Property Vari_History As String
        Public Property Vari_Setting_kotutype As String
        Public Property Vari_Bus_companytoeki As String
        Public Property Vari_Bus_stationtoeki As String
        Public Property Vari_Bus_kyoritoeki As String
        Public Property Vari_Bus_tohomintoeki As String
        Public Property Vari_Kotu_syudan As String
        Public Property Vari_Car_kyori As String

    End Class

#End Region

#Region "物件近隣駐車場情報(汎用ツールのみ)"

    Public Class Bkdata_parkingother_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Parking_no As String
        Public Property Vari_Naiyo As String
        Public Property Vari_Kyori As String
        Public Property Vari_Gak As String
        Public Property Vari_Zeikbn As String

    End Class

#End Region

#Region "物件参照ファイル情報(汎用ツールのみ)"

    Public Class Bkdata_relfile_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_File_no As String
        Public Property Vari_Fullpath As String
        Public Property Vari_Addtime As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "物件接道情報"

    Public Class Bkdata_setudo_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Setudo_cnt As String
        Public Property Vari_Setudo_muki As String
        Public Property Vari_Setudo_roadpattern As String
        Public Property Vari_Setudo_haba As String
        Public Property Vari_Setudo_setudokyori As String
        Public Property Vari_Setudo_pointselectflg As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "物件周辺情報"

    Public Class Bkdata_syuhen_Model

        Public Property Vari_Syuhen_guid As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Syuhen_cnt As String
        Public Property Vari_Sisetu_kbn As String
        Public Property Vari_Sisetu_name As String
        Public Property Vari_Sisetu_kyori As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "×物件周辺画像情報"

    Public Class Bkdata_syuhengazo_Model

        Public Property Vari_Syuhen_guid As String
        Public Property Vari_Syuhen_gazono As String
        Public Property Vari_Syuhen_gazocomment As String
        Public Property Vari_Syuhen_gazolastupdate As String
        Public Property Vari_Syuhen_gazofilesizebyte As String
        Public Property Vari_Syuhen_gazowidth As String
        Public Property Vari_Syuhen_gazoheight As String
        Public Property Vari_Syuhen_gazothumbimage As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "物件修繕維持管理連絡先情報"

    Public Class Bkdata_szeniji_Model

        Public Property Vari_Bk_guid As String
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

#Region "×物件bkdata_youtube情報"

    Public Class Bkdata_youtube_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_No As String
        Public Property Vari_Uri As String
        Public Property Vari_Comment As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "物件メモ情報"

    Public Class Bkdata_memo_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
