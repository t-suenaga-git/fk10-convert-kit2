Namespace Njc.Model

#Region "物件分類モデル"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_bk_rui_Model

    '    Public Property Vari_Bk_ruino As String
    '    Public Property Vari_Bk_ruiname As String
    '    Public Property Vari_Bk_ruibiko As String
    '    Public Property Vari_Bk_ruiuseflg As String
    '    Public Property Vari_Sincyoku_keiyakukbn As String
    '    Public Property Vari_Sincyoku_kosinkbn As String
    '    Public Property Vari_Sincyoku_kaiyakukbn As String
    '    Public Property Vari_History As String
    '    Public Property Vari_Homemate_ruikbn As String
    '    Public Property Vari_Bk_endofmonthflg As String
    '    Public Property Vari_Ikkatukariage_siwakekbn As String

    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end

#End Region

#Region "バスモデル"

    Public Class M_buskotu_Model

        Public Property Vari_Buskotu_no As String
        Public Property Vari_Buskotu_name As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "バス交通モデル"

    Public Class M_buskotu_stop_Model

        Public Property Vari_Buskotu_no As String
        Public Property Vari_Sortorder As String
        Public Property Vari_Keito_name As String
        Public Property Vari_Busstop_name As String

    End Class

#End Region

#Region "鍵タイトルモデル"

    Public Class M_kagi_title_Model

        Public Property Vari_Kagi_kbn As String     '鍵区分
        Public Property Vari_Kagi_no As String      '鍵タイトルNo
        Public Property Vari_Kagi_name As String    '鍵タイトル名
        Public Property Vari_History As String

    End Class

#End Region

#Region "箇所分類マスタモデル"

    Public Class M_claim_rui_Model

        Public Property Vari_Claim_ruikbn As String         '1:箇所区分 2:クレーム分類
        Public Property Vari_Claim_ruino As String
        Public Property Vari_Claim_ruisortorder As String
        Public Property Vari_Claim_name As String
        Public Property Vari_Claim_useflg As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "クレーム分類マスタモデル"

    'Public Class M_claim_rui_claim_Model

    '    Public Property Vari_Claim_ruikbn As String         '1:箇所区分 2:クレーム分類
    '    Public Property Vari_Claim_ruino As String
    '    Public Property Vari_Claim_ruisortorder As String
    '    Public Property Vari_Claim_name As String
    '    Public Property Vari_Claim_useflg As String
    '    Public Property Vari_History As String

    'End Class

#End Region

#Region "特約事項マスタモデル"

    Public Class M_tokuyaku_Model

        Public Property Vari_Tokuyaku_grpno As String       '特約グループNo
        Public Property Vari_Tokuyaku_no As String          '特約No
        Public Property Vari_Tokuyaku_title As String       '特約タイトル
        Public Property Vari_Tokuyaku_template As String    '特約テンプレート
        Public Property Vari_History As String

    End Class

#End Region

#Region "原状回復特約マスタモデル"

    'Public Class M_tokuyaku_genjo_Model

    '    Public Property Vari_Tokuyaku_grpno As String       '特約グループNo
    '    Public Property Vari_Tokuyaku_no As String          '特約No
    '    Public Property Vari_Tokuyaku_title As String       '特約タイトル
    '    Public Property Vari_Tokuyaku_template As String    '特約テンプレート
    '    Public Property Vari_History As String

    'End Class

#End Region

#Region "契約分類マスタモデル"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_ky_rui_Model

    '    Public Property Vari_Ky_ruino As String         '契約分類No
    '    Public Property Vari_Ky_ruiname As String       '契約分類名
    '    Public Property Vari_Ky_ruibiko As String       '契約分類備考
    '    Public Property Vari_Ky_ruitutimm As String     '契約終了通知月（◯ヶ月前）
    '    Public Property Vari_Ky_ruicolor As String      '契約分類色
    '    Public Property Vari_Teisyaku_flg As String     '1:使用する 2:使用しない
    '    Public Property Vari_History As String
    '    Public Property Vari_Useflg As String           '1:検索対象 0:検索対象外

    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end

#End Region

#Region "保険種類マスタモデル"

    Public Class M_hoken_rui_Model

        Public Property Vari_Hoken_ruino As String
        Public Property Vari_Hoken_ruiname As String
        Public Property Vari_Hoken_ruikana As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_History As String
        Public Property Vari_Useflg As String

    End Class

#End Region

#Region "入金区分マスタモデル"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_nkbn_Model

    '    Public Property Vari_Nkbn_no As String          '1-99のキー(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
    '    Public Property Vari_Nkbn_order As String       '表示順
    '    Public Property Vari_Nkbn_name As String        '入金区分名(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
    '    Public Property Vari_Nkbn_shortname As String   '入金区分略称
    '    Public Property Vari_Nkbn_zokusei As String     '1:現金, 2:振込, 3:振替, 4:その他, 99:移動
    '    Public Property Vari_History As String

    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end
#End Region

#Region "学校区マスタモデル"

    Public Class M_koku_add_Model

        Public Property Vari_Ken_no As String           '県No
        Public Property Vari_Si_no As String            '市No
        Public Property Vari_No As String               '学校区No
        Public Property Vari_Add_cyo As String          '町地域
        Public Property Vari_Add_banti As String        '番地
        Public Property Vari_Syogaku_name As String     '小学校名
        Public Property Vari_Cyugaku_name As String     '中学校名
        Public Property Vari_History As String
        Public Property Vari_Add_cyome As String        'NULL
       
    End Class

#End Region

#Region "エリアマスタモデル"

    Public Class M_area_Model

        Public Property Vari_Area_no As String          'エリアNo
        Public Property Vari_Area_name As String        'エリア名
        Public Property Vari_Area_sortorder As String   '表示順
        Public Property Vari_Area_useflg As String      '使用有無
        Public Property Vari_History As String

    End Class

#End Region

#Region "入金項目マスタモデル"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_nkin_Model

    '    Public Property Vari_Nkin_no As String
    '    Public Property Vari_Nkin_name As String
    '    Public Property Vari_Nkin_printname As String
    '    Public Property Vari_Nkin_ruino As String
    '    Public Property Vari_Nkin_zkseino As String
    '    Public Property Vari_Nkin_useflg As String          '1:使用する, 2:使用しない
    '    Public Property Vari_Keiyaku_nkinno As String
    '    Public Property Vari_Keiyakuhiki_nkinno As String
    '    Public Property Vari_Kosin_nkinno As String
    '    Public Property Vari_Kaiyaku_nkinno As String
    '    Public Property Vari_Kaiyakuhiki_nkinno As String
    '    Public Property Vari_Biko_nkin As String
    '    Public Property Vari_History As String
    '    Public Property Vari_Ryosyu_flg As String           '1:有り, 2:なし, -1:未定義

    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end

#End Region

#Region "口座種別マスタモデル"

    Public Class M_kozasyu_Model

        Public Property Vari_Kosyu_no As String
        Public Property Vari_Kosyu_name As String
        Public Property Vari_Kosyu_sname As String
        Public Property Vari_History As String
     
    End Class

#End Region

#Region "変動費マスタモデル"

    Public Class M_hendorule_Model

        Public Property Vari_Rule_no As String
        Public Property Vari_Rule_name As String
        Public Property Vari_Tani As String
        Public Property Vari_Tanisjis As String
        Public Property Vari_Ryokin_cnt As String
        Public Property Vari_Hendo_rui As String
        Public Property Vari_Hasuusyori_sel As String
        Public Property Vari_Hasuusyosu_ptn As String
        Public Property Vari_Koukei As String
        Public Property Vari_Biko_kihon As String
        Public Property Vari_Keijo_rui As String
        Public Property Vari_Zei_kbnhayami As String
        Public Property Vari_Zei_kbntanka As String
        Public Property Vari_History As String
        Public Property Vari_Useflg As String

    End Class

#End Region

#Region "変動費一覧マスタモデル"

    Public Class M_hendorule_itiran_Model

        Public Property Vari_Rule_no As String
        Public Property Vari_Item_no As String
        Public Property Vari_Siyoryo As String
        Public Property Vari_Ryokin1 As String
        Public Property Vari_Ryokin2 As String
        Public Property Vari_Ryokin3 As String
        Public Property Vari_Ryokin4 As String
        Public Property Vari_Ryokin5 As String

    End Class

#End Region

#Region "備考タイトルマスタ" '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add

    Public Class M_memo_Model

        Public Property Vari_Memo_kbn As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo_sortorder As String
        Public Property Vari_Memo_name As String
        Public Property Vari_Memo_useflg As String
        Public Property Vari_Memo_usehojoitems As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "備考入力補助リストマスタ"  '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add

    Public Class M_memo_lst_Model

        Public Property Vari_Memo_kbn As String
        Public Property Vari_Memo_no As String
        Public Property Vari_Memo_lstno As String
        Public Property Vari_Memo_lstname As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "画像タイトルマスタ" '20160720 連動情報構築

    Public Class M_gazo_title_Model

        Public Property Vari_Gazo_kbn As String
        Public Property Vari_Gazo_no As String
        Public Property Vari_Gazo_name As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
