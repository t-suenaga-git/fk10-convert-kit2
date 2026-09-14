Imports Converter10.Njc.Common
Imports Converter10.Njc.N3Lib.Utys
Imports System.IO
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Microsoft.Win32             '2016.02.22 V7DB接続情報取得処理追加
Imports System.Data.OleDb

Namespace Njc.Frm

    Public Class MainFrm

#Region "宣言"

        '★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        ' 20160705 確認メモ
        ' 使用する前に初期化を行っているかを再確認する
        ' (CV後に再度CVをそのまま行った時に問題ない)
        '★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

        '[接続情報]
        Private sqlcnnv7 As System.Data.SqlClient.SqlConnection     'V7接続情報
        Private sqlcnnv10 As System.Data.SqlClient.SqlConnection    '10接続情報
        Private fstmodelv7 As New Njc.Model.DefSQLConnection        'V7接続情報格納モデル
        Private fstmodelv10 As New Njc.Model.DefSQLConnection       '10接続情報格納モデル
        Private cnnv7 As New Njc.Common.DBConnection                'V7接続処理用クラス
        Private cnnv10 As New Njc.Common.DBConnection               '10接続処理用クラス

        '[各タブ生成オブジェクト]
        Private tabPageManager As Njc.Common.TabPageManager         'タブページ表示用(メイン画面タブ)
        Private tabPageManagerCvitem As Njc.Common.TabPageManager   'タブページ表示用(項目選択タブ)
        Private tabPageManagerJizen As Njc.Common.TabPageManager    'タブページ表示用(事前作業タブ)
        Private tabPageManagerJigo As Njc.Common.TabPageManager     'タブページ表示用(事後作業タブ)
        Private tabPageManagerHojyo As Njc.Common.TabPageManager    'タブページ表示用(補助機能タブ)

        '[その他]
        Private logset As New Njc.Common.LogSetting                 'ログ出力用
        Private model_cvitem As New Njc.Model.CVItemInfo            '移行項目名称格納用
        Private obj_com As New Njc.Common.CommonRepository          '共通処理用
        Private hash_deltblqrywhere As New Hashtable                'テーブル名と削除対象を紐付値格納用リスト

        '[処理件数関連]
        Private pgbtotalcnt As Integer                              'プログレスバー全件数
        Private pgbpartcnt As Integer                               'プログレスバー個別件数
        Private beforemidfilechkflg As Boolean = False              '中間ファイルチェック前準備完了フラグ

        '[汎用コンバートキット用]
        Private list_baseCVchkitem As New List(Of CheckBox)         '20160516 汎用のみ表示するチェックボックスのチェックをOFFにする処理を追加 -add
        Private list_existCVchkitem As New List(Of CheckBox)        '20160905 汎用の場合の移行対象外項目(既存のみ移行対象)選定の修正 -add

        '[パス]
        Private dcv_exepath As String = System.Reflection.Assembly.GetExecutingAssembly().Location                  'データコンバーター実行ファイル格納パス]
        Private dcv_exedir As String = IO.Path.GetDirectoryName(dcv_exepath)                                        'データコンバーターディレクトリパス

        '[コンバーター対象項目選択タブ名]
        Private Const DC_SELTAB_K01 As String = "tabPageKizon110"
        Private Const DC_SELTAB_K02 As String = "tabPageKizon120"
        Private Const DC_SELTAB_K03 As String = "tabPageKizon130"
        Private Const DC_SELTAB_B01 As String = "tabPageBase110"
        Private Const DC_SELTAB_B02 As String = "tabPageBase120"
        Private Const DC_SELTAB_B03 As String = "tabPageBase130"
        Private Const DC_SELTAB_B04 As String = "tabPageBase140"
        Private Const DC_SELTAB_B05 As String = "tabPageBase150"
        Private Const DC_SELTAB_B06 As String = "tabPageBase160"
        Private Const DC_SELTAB_B07 As String = "tabPageBase170"
        Private Const DC_SELTAB_B08 As String = "tabPageBase180"
        Private Const DC_SELTAB_B09 As String = "tabPageBase190"
        Private Const DC_SELTAB_B10 As String = "tabPageBase200"
        Private Const DC_SELTAB_B11 As String = "tabPageBase210"                                                    '20160720 連動情報構築 -add
        Private Const DC_SELTAB_H01 As String = "tabPageHanyo110"
        Private Const DC_SELTAB_H02 As String = "tabPageHanyo120"
        Private Const DC_SELTAB_H03 As String = "tabPageHanyo130"
        Private Const M_SYOKITAB_OPS As String = "pnlOptionSelect"                                                  'メイン.初期設定.オプション選択パネル
        '20160719 V7オプション選択による表示設定処理追加_不足分対応 -add sta
        '画面起動時にチェックボックスイベントに入るため起動時は入らないようにするフラグ
        '画面起動イベント(FrmMain_Load)の最後にFalseに変更して保持しておく
        Private frmloadflg_kysq As Boolean = True                   '画面起動時判別用フラグ_契約請求関連チェックボックス制御用(True:画面起動時 False:左以外)
        Private frmloadflg_clsz As Boolean = True                   '画面起動時判別用フラグ_クレーム修繕関連チェックボックス制御用(True:画面起動時 False:左以外)
        '20160719 V7オプション選択による表示設定処理追加_不足分対応 -add end
        '20160720 紐付設定ファイル出力機能の追加 -add
        Private relitemstr As String = ""       '紐付呼出時の項目格納用
        '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add                                       
        Private relexeerr As Integer = 0       '紐付ツール終了時の終了コード格納用(0:正常終了 1:キャンセル 99:異常終了)

#End Region

#Region "※既存ユーザ用処理"

        ''' <summary>
        ''' 中間ファイル書込Repositoryの生成
        ''' </summary>
        ''' <param name="chkboxtext"></param>
        ''' <remarks></remarks>
        Private Function Get_ObjRep_Mid(ByVal chkboxtext As String) As Object

            Dim obj_rep As Object = Nothing

            Select Case chkboxtext

                '各マスタ情報
                Case model_cvitem.MstBus : obj_rep = New Njc.Repository.M_bus_V10_Repository.SubConv                        'バス交通マスタ(革命10用)
                Case model_cvitem.MstBusKotu : obj_rep = New Njc.Repository.M_bus_Repository.SubConv                        'バス停マスタ
                Case model_cvitem.MstKagititle : obj_rep = New Njc.Repository.M_biko4_Repository.SubConv                    '鍵タイトルマスタ
                Case model_cvitem.MstKasyoClaimrui : obj_rep = New Njc.Repository.Claim_data_Rev7_Repository.SubConv        '箇所クレーム分類マスタ
                Case model_cvitem.MstTokuyaku : obj_rep = New Njc.Repository.M_tokuyaku_Rev7_Repository.SubConv             '特約事項マスタ
                Case model_cvitem.MstKeiyakurui : obj_rep = New Njc.Repository.M_keirui_Repository.SubConv                  '契約分類マスタ
                Case model_cvitem.MstHokenrui : obj_rep = New Njc.Repository.M_hoken_rui_Rev7_Repository.SubConv            '保険種類マスタ
                Case model_cvitem.MstSchool : obj_rep = New Njc.Repository.M_school_Repository.SubConv                      '学校区マスタ
                Case model_cvitem.MstArea : obj_rep = New Njc.Repository.M_addkbn_Repository.SubConv                        'エリアマスタ
                Case model_cvitem.MstHendo : obj_rep = New Njc.Repository.M_hendo_Repository.SubConv                        '変動費マスタ
                Case model_cvitem.MstHendoitiran : obj_rep = New Njc.Repository.M_hendo_mei_Repository.SubConv              '変動費一覧マスタ
                Case model_cvitem.MstBikotitle : obj_rep = New Njc.Repository.M_biko_Repository.SubConv                     '備考タイトルマスタ 
                Case model_cvitem.MstBikolst : obj_rep = New Njc.Repository.M_biko_lst_Repository.SubConv                   '備考入力補助リストマスタ
                Case model_cvitem.MstGazotitle : obj_rep = New Njc.Repository.M_gazo_syoki_Repository.SubConv               '画像タイトルマスタ            '20160801 タイトルマスタ統合処理 -add

                    '業者情報
                Case model_cvitem.GyCyukaiBase : obj_rep = New Njc.Repository.M_gy_Repository.SubConv                       '仲介業者情報
                Case model_cvitem.GyCyukaiKoza : obj_rep = New Njc.Repository.M_gy_Koza_Repository.SubConv                  '仲介業者口座情報
                Case model_cvitem.GyCyukaiMemo : obj_rep = New Njc.Repository.M_gy_Biko_Repository.SubConv                  '仲介業者メモ情報
                Case model_cvitem.GyHokenBase : obj_rep = New Njc.Repository.M_hoken_gy_Repository.SubConv                  '保険業者情報
                Case model_cvitem.GyHokenKoza : obj_rep = New Njc.Repository.M_hoken_gy_Koza_Repository.SubConv             '保険業者口座情報
                Case model_cvitem.GyHokenMemo : obj_rep = New Njc.Repository.M_hoken_gy_Biko_Repository.SubConv             '保険業者メモ情報
                Case model_cvitem.GyYatinhosyoBase : obj_rep = New Njc.Repository.M_hosyo_gy_Repository.SubConv             '家賃保証業者情報
                Case model_cvitem.GyYatinhosyoMemo : obj_rep = New Njc.Repository.M_hosyo_gy_Biko_Repository.SubConv        '家賃保証業者メモ情報
                Case model_cvitem.GySyuzenBase : obj_rep = New Njc.Repository.M_gysyuzen_Repository.SubConv                 '修繕業者情報
                Case model_cvitem.GySyuzenKoza : obj_rep = New Njc.Repository.M_gysyuzen_Koza_Repository.SubConv            '修繕業者口座情報
                Case model_cvitem.GySyuzenMemo : obj_rep = New Njc.Repository.M_gysyuzen_Biko_Repository.SubConv            '修繕業者メモ情報
                Case model_cvitem.GyLifelineBase : obj_rep = New Njc.Repository.M_kokyo_Repository.SubConv                  'ライフライン業者情報
                Case model_cvitem.GySekoBase : obj_rep = New Njc.Repository.M_seko_kaisya_Repository.SubConv                '施工業者情報
                Case model_cvitem.GySisetuBase : obj_rep = New Njc.Repository.M_hosyu_kaisya_Repository.SubConv             '保守業者情報

                    '自社情報
                Case model_cvitem.JisyaBase : obj_rep = New Njc.Repository.M_jisya_Repository.SubConv                       '自社情報
                Case model_cvitem.JisyaKoza : obj_rep = New Njc.Repository.M_jisya_Koza_Repository.SubConv                  '自社口座情報
                Case model_cvitem.JisyaTanto : obj_rep = New Njc.Repository.M_tanto_Repository.SubConv                      '自社担当者情報      
                Case model_cvitem.JisyaMemo : obj_rep = New Njc.Repository.M_jisya_Biko_Repository.SubConv                  '自社メモ情報
                Case model_cvitem.FBFuriirai : obj_rep = New Njc.Repository.M_furi_irai_Repository.SubConv                  '振込依頼人情報

                    '口座関連情報
                Case model_cvitem.FBFuritesuryo : obj_rep = New Njc.Repository.M_furi_tesu_Repository.SubConv               '振込手数料情報
                Case model_cvitem.FBKozafurikae : obj_rep = New Njc.Repository.M_koza_furi_Repository.SubConv               '口座振替情報
                Case model_cvitem.FBNsSyutoku : obj_rep = New Njc.Repository.M_ns_syutoku_Repository.SubConv                '入出金取得情報
                Case model_cvitem.MstYatinKoza : obj_rep = New Njc.Repository.M_furi_koza_Repository.SubConv                '家賃入金口座情報
                Case model_cvitem.MstANSERArea : obj_rep = New Njc.Repository.Mspc_area_Repository.SubConv                  'ANSERエリア情報              '20160704 ANSER情報の移行処理追加 -add
                Case model_cvitem.MstANSERAccpoint : obj_rep = New Njc.Repository.Mspc_accpoint_Repository.SubConv          'ANSERアクセスポイント情報    '20160704 ANSER情報の移行処理追加 -add
                Case model_cvitem.FBANSERSetuzoku : obj_rep = New Njc.Repository.Mspc_setuzoku_Repository.SubConv           'ANSER接続情報                '20160704 ANSER情報の移行処理追加 -add

                    '家主情報
                Case model_cvitem.OwBase : obj_rep = New Njc.Repository.M_yanu_so_Repository.SubConv                        '家主情報
                Case model_cvitem.OwKoza : obj_rep = New Njc.Repository.M_yanu_so_koza_Repository.SubConv                   '家主口座情報
                Case model_cvitem.OwEvent : obj_rep = New Njc.Repository.M_yanu_so_event_Repository.SubConv                 '家主イベント情報
                Case model_cvitem.OwMemo : obj_rep = New Njc.Repository.M_yanu_so_biko_Repository.SubConv                   '家主メモ情報

                    '契約者情報
                Case model_cvitem.KysBase : obj_rep = New Njc.Repository.Kys_mst_Repository.SubConv                         '契約者情報
                Case model_cvitem.KysKoza : obj_rep = New Njc.Repository.Kys_mst_koza_Repository.SubConv                    '契約者口座情報
                Case model_cvitem.KysSyogoKana : obj_rep = New Njc.Repository.Kys_mst_syogokana_Repository.SubConv          '契約者照合用カナ情報
                Case model_cvitem.KysHosyonin : obj_rep = New Njc.Repository.Kys_mst_hosyonin_Repository.SubConv            '契約保証人者情報
                Case model_cvitem.KysMemo : obj_rep = New Njc.Repository.Kys_biko_Repository.SubConv                        '契約者メモ情報

                    '物件情報
                Case model_cvitem.BkBase : obj_rep = New Njc.Repository.Bk_mst_Repository.SubConv                           '物件基本情報
                Case model_cvitem.BkSyosai : obj_rep = New Njc.Repository.Bk_mst_syosai_Repository.SubConv                  '物件詳細情報
                Case model_cvitem.Bksyo : obj_rep = New Njc.Repository.Bk_mst_syo_Repository.SubConv                        '物件所有者情報
                Case model_cvitem.BkGomi : obj_rep = New Njc.Repository.Bk_mst_gomi_Repository.SubConv                      '物件ゴミ情報
                Case model_cvitem.BkKenri : obj_rep = New Njc.Repository.Bk_Jyuyo_Repository.SubConv                        '物件権利情報
                Case model_cvitem.BkKotu : obj_rep = New Njc.Repository.Bk_eki_Repository.SubConv                           '物件交通情報
                Case model_cvitem.BkSetudo : obj_rep = New Njc.Repository.Bk_mst_setudo_Repository.SubConv                  '物件接道情報
                Case model_cvitem.BkSyuhen : obj_rep = New Njc.Repository.Bk_mst_syuhen_Repository.SubConv                  '物件周辺情報
                Case model_cvitem.BkSzeniji : obj_rep = New Njc.Repository.Bk_setubi_ren_bk_Repository.SubConv              '物件修繕維持管理連絡先情報
                Case model_cvitem.BkMemo : obj_rep = New Njc.Repository.Bk_biko_Repository.SubConv                          '物件メモ情報
                Case model_cvitem.BkKagi : obj_rep = New Njc.Repository.Bk_mst_kagi_Repository.SubConv                      '物件鍵情報                  ※汎用分(V7には無い)
                Case model_cvitem.BkHendo : obj_rep = New Njc.Repository.Bk_mst_hendo_Repository.SubConv                    '物件変動費親メーター情報    ※汎用分(V7には無い)
                Case model_cvitem.BkKinrincyusyajo : obj_rep = New Njc.Repository.Bk_mst_kinrincyushajo_Repository.SubConv  '物件近隣駐車場情報          ※汎用分(V7には無い)
                Case model_cvitem.BkSansyofile : obj_rep = New Njc.Repository.Bk_mst_sansyofile_Repository.SubConv          '物件参照ファイル情報        ※汎用分(V7には無い)

                    '部屋情報
                Case model_cvitem.HyBase : obj_rep = New Njc.Repository.Hy_mst_Repository.SubConv                           '部屋基本情報
                Case model_cvitem.HySyosai : obj_rep = New Njc.Repository.Hy_mst_syosai_Repository.SubConv                  '部屋詳細情報
                Case model_cvitem.HyParking : obj_rep = New Njc.Repository.Hy_mst_parking_Repository.SubConv                '部屋駐車場情報
                Case model_cvitem.HyTokuyaku : obj_rep = New Njc.Repository.Hy_tokuyaku_Repository.SubConv                  '部屋特約情報
                Case model_cvitem.HyKagi : obj_rep = New Njc.Repository.Hy_kagi_Repository.SubConv                          '部屋鍵情報
                Case model_cvitem.HyMadoriutiwake : obj_rep = New Njc.Repository.Hy_mst_madriutiwake_Repository.SubConv     '部屋間取内訳情報
                Case model_cvitem.HyMenseki : obj_rep = New Njc.Repository.Hy_mst_menseki_Repository.SubConv                '部屋面積情報
                Case model_cvitem.HySetubi : obj_rep = New Njc.Repository.Hy_setubi_Repository.SubConv                      '部屋設備情報
                Case model_cvitem.HyNkinkomk : obj_rep = New Njc.Repository.Hy_kanri_Repository.SubConv                     '部屋入金項目情報
                Case model_cvitem.HyHendo : obj_rep = New Njc.Repository.Hy_kanri_hendo_Repository.SubConv                  '部屋変動費各戸メーター情報
                Case model_cvitem.HyMemo : obj_rep = New Njc.Repository.Hy_biko_Repository.SubConv                          '部屋メモ情報
                Case model_cvitem.HyCommonsalespoint : obj_rep = New Njc.Repository.Hy_commonsalespoint_Repository.SubConv  '部屋共通セールスポイント情報
                Case model_cvitem.HyConfirm : obj_rep = New Njc.Repository.Hy_confirm_Repository.SubConv                    '部屋契約解約確認事項情報
                Case model_cvitem.HyKenri : obj_rep = New Njc.Repository.Hy_kenri_Repository.SubConv                        '部屋権利情報
                Case model_cvitem.HySansyofile : obj_rep = New Njc.Repository.Hy_relfile_Repository.SubConv                 '部屋参照ファイル情報
                Case model_cvitem.HyGenjotanka : obj_rep = New Njc.Repository.Hy_szen_Repository.SubConv                    '部屋原状回復目安単価情報
                Case model_cvitem.HySzeniji : obj_rep = New Njc.Repository.Hy_szeniji_Repository.SubConv                    '部屋修繕維持管理連絡先情報

                    '送金ルール情報
                Case model_cvitem.SoruleBase : obj_rep = New Njc.Repository.Bk_kanri_Repository.SubConv                     '送金ルール基本情報
                Case model_cvitem.SoruleSosaki : obj_rep = New Njc.Repository.Bk_sokin_Repository.SubConv                   '送金ルール送金先情報
                Case model_cvitem.SoruleNkin : obj_rep = New Njc.Repository.Hy_kanri_sorulenkin_Repository.SubConv          '送金ルール入金項目情報
                Case model_cvitem.SoruleKojo : obj_rep = New Njc.Repository.Hy_kanri_sorulekojo_Repository.SubConv          '送金ルール控除項目情報

                    '契約情報
                Case model_cvitem.KyBase : obj_rep = New Njc.Repository.Ky_kosinkai_Repository.SubConv                      '契約基本情報
                Case model_cvitem.KyRireki : obj_rep = New Njc.Repository.Ky_kosinkai_rireki_Repository.SubConv             '契約履歴情報
                Case model_cvitem.KyCar : obj_rep = New Njc.Repository.Ky_car_Repository.SubConv                            '契約車情報
                Case model_cvitem.KyKys : obj_rep = New Njc.Repository.Ky_kosinkai_kys_Repository.SubConv                   '契約契約者情報
                Case model_cvitem.KyHosyonin : obj_rep = New Njc.Repository.Ky_kosinkai_hosyonin_Repository.SubConv         '契約保証人情報
                Case model_cvitem.KyNyukyo : obj_rep = New Njc.Repository.Ky_kazoku_Repository.SubConv                      '契約入居者情報
                Case model_cvitem.KyTokuyaku : obj_rep = New Njc.Repository.ky_tokuyaku_Repository.SubConv                  '契約特約事項情報
                Case model_cvitem.KyHoken : obj_rep = New Njc.Repository.ky_Hoken_Repository.SubConv                        '契約保険情報
                Case model_cvitem.KyMemo : obj_rep = New Njc.Repository.Ky_biko_Repository.SubConv                          '契約メモ情報
                Case model_cvitem.KyNkinkomk : obj_rep = New Njc.Repository.Ky_sqdata_Repository.SubConv                    '契約入金項目情報
                Case model_cvitem.KyNkinkomkNx : obj_rep = New Njc.Repository.Ky_sqdata_nx_Repository.SubConv               '契約次回入金項目情報
                Case model_cvitem.KyHendo : obj_rep = New Njc.Repository.Ky_sqdata_hendo_Repository.SubConv                 '契約変動費各戸メーター情報
                Case model_cvitem.KyKojoRule : obj_rep = New Njc.Repository.Ky_sqdata_kojo_Repository.SubConv               '契約控除ルール情報
                Case model_cvitem.KySorule : obj_rep = New Njc.Repository.Ky_sqdata_sorule_Repository.SubConv               '契約送金ルール情報
                Case model_cvitem.KyKai : obj_rep = New Njc.Repository.Ky_kosinkai_kai_Repository.SubConv                   '契約解約情報
                    '20160531 鍵情報移行処理の修正 -del sta
                    'Case model_cvitem.KyKagi : obj_rep = New Njc.Repository.Ky_kagi_Repository.SubConv                          '契約鍵情報
                    '20160531 鍵情報移行処理の修正 -del end
                Case model_cvitem.KySzen : obj_rep = New Njc.Repository.Reform_data_kaiszen_Repository.SubConv              '契約修繕見積情報   
                Case model_cvitem.KySzenmeisai : obj_rep = New Njc.Repository.Reform_item_kaiszen_Repository.SubConv        '契約修繕見積詳細情報

                    '請求情報
                Case model_cvitem.SqKajyo : obj_rep = New Njc.Repository.Sq_kajo_Repository.SubConv
                Case model_cvitem.SqUnyotaino : obj_rep = New Njc.Repository.Sq_unyomisyutaino_Repository.SubConv
                Case model_cvitem.SqSq : obj_rep = New Njc.Repository.Sq_sq_Repository.SubConv
                Case model_cvitem.SqHendokensin : obj_rep = New Njc.Repository.Sq_hendokensin_Repository.SubConv
                Case model_cvitem.SqKoteiKojo : obj_rep = New Njc.Repository.Sq_koteikojo_Repository.SubConv
                Case model_cvitem.SqSqKojo : obj_rep = New Njc.Repository.Sq_sqkojo_Repository.SubConv

                    'クレーム情報
                Case model_cvitem.ClaimBase : obj_rep = New Njc.Repository.Claim_data_Base_Repository.SubConv
                Case model_cvitem.ClaimTaiorireki : obj_rep = New Njc.Repository.Claim_data_taiorireki_Repository.SubConv
                Case model_cvitem.ClaimRelfile : obj_rep = New Njc.Repository.Claim_data_relfile_Repository.SubConv

                    '修繕関連移
                Case model_cvitem.SzenBase : obj_rep = New Njc.Repository.Reform_data_Base_Repository.SubConv
                Case model_cvitem.SzenSzen : obj_rep = New Njc.Repository.Reform_data_Mitumori_Repository.SubConv
                Case model_cvitem.SzenSzenmeisai : obj_rep = New Njc.Repository.Reform_item_Repository.SubConv
                Case model_cvitem.SzenClaim : obj_rep = New Njc.Repository.Reform_data_claim_Repository.SubConv
                Case model_cvitem.SzenRelfile : obj_rep = New Njc.Repository.Reform_data_relfile_Repository.SubConv         '20160627 修繕関連ファイル移行修正
                Case model_cvitem.SzenMemo : obj_rep = New Njc.Repository.Reform_jyuyo_Repository.SubConv

                    '初期設定情報
                Case model_cvitem.SyskanriBase : obj_rep = New Njc.Repository.M_kan_Base_Repository.SubConv
                Case model_cvitem.SyskanriZei : obj_rep = New Njc.Repository.M_kan_Zei_Repository.SubConv
                Case model_cvitem.SyskanriHenkanmoji : obj_rep = New Njc.Repository.M_kan_Henkanmoji_Repository.SubConv
                Case model_cvitem.SyskanriNkinkomkmerge : obj_rep = New Njc.Repository.M_kan_Nkinkomkmerge_Repository.SubConv

                    '物件データ連動情報 '20160720 連動情報構築 -add
                Case model_cvitem.RendoSosinBase : obj_rep = New Njc.Repository.Rendo_sosin_Base_Repository.SubConv
                Case model_cvitem.RendoSosinJisyaweb : obj_rep = New Njc.Repository.Rendo_sosin_jisyaweb_Repository.SubConv
                Case model_cvitem.RendoSosinHomes : obj_rep = New Njc.Repository.Rendo_sosin_homes_Repository.SubConv
                Case model_cvitem.RendoSosinAthome : obj_rep = New Njc.Repository.Rendo_sosin_athome_Repository.SubConv
                Case model_cvitem.RendoSosinSuumo : obj_rep = New Njc.Repository.Rendo_sosin_suumo_Repository.SubConv
                Case model_cvitem.RendoKokokuJisyaweb : obj_rep = New Njc.Repository.Rendo_kokoku_jisyaweb_Repository.SubConv
                Case model_cvitem.RendoKokokuHomes : obj_rep = New Njc.Repository.Rendo_kokoku_homes_Repository.SubConv
                Case model_cvitem.RendoKokokuAthome : obj_rep = New Njc.Repository.Rendo_kokoku_athome_Repository.SubConv
                Case model_cvitem.RendoKokokuSuumo : obj_rep = New Njc.Repository.Rendo_kokoku_suumo_Repository.SubConv
                Case model_cvitem.RendoHyrui : obj_rep = New Njc.Repository.Rendo_hyrui_Repository.SubConv
                Case model_cvitem.RendoHysosin : obj_rep = New Njc.Repository.Rendo_hysosin_Repository.SubConv
                Case model_cvitem.RendoBtoBgroup : obj_rep = New Njc.Repository.Rendo_BtoBgroup_Repository.SubConv
                Case model_cvitem.RendoMapdisp : obj_rep = New Njc.Repository.Rendo_mapdisp_Repository.SubConv

                Case Else   '空のまま返却

            End Select

            Return obj_rep

        End Function

        ''' <summary>
        ''' 移行元読込→中間ファイル作成処理
        ''' </summary>
        ''' <param name="list_cv"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_MiddleFile(ByVal list_cv As List(Of String), ByRef errstr As String) As Boolean

            Dim normalflg As Boolean = True
            Dim pgbcnt As Integer = 0
            Dim obj_pgb As New ProgressBarManager
            Dim tmp_totalsituation As String = ""
            Dim tmp_partsituation As String = ""
            Dim rowcnt As Integer = 0


            '--------------------------------------------------
            ' ログ/状況出力
            '--------------------------------------------------
            'ログ出力
            Dim tmp_sql_sta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MAKE_MIDSTA), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_sta, rowcnt)
            '状況出力
            tmp_partsituation = SITUATION_MIDSTA
            Call Me.Set_Situation(tmp_partsituation, 1)


            '--------------------------------------------------
            'メイン処理
            '--------------------------------------------------
            For Each cvitem In list_cv

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Exit For
                End If

                'グループボックス/チェックボックス名取得
                Dim tmp_item() As String = cvitem.Split("-")
                Dim syorigrp As String = tmp_item(0)
                Dim syoriitem As String = tmp_item(1)

                '処理Repositoryの生成(オブジェクト生成(各ブロック毎処理))
                Dim obj_rep As Object = Get_ObjRep_Mid(syoriitem)

                '----- デバッグ用処理 ----- sta
                'Dim list As New List(Of String) From {"送信設定athome情報"}
                'Dim list As New List(Of String) From {"送信設定自社web情報", "送信設定HOMES情報", "送信設定athome情報", "送信設定SUUMO情報"}
                ''Dim list As New List(Of String) From {"部屋毎送信情報", "BtoBグループ設定情報", "地図表示詳細設定情報"}
                ''Dim list As New List(Of String) From {"広告補足自社web情報", "広告補足HOMES情報", "広告補足athome情報", "広告補足SUUMO情報"}
                'Dim list As New List(Of String) From {"未収滞納金情報"}
                'If list.Contains(syoriitem) = False Then
                '    obj_rep = Nothing
                'End If
                '----- デバッグ用処理 ----- end

                '処理開始
                If obj_rep IsNot Nothing Then

                    '状況出力
                    tmp_totalsituation = syoriitem & INDENT_1 & SITUATION_MIDSTA
                    Call Me.Set_Situation(tmp_totalsituation, 0)
                    Me.lblCVItem.Text = syorigrp & " ： " & syoriitem

                    '個別処理
                    '20160829 連動中間ファイル作成処理修正 -chg sta
                    'Call Me.Setting_BeforeMakeMidFile(syorigrp)
                    Call Me.Setting_BeforeMakeMidFile(syorigrp, syoriitem)
                    '20160829 連動中間ファイル作成処理修正 -chg end

                    'VIEW作成クエリ取得(抽出クエリ＋ソート用文字列取得)
                    Dim sortstr As String = ""
                    Dim sql_makeview As String = obj_rep.Get_UseQry(sortstr)

                    '中間ファイル作成メイン処理
                    Dim makecnt As Integer = 0
                    Dim notmakecnt As Integer = 0
                    normalflg = Me.Make_MidFile(syorigrp, syoriitem, sql_makeview, sortstr, makecnt, notmakecnt, errstr)

                    'ログ出力
                    Dim tmp_sql_itemsta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(21, syoriitem, "-", Me.Chg_FlgToStr(normalflg, 2)), False)
                    DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_itemsta, rowcnt)

                    '状況出力
                    Dim tmp_str As String = ""
                    tmp_str = tmp_str & vbCrLf & INDENT_2 & SITUATION_CVITEMCVCNT & INDENT_1 & makecnt.ToString         '作成件数
                    tmp_str = tmp_str & vbCrLf & INDENT_2 & SITUATION_CVITEMNOTCVCNT & INDENT_1 & notmakecnt.ToString   '未作成件数
                    tmp_totalsituation = syoriitem & INDENT_1 & SITUATION_MIDEND
                    tmp_partsituation = INDENT_1 & Me.Chg_FlgToStr(normalflg, 1) & syoriitem & tmp_str
                    Call Me.Set_Situation(tmp_totalsituation, 0)
                    Call Me.Set_Situation(tmp_partsituation, 1)
                    If normalflg = False Then
                        Return normalflg
                    End If
                End If

                '----- 全体用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsettingTotal(pgbcnt)
                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt, True)
            Next


            '--------------------------------------------------
            ' ログ/状況出力
            '--------------------------------------------------
            'ログ出力
            Dim tmp_sql_end As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, LOG_SYORIKOMK_MAKE_MIDEND), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_end, rowcnt)
            '状況出力
            tmp_partsituation = SITUATION_MIDEND
            Call Me.Set_Situation(tmp_partsituation, 1)

            Return normalflg

        End Function

        ''' <summary>
        ''' 中間ファイル作成処理
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="sheetname"></param>
        ''' <param name="qry_view"></param>
        ''' <param name="sortstr"></param>
        ''' <param name="makecnt"></param>
        ''' <param name="notmakecnt"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Make_MidFile(ByVal filename As String, _
                                      ByVal sheetname As String, _
                                      ByVal qry_view As String, _
                                      ByVal sortstr As String, _
                                      ByRef makecnt As Integer, _
                                      ByRef notmakecnt As Integer, _
                                      ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim rowcnt As Integer = 0
            Dim totalrowcnt As Integer = 0
            Dim cvrowcnt As Integer = 0
            Dim tmp_totalsituation As String = ""
            Dim tmp_partsituation As String = ""
            Dim tmp_str As String = ""
            Dim obj_pgb As New ProgressBarManager
            Dim pgbcnt As Integer = 0


            '----- 個別用プログレスバー初期化 -----
            pgbpartcnt = 4                                  '4行程で中間ファイルを作成
            Call obj_pgb.pgbInitPart(pgbpartcnt)
            Call obj_com.ProgressOutPut(0, pgbpartcnt)


            '--------------------------------------------------
            ' 移行元を成形したVIEWの作成
            '--------------------------------------------------
            '書込先中間ファイルパスを取得
            Dim midfilepath As String = EtcMethod.Set_Path(MiddleDirPath, filename) & ".xlsx"
            '作成するVIEWの名称設定
            Dim viewname As String = PRE_VIEW_NAME & sheetname

            '出力ファイルパス設定 (拡張子が「.csv」の場合、展開時に列定義が不可になるため「.tmp」にしておく)
            Dim tmp_strpath As String = EtcMethod.Set_Path(MiddleDirPath, DIR_TEMP_CSV)
            Dim csvfilepath As String = EtcMethod.Set_Path(tmp_strpath, sheetname) & ".tmp"

            '既存VIEW削除                                                               
            Dim tmp_sqldrop As String = DBQuery.Qry_DropInfo(viewname, False)
            rtn = DBExec.Exec_NonQuery(sqlcnnv7, tmp_sqldrop, rowcnt)
            If rtn = False Then
                Return rtn
            End If

            '新規VIEW作成
            Dim tmp_sqlview As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY & qry_view
            rtn = DBExec.Exec_NonQuery(sqlcnnv7, tmp_sqlview, rowcnt)
            If rtn = False Then
                errstr = SITUATION_MIDERREND
                Return rtn
            End If

            '作成したVIEWから列数を取得(CSV展開時の列定義用)
            Dim tmp_sqlgetcolcnt As String = DBQuery.Qry_GetColCount(viewname)
            Dim colcnt As Integer = DBExec.Exec_Scalar(tmp_sqlgetcolcnt, sqlcnnv7)

            '作成したVIEWから行数を取得(状況出力用)
            Dim tmp_sqlgetrowcnt As String = DBQuery.Qry_GetRowCount(viewname)
            totalrowcnt = DBExec.Exec_Scalar(tmp_sqlgetrowcnt, sqlcnnv7)
            tmp_str = tmp_str & vbCrLf & INDENT_2 & SITUATION_CVITEMTOTALCNT & INDENT_1 & totalrowcnt.ToString

            '----- 個別用プログレスバー更新 -----                               
            pgbcnt = pgbcnt + 1
            Call obj_pgb.pgbsettingPart(pgbcnt)
            Call obj_com.ProgressOutPut(pgbcnt, pgbpartcnt)
            Me.tabPageJikko.Refresh()


            '--------------------------------------------------
            ' 作成したVIEWからCSVファイルを作成
            '--------------------------------------------------
            'CSV出力                                                                   
            rtn = FileMethod.TblView_Output_CSV(sqlcnnv7, csvfilepath, viewname, sortstr, errstr)
            If rtn = False Then
                Return rtn
            End If

            '----- 個別用プログレスバー更新 -----
            pgbcnt = pgbcnt + 1
            Call obj_pgb.pgbsettingPart(pgbcnt)
            Call obj_com.ProgressOutPut(pgbcnt, pgbpartcnt)
            Me.tabPageJikko.Refresh()


            '--------------------------------------------------
            ' 作成したCSVを中間ファイルへ書込み
            '--------------------------------------------------
            '書込処理(データが0件の場合は書込み処理を行わない (Excelファイルにゴミデータが混入するため))
            If totalrowcnt <> 0 Then
                rtn = FileMethod.CSV_To_Excel(csvfilepath, midfilepath, sheetname, colcnt, cvrowcnt)
                If rtn = False Then
                    Return rtn
                End If
            End If

            '----- 個別用プログレスバー更新 -----
            pgbcnt = pgbcnt + 1
            Call obj_pgb.pgbsettingPart(pgbcnt)
            Call obj_com.ProgressOutPut(pgbcnt, pgbpartcnt)
            Me.tabPageJikko.Refresh()


            '--------------------------------------------------
            ' 作成VIEWの削除(開発用タブ内のチェックで制御)
            '--------------------------------------------------
            'VIEW削除
            Dim viewdropflg As Boolean = Me.chkV7ViewDrop.Checked
            If viewdropflg Then
                rtn = DBExec.Exec_NonQuery(sqlcnnv7, tmp_sqldrop, rowcnt)
                If Not rtn Then
                    Return rtn
                End If
            End If

            '----- 個別用プログレスバー更新 -----
            pgbcnt = pgbcnt + 1
            Call obj_pgb.pgbsettingPart(pgbcnt)
            Call obj_com.ProgressOutPut(pgbcnt, pgbpartcnt)
            Me.tabPageJikko.Refresh()


            '--------------------------------------------------
            '件数格納
            '--------------------------------------------------
            '作成件数格納
            makecnt = cvrowcnt

            '未作成件数格納
            notmakecnt = totalrowcnt - cvrowcnt

            Return rtn

        End Function

        ''' <summary>
        ''' 各移行項目参考件数表示処理 '20160707 各抽出件数の出力処理追加 20160905 汎用コンバートの件数表示修正(メソッド名Set_CVItemCnt→Set_ExistCVItemCnt)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_ExistCVItemCnt()

            Dim tblname As String = ""                  '抽出テーブル名格納用
            Dim maxhysetubino As Integer = 77           '部屋設備の最大項目数
            Dim list_cvitem As New List(Of String)      '項目選択1～3内の全チェックボックス名称格納用
            list_cvitem = Me.Get_ListCVChkitemAll()

            '----- 中間用プログレスバー更新初期化 -----
            Dim obj_pgb As New ProgressBarManager
            Dim pgbcnt As Integer = 0
            Dim pgbtotalcnt As Integer = list_cvitem.Count
            Call obj_pgb.pgbInitChkTotal(pgbtotalcnt)
            Me.lblCheckSituation.Text = SITUATION_EXTRACTION
            Me.lblPgbCheck.Text = "0 %"
            Me.pgbCheck.Value = 0
            Me.pnlPrgChk.Visible = True
            Me.tabCtrlMain.Enabled = False
        

            For Each cvitem In list_cvitem

                Dim tmp_sql_exec As String = ""         '実行時のクエリ格納用
                Dim tmp_sql_other As String = ""        '単純に抽出できない場合のクエリ格納用
                Dim outputcnt As String = ""            '出力する件数(文字列)
                Dim lbl_output As New Label             '出力するラベル格納用

                '各移行項目毎処理
                Select Case cvitem
                    Case "バス交通マスタ"
                        lbl_output = Me.lblKiMstBusCnt
                        tblname = "m_bus"
                    Case "学校区マスタ"
                        lbl_output = Me.lblKiMstSchoolCnt
                        tmp_sql_other = " SELECT COUNT(*) FROM (SELECT DISTINCT syogaku_name,cyugaku_name FROM bk_mst WHERE syogaku_name <> '' OR cyugaku_name <> '') AS VW "
                    Case "エリアマスタ"
                        lbl_output = Me.lblKiMstAreaCnt
                        tblname = "m_addkbn"
                    Case "保険種類マスタ"
                        lbl_output = Me.lblKiMstHokenruiCnt
                        tblname = "m_hoken_rui"
                    Case "特約マスタ"
                        lbl_output = Me.lblKiMstTokuyakuCnt
                        tmp_sql_other = " SELECT (SELECT COUNT(*) FROM m_tokuyaku) + (SELECT COUNT(*) FROM m_tokuyaku_genjo) "
                    Case "クレーム分類設定内容"
                        lbl_output = Me.lblKiMstKasyoClaimruiCnt
                        tmp_sql_other = " SELECT COUNT(*) FROM (SELECT DISTINCT kasho_rui FROM claim_data UNION SELECT DISTINCT claim_rui FROM claim_data) AS VW "
                    Case "変動費設定内容"
                        lbl_output = Me.lblKiMstHendoCnt
                        tblname = "m_hendo"
                        '20160818 タイトルマスタ件数表示修正 -chg sta
                        'Case "鍵タイトルマスタ"
                        '    lbl_output = Me.lblKiMstKagititleCnt
                        '    tmp_sql_other = " SELECT COUNT(*) FROM (SELECT * FROM m_biko WHERE biko_kbn = 4) AS VW "
                        'Case "備考タイトルマスタ"
                        '    lbl_output = Me.lblKiMstBikotitleCnt
                        '    tmp_sql_other = " SELECT (SELECT COUNT(*) FROM m_biko WHERE biko_kbn <> 4) + (SELECT COUNT(*) FROM m_jyuyo WHERE jyuyo_kbn = 4) "
                    Case "タイトルマスタ"
                        Call Me.Set_CVItemCnt_TitleMst()
                        '20160818 タイトルマスタ件数表示修正 -chg end
                    Case "初期設定"
                        lbl_output = Me.lblKiSyskanriBaseCnt
                        tblname = "m_kan"
                    Case "自社情報"
                        lbl_output = Me.lblKiJisyaBaseCnt
                        tmp_sql_other = " SELECT (SELECT COUNT(*) FROM m_jisya) + (SELECT COUNT(*) FROM m_siten) "
                    Case "家主情報"
                        lbl_output = Me.lblKiOwBaseCnt
                        tblname = "m_yanu_so"
                    Case "仲介・管理業者情報"
                        lbl_output = Me.lblKiGyCyukaiBaseCnt
                        tblname = "m_gy"
                    Case "修繕業者情報"
                        lbl_output = Me.lblKiGySyuzenBaseCnt
                        tblname = "m_gysyuzen"
                    Case "ライフライン業者情報"
                        lbl_output = Me.lblKiGyLifelineBaseCnt
                        tblname = "m_kokyo"
                    Case "保険業者情報"
                        lbl_output = Me.lblKiGyHokenBaseCnt
                        tblname = "m_hoken_gy"
                    Case "家賃保証業者情報"
                        lbl_output = Me.lblKiGyYatinhosyoBaseCnt
                        tblname = "m_hosyo_gy"
                    Case "施設保守業者情報"
                        lbl_output = Me.lblKiGySisetuBaseCnt
                        tblname = "m_hosyu_kaisya"
                    Case "施工業者情報"
                        lbl_output = Me.lblKiGySekoBaseCnt
                        tblname = "m_seko_kaisya"
                    Case "物件情報"
                        lbl_output = Me.lblKiBkBaseCnt
                        tblname = "bk_mst"
                    Case "部屋情報"
                        lbl_output = Me.lblKiHyBaseCnt
                        tblname = "hy_mst"
                    Case "部屋設備情報"
                        lbl_output = Me.lblKiHySetubiCnt
                        For cntii = 1 To maxhysetubino
                            Dim tmp_sql As String = ""
                            If cntii < maxhysetubino Then
                                tmp_sql = " (SELECT COUNT(*) FROM (SELECT bk_no,hy_no,setubi" & cntii.ToString & " FROM hy_setubi WHERE setubi" & cntii.ToString & " <> '') AS VW) + "
                            Else
                                tmp_sql = " (SELECT COUNT(*) FROM (SELECT bk_no,hy_no,setubi" & cntii.ToString & " FROM hy_setubi WHERE setubi" & cntii.ToString & " <> '') AS VW) "
                            End If
                            tmp_sql_other = tmp_sql_other & tmp_sql
                        Next
                        tmp_sql_other = " SELECT " & tmp_sql_other
                    Case "契約者情報"
                        lbl_output = Me.lblKiKysBaseCnt
                        tblname = "kys_mst"
                    Case "契約情報"
                        lbl_output = Me.lblKiKyBaseCnt
                        tblname = "ky_kosinkai"
                    Case "請求情報"
                        Call Me.Set_CVItemCnt_Sqdata()
                    Case "クレーム情報"
                        lbl_output = Me.lblKiClaimBaseCnt
                        tblname = "claim_data"
                    Case "修繕情報"
                        lbl_output = Me.lblKiSzenBaseCnt
                        tblname = "reform_data"
                    Case "ポータル連動情報"             '20160905 既存コンバートの連動件数表示 -add
                        lbl_output = Me.lblKiRendoBaseCnt
                        Dim tmptblname As String = PRE_TBL_NAME & "物件データ連動情報"
                        Call Me.Make_RendoTmpTable(tmptblname)
                        tmp_sql_other = " SELECT COUNT(*) FROM CVTBL_物件データ連動情報 WHERE [Enabled] <> 0 "
                End Select

                'クエリ格納
                If tblname <> "" And tmp_sql_other = "" Then
                    tmp_sql_exec = " SELECT COUNT(*) FROM " & tblname
                Else
                    tmp_sql_exec = tmp_sql_other
                End If

                '件数取得実行
                If tmp_sql_exec <> "" Then
                    outputcnt = DBExec.Exec_Scalar(tmp_sql_exec, sqlcnnv7)
                    lbl_output.Text = Int32.Parse(outputcnt).ToString("#,0") & " 件"
                Else
                    lbl_output.Text = ""
                End If

                '----- 中間用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsettingChkTotal(pgbcnt)
                Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt, True)
                Me.Refresh()

            Next

            '----- 中間用プログレスバー非表示 -----
            Me.pnlPrgChk.Visible = False
            Me.tabCtrlMain.Enabled = True

        End Sub

        ''' <summary>
        ''' 請求情報参考件数表示処理 '20160707 各抽出件数の出力処理追加
        ''' </summary>
        ''' <remarks>
        ''' ・請求情報は未収・滞納金と預り金の2つに分けて表示する
        ''' </remarks>
        Private Sub Set_CVItemCnt_Sqdata()

            Dim outputcnt As String = "0"

            '預り金
            Dim tmp_sql_azu As String = "SELECT COUNT(*) FROM sq_meisai WHERE nkin_no = 6020"
            outputcnt = DBExec.Exec_Scalar(tmp_sql_azu, sqlcnnv7)
            lblKiSqAzBaseCnt.Text = Int32.Parse(outputcnt).ToString("#,0") & " 件"
            outputcnt = "0"

            '未収・滞納金
            Dim tmp_sql_misyu As String = ""
            tmp_sql_misyu = tmp_sql_misyu & " SELECT "
            tmp_sql_misyu = tmp_sql_misyu & " /*未収件数 sta*/ "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*その他請求の移行*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sq_meisai AS SQ "
            tmp_sql_misyu = tmp_sql_misyu & " 	WHERE nkin_ymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   data_kbn = 6 "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.nkin_no <> 6020			/*入金項目No:6020 は過剰金*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   EXISTS						/*該当年月が適用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 			( "
            tmp_sql_misyu = tmp_sql_misyu & " 				SELECT * FROM bk_kanri AS BKK "
            tmp_sql_misyu = tmp_sql_misyu & " 				WHERE SQ.bk_no = BKK.bk_no "
            tmp_sql_misyu = tmp_sql_misyu & " 				AND   SQ.gt_ym >= BKK.start_ym "
            tmp_sql_misyu = tmp_sql_misyu & " 			) "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   so_kakuymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " + "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*部分入金の移行*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sq_meisai AS SQ "
            tmp_sql_misyu = tmp_sql_misyu & " 	WHERE rec_kbn <> 0 "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   nkin_ymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   data_kbn <> 6 "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   EXISTS					/*該当年月が適用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 			( "
            tmp_sql_misyu = tmp_sql_misyu & " 				SELECT * FROM bk_kanri AS BKK "
            tmp_sql_misyu = tmp_sql_misyu & " 				WHERE SQ.bk_no = BKK.bk_no "
            tmp_sql_misyu = tmp_sql_misyu & " 				AND   SQ.gt_ym >= BKK.start_ym "
            tmp_sql_misyu = tmp_sql_misyu & " 			) "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   so_kakuymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " + "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*未入金かつ送金確定されているデータの移行*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*その他請求以外の送金確定*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sq_meisai AS SQ "
            tmp_sql_misyu = tmp_sql_misyu & " 	WHERE nkin_ymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   data_kbn <> 6 "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   EXISTS					/*該当年月が適用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 			( "
            tmp_sql_misyu = tmp_sql_misyu & " 				SELECT * FROM bk_kanri AS BKK "
            tmp_sql_misyu = tmp_sql_misyu & " 				WHERE SQ.bk_no = BKK.bk_no "
            tmp_sql_misyu = tmp_sql_misyu & " 				AND   SQ.gt_ym >= BKK.start_ym "
            tmp_sql_misyu = tmp_sql_misyu & " 			) "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   so_kakuymd IS NOT NULL "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " + "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*未入金かつ送金確定されているデータの移行*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*その他請求の送金確定*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sq_meisai AS SQ "
            tmp_sql_misyu = tmp_sql_misyu & " 	WHERE nkin_ymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   data_kbn = 6 "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.nkin_no <> 6020		/*入金項目No:6020 は過剰金*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   EXISTS					/*該当年月が適用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 			( "
            tmp_sql_misyu = tmp_sql_misyu & " 				SELECT * FROM bk_kanri AS BKK "
            tmp_sql_misyu = tmp_sql_misyu & " 				WHERE SQ.bk_no = BKK.bk_no "
            tmp_sql_misyu = tmp_sql_misyu & " 				AND   SQ.gt_ym >= BKK.start_ym "
            tmp_sql_misyu = tmp_sql_misyu & " 			) "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   so_kakuymd IS NOT NULL "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " + "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*控除支払登録データから請求のフラグが立っているデータを抽出*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sh_kojo AS SHK WHERE SHK.sq_umu = 1 "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " + "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*控除支払登録データから支払に金額が設定されているデータの抽出→自社支払いへ移行する*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sh_kojo AS SHK WHERE siharai_gak <> 0 AND SHK.nkin_no <> 9910 "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " /*未収件数 end*/ "
            tmp_sql_misyu = tmp_sql_misyu & " /*滞納件数 sta*/ "
            tmp_sql_misyu = tmp_sql_misyu & " + "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*過去送金実績が無い場合*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sq_meisai AS SQ "
            tmp_sql_misyu = tmp_sql_misyu & " 	WHERE nkin_ymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.nkin_no <> 6020 "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   so_kakuymd IS NULL			/*送金実績が無い*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   NOT "
            tmp_sql_misyu = tmp_sql_misyu & " 			( "
            tmp_sql_misyu = tmp_sql_misyu & " 				SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/			 "
            tmp_sql_misyu = tmp_sql_misyu & " 				AND "
            tmp_sql_misyu = tmp_sql_misyu & " 				EXISTS					/*該当年月が適用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 				( "
            tmp_sql_misyu = tmp_sql_misyu & " 					SELECT * FROM bk_kanri AS BKK "
            tmp_sql_misyu = tmp_sql_misyu & " 					WHERE SQ.bk_no = BKK.bk_no "
            tmp_sql_misyu = tmp_sql_misyu & " 					AND   SQ.gt_ym >= BKK.start_ym				 "
            tmp_sql_misyu = tmp_sql_misyu & " 				) "
            tmp_sql_misyu = tmp_sql_misyu & " 			) "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " + "
            tmp_sql_misyu = tmp_sql_misyu & " ( "
            tmp_sql_misyu = tmp_sql_misyu & " 	/*過去送金実績が有る場合*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	SELECT COUNT(*) FROM sq_meisai AS SQ "
            tmp_sql_misyu = tmp_sql_misyu & " 	WHERE nkin_ymd IS NULL "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   SQ.nkin_no <> 6020 "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   so_kakuymd IS NOT NULL		/*送金実績が有る*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 	AND   NOT "
            tmp_sql_misyu = tmp_sql_misyu & " 			( "
            tmp_sql_misyu = tmp_sql_misyu & " 				SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/			 "
            tmp_sql_misyu = tmp_sql_misyu & " 				AND "
            tmp_sql_misyu = tmp_sql_misyu & " 				EXISTS					/*該当年月が適用開始年月以降*/ "
            tmp_sql_misyu = tmp_sql_misyu & " 				( "
            tmp_sql_misyu = tmp_sql_misyu & " 					SELECT * FROM bk_kanri AS BKK "
            tmp_sql_misyu = tmp_sql_misyu & " 					WHERE SQ.bk_no = BKK.bk_no "
            tmp_sql_misyu = tmp_sql_misyu & " 					AND   SQ.gt_ym >= BKK.start_ym				 "
            tmp_sql_misyu = tmp_sql_misyu & " 				) "
            tmp_sql_misyu = tmp_sql_misyu & " 			) "
            tmp_sql_misyu = tmp_sql_misyu & " ) "
            tmp_sql_misyu = tmp_sql_misyu & " /*滞納件数 end*/ "

            Dim tmp_unyoymd As String = Me.txtUnyoYYYYMM.Text & "/01"
            tmp_sql_misyu = tmp_sql_misyu.Replace("運用開始年月置換用文字列", tmp_unyoymd)
            outputcnt = DBExec.Exec_Scalar(tmp_sql_misyu, sqlcnnv7)
            lblKiSqMiBaseCnt.Text = Int32.Parse(outputcnt).ToString("#,0") & " 件"
            outputcnt = "0"

        End Sub

        ''' <summary>
        ''' タイトルマスタ参考件数表示処理 '20160818 タイトルマスタ件数表示修正
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_CVItemCnt_TitleMst()

            Dim tmp_sql As String = ""
            Dim outputcnt As String = ""            '出力する件数(文字列)

            '鍵タイトルマスタ
            tmp_sql = " SELECT COUNT(*) FROM (SELECT * FROM m_biko WHERE biko_kbn = 4) AS VW "
            outputcnt = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            If outputcnt <> "0" Then
                lblKiMstKagititleCnt.Text = Int32.Parse(outputcnt).ToString("#,0") & " 件"
            End If

            '備考タイトルマスタ
            tmp_sql = " SELECT (SELECT COUNT(*) FROM m_biko WHERE biko_kbn <> 4) + (SELECT COUNT(*) FROM m_jyuyo WHERE jyuyo_kbn = 4) "
            outputcnt = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            If outputcnt <> "0" Then
                lblKiMstBikotitleCnt.Text = Int32.Parse(outputcnt).ToString("#,0") & " 件"
            End If

            '画像タイトルマスタ
            tmp_sql = " SELECT COUNT(*) FROM m_gazo_syoki WHERE gazo_kbn IN (1,2,3,5) AND gazo_name <> '' "
            outputcnt = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            If outputcnt <> "0" Then
                lblKiMstGazotitleCnt.Text = Int32.Parse(outputcnt).ToString("#,0") & " 件"
            End If

        End Sub

        '20160707 V7用処理へ移動 -add sta
        ''' <summary>
        ''' 移行項目毎の個別処理 '20160829 連動中間ファイル作成処理修正 引数に処理項目名を追加
        ''' </summary>
        ''' <param name="syorigrp"></param>
        ''' <remarks></remarks>
        Private Sub Setting_BeforeMakeMidFile(ByVal syorigrp As String, ByVal syorikomkname As String)

            Select Case CNVNO
                Case ConvertTypes._汎用

                Case ConvertTypes._既存ユーザ用
                    Select Case syorigrp
                        Case "契約情報"
                            Dim tmptblname As String = PRE_TBL_NAME & syorigrp
                            Call Me.Make_KytmpTable(tmptblname)
                        Case "物件データ連動情報" '20160720 連動情報構築 -add
                            Dim tmptblname As String = PRE_TBL_NAME & syorigrp
                            Call Me.Make_RendoTmpTable(tmptblname)
                            '20160829 連動中間ファイル作成処理修正 -add sta
                            If syorikomkname = "送信設定athome情報" Then
                                Call Me.Make_SubView_NotGazo(syorikomkname)
                                Call Me.Make_SubView_OnlyGazo(syorikomkname)
                            End If
                            '20160829 連動中間ファイル作成処理修正 -add end
                    End Select
                Case 2
            End Select

        End Sub

        ''' <summary>
        ''' V7の契約情報テーブルを10の契約情報テーブルをへ合わせる仮テーブルを作成　　　　
        ''' </summary>
        ''' <param name="tmptblname"></param>
        ''' <remarks>
        ''' ・契約情報中間ファイル作成用の作業用テーブル作成
        ''' </remarks>
        Private Sub Make_KytmpTable(ByVal tmptblname As String)

            Dim tmpcnt As Integer = 0

            '初期化
            Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(tmptblname, True)
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sql_drop, tmpcnt)

            '仮テーブル作成クエリ
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " /*変数宣言*/ "
            tmp_sql = tmp_sql & " DECLARE @bkno INT "
            tmp_sql = tmp_sql & " DECLARE @hyno VARCHAR(8) "
            tmp_sql = tmp_sql & " DECLARE @kyno INT "
            tmp_sql = tmp_sql & " DECLARE @kono INT "
            tmp_sql = tmp_sql & " DECLARE @kokbn INT "
            tmp_sql = tmp_sql & " DECLARE @kono10 INT "
            tmp_sql = tmp_sql & " DECLARE @henkono10 INT "
            tmp_sql = tmp_sql & " DECLARE @tmpoldkey VARCHAR(50) "
            tmp_sql = tmp_sql & " DECLARE @tmpoldkono INT "
            tmp_sql = tmp_sql & " DECLARE @tmpoldkokbn INT "
            tmp_sql = tmp_sql & " DECLARE @tmpoldhenkono INT "
            tmp_sql = tmp_sql & " DECLARE @tmpnewkey VARCHAR(50) "
            tmp_sql = tmp_sql & " DECLARE @tmpnewkono INT "
            tmp_sql = tmp_sql & " DECLARE @tmpnewkokbn INT "
            tmp_sql = tmp_sql & " DECLARE @tmpnewhenkono INT "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " /*カーソル宣言*/ "
            tmp_sql = tmp_sql & " DECLARE kycursor CURSOR FOR "
            tmp_sql = tmp_sql & " 	SELECT bk_no,hy_no,ky_no,ko_no,ko_kbn FROM ky_kosinkai "
            tmp_sql = tmp_sql & " 	ORDER BY bk_no,hy_no,ky_no,ko_no "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " /*仮テーブル作成*/ "
            tmp_sql = tmp_sql & " CREATE TABLE CVTBL_契約情報 "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		 tmp_bkno INT "
            tmp_sql = tmp_sql & " 		,tmp_hyno VARCHAR(8) "
            tmp_sql = tmp_sql & " 		,tmp_kyno INT "
            tmp_sql = tmp_sql & " 		,tmp_kono INT "
            tmp_sql = tmp_sql & " 		,tmp_kokbn INT "
            tmp_sql = tmp_sql & " 		,kono10 INT "
            tmp_sql = tmp_sql & " 		,henkono10 INT "
            tmp_sql = tmp_sql & " 	) "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " /*開始処理*/ "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " SET NOCOUNT ON "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " OPEN kycursor "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " SET @tmpoldkey = '' "
            tmp_sql = tmp_sql & " SET @tmpnewkey = '' "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " FETCH NEXT FROM kycursor INTO @bkno,@hyno,@kyno,@kono,@kokbn "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " WHILE @@FETCH_STATUS = 0 "
            tmp_sql = tmp_sql & " BEGIN "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*現在行のキー(同一契約を示す部分)を取得*/ "
            tmp_sql = tmp_sql & " 	SET @tmpnewkey = CONVERT(VARCHAR,@bkno) + '-' + CONVERT(VARCHAR,@hyno) + '-' + CONVERT(VARCHAR,@kyno) "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " 	/*前のデータと照合*/ "
            tmp_sql = tmp_sql & " 	IF @tmpoldkey <> @tmpnewkey		/*契約情報が異なっている場合*/ "
            tmp_sql = tmp_sql & " 		BEGIN "
            tmp_sql = tmp_sql & " 			SET @kono10 = 1			/*更新No初期化*/ "
            tmp_sql = tmp_sql & " 			SET @henkono10 = 1		/*変更No初期化*/ "
            tmp_sql = tmp_sql & " 		END "
            tmp_sql = tmp_sql & " 	ELSE IF @kokbn = 1				/*契約情報が同一で更新されている場合*/ "
            tmp_sql = tmp_sql & " 		BEGIN "
            tmp_sql = tmp_sql & " 			SET @tmpnewkono = @tmpoldkono + 1	/*更新Noをカウントアップしてセット*/ "
            tmp_sql = tmp_sql & " 			SET @kono10 = @tmpnewkono "
            tmp_sql = tmp_sql & " 			SET @henkono10 = 1					/*変更Noを初期化*/ "
            tmp_sql = tmp_sql & " 		END "
            tmp_sql = tmp_sql & " 	ELSE IF @kokbn = 2				/*契約情報が同一で改定されている場合*/ "
            tmp_sql = tmp_sql & " 		BEGIN "
            tmp_sql = tmp_sql & " 			SET @kono10 = @tmpoldkono					/*更新Noを引き継ぎ*/ "
            tmp_sql = tmp_sql & " 			SET @tmpnewhenkono = @tmpoldhenkono + 1		/*変更Noをカウントアップしてセット*/ "
            tmp_sql = tmp_sql & " 			SET @henkono10 = @tmpnewhenkono "
            tmp_sql = tmp_sql & " 		END "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*挿入処理*/ "
            tmp_sql = tmp_sql & " 	INSERT INTO CVTBL_契約情報 VALUES (@bkno,@hyno,@kyno,@kono,@kokbn,@kono10,@henkono10) "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " 	/*次の行との照合用に値を退避させておく*/ "
            tmp_sql = tmp_sql & " 	SET @tmpoldkey = CONVERT(VARCHAR,@bkno) + '-' + CONVERT(VARCHAR,@hyno) + '-' + CONVERT(VARCHAR,@kyno) "
            tmp_sql = tmp_sql & " 	SET @tmpoldkono = @kono10 "
            tmp_sql = tmp_sql & " 	SET @tmpoldhenkono = @henkono10 "
            tmp_sql = tmp_sql & " 	SET @tmpoldkokbn = @kokbn "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " 	FETCH NEXT FROM kycursor INTO @bkno,@hyno,@kyno,@kono,@kokbn "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " END "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " /*終了処理*/ "
            tmp_sql = tmp_sql & " CLOSE kycursor "
            tmp_sql = tmp_sql & " DEALLOCATE kycursor "

            '仮テーブル作成実行
            Dim normalflg As Boolean = DBExec.Exec_NonQuery(sqlcnnv7, tmp_sql, tmpcnt)

        End Sub
        '20160707 V7用処理へ移動 -add end

        ''' <summary>
        ''' レジストリに書き込まれているV7連動送信設定情報から仮テーブルを作成 '20160720 連動情報構築 -add
        ''' </summary>
        ''' <param name="tmptblname"></param>
        ''' <remarks></remarks>
        Private Sub Make_RendoTmpTable(ByVal tmptblname As String)

            Dim setno As Integer = 10   '送信データベース数
            Dim itemno As Integer = 11  '要素数
            Dim tmp_enabled(setno) As String
            Dim tmp_system(setno) As String
            Dim tmp_regSet(setno) As String
            Dim tmp_server(setno) As String
            Dim tmp_catalog(setno) As String
            Dim tmp_user(setno) As String
            Dim tmp_pass(setno) As String
            Dim tmp_imgPath(setno) As String
            Dim tmp_comment(setno) As String
            Dim tmp_timer(setno) As String
            Dim tmp_backup(setno) As String

            Dim tmp_regitemprestr As String = "SET_DB"
            Dim tmp_regitempoststr() As String = {"", "Enabled", "System", "RegSet", "Server", "Catalog", "User", "Pass", "ImgPath", "Comment", "Timer", "Backup"}


            '-------------------------------------------------------
            '送信設定基本情報格納用仮テーブル作成
            '-------------------------------------------------------
            '初期化
            Dim tmp_cnt As Integer = 0
            Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(tmptblname, True)
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sql_drop, tmp_cnt)

            '作成
            Dim tmp_sql_create As String = ""
            tmp_sql_create = tmp_sql_create & " CREATE TABLE " & tmptblname
            tmp_sql_create = tmp_sql_create & " 	( "
            tmp_sql_create = tmp_sql_create & " 		[No] INT, "
            tmp_sql_create = tmp_sql_create & " 		[Enabled] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[System] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[RegSet] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[Server] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[Catalog] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[User] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[Pass] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[ImgPath] NCHAR(100), "
            tmp_sql_create = tmp_sql_create & " 		[Comment] NCHAR(500), "
            tmp_sql_create = tmp_sql_create & " 		[Timer] NCHAR(100),"
            tmp_sql_create = tmp_sql_create & " 		[Backup] NCHAR(100) "
            tmp_sql_create = tmp_sql_create & " 	) "
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sql_create, tmp_cnt)


            '-------------------------------------------------------
            'V7送信設定情報が格納されているレジストリパス設定
            '-------------------------------------------------------
            Dim regpath_rendomain As String = "Software\VB and VBA Program Settings\NJCAgent\NJCAgent"      'メイン
            Dim regpath_rendosub As String = "Software\VB and VBA Program Settings\Fkanri50\NJCAgent"       '古い分？


            '-------------------------------------------------------
            'V7送信設定情報をレジストリから取得
            '-------------------------------------------------------
            Dim regkeymain As RegistryKey
            Dim regkeysub As RegistryKey
            Dim regkeycvuse As RegistryKey
            regkeymain = Registry.CurrentUser.OpenSubKey(regpath_rendomain, False)
            regkeysub = Registry.CurrentUser.OpenSubKey(regpath_rendosub, False)

            If regkeymain IsNot Nothing Then
                'メインにデータが存在する場合はメインを使う
                regkeycvuse = regkeymain
            ElseIf regkeysub IsNot Nothing Then
                'メインにデータが存在せず古い分？に存在する場合は古い分？を使う
                regkeycvuse = regkeysub
            Else
                'どちらも存在しない場合は送信設定情報無しとして移行対象外
                Exit Sub
            End If


            '-------------------------------------------------------
            '取得したレジストリから値を取得
            '-------------------------------------------------------
            For cntii = 1 To setno

                Dim tmp_value As String = ""

                '各要素の値を取得→連結
                For cntjj = 1 To itemno

                    Dim tmp_itemname As String = tmp_regitemprestr & cntii.ToString & tmp_regitempoststr(cntjj)
                    tmp_value = tmp_value & "," & "'" & DirectCast(regkeycvuse.GetValue(tmp_itemname, ""), String) & "'"
                Next

                tmp_value = cntii.ToString & tmp_value

                '挿入処理
                Dim tmp_sql_insert As String = " INSERT INTO " & tmptblname & " VALUES(" & tmp_value & ")"
                DBExec.Exec_NonQuery(sqlcnnv7, tmp_sql_insert, tmp_cnt)

            Next

        End Sub

        ''' <summary>
        ''' 送信設定athome情報の画像以外の項目VIEW化 '20160829 連動中間ファイル作成処理修正 -add
        ''' </summary>
        ''' <param name="syorikomkname"></param>
        ''' <remarks></remarks>
        Public Sub Make_SubView_NotGazo(ByVal syorikomkname As String)

            Dim tmp_sql As String = ""

            For cntii = 0 To 9

                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & "        送信設定No置換文字列 AS [送信設定順]  "
                tmp_sql = tmp_sql & " 		,30 AS [サイトNo] "
                tmp_sql = tmp_sql & " 		/*******************************************フィールドへ移行するデータ sta*******************************************/ "
                tmp_sql = tmp_sql & " 		,(SELECT REPLACE(set_value1,RIGHT(set_value1,4),'') FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061002') AS [ポータルサイトID] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01061001','11061001','21061001','31061001','41061001','51061001','61061001','71061001','81061001','91061001')) = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01061001','11061001','21061001','31061001','41061001','51061001','61061001','71061001','81061001','91061001')) = 1 THEN 0 "
                tmp_sql = tmp_sql & " 		 END AS [サイト別送信有無] "
                tmp_sql = tmp_sql & " 		,'' AS [ポータルサイトパスワード] "
                tmp_sql = tmp_sql & " 		/*******************************************フィールドへ移行するデータ end*******************************************/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/************************************************xmlへ移行するデータ sta************************************************/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/***********表示設定タブ***********/ "
                tmp_sql = tmp_sql & " 		,(SELECT RIGHT(set_value1,4) FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061002') AS [店舗ID] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 3 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 1 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 1 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 2 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 3 THEN 4 "
                tmp_sql = tmp_sql & " 		 END AS [物件名・部屋Noの表示(一般ユーザー向け)] "
                tmp_sql = tmp_sql & "  		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 3 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 1 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 1 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 2 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 3 THEN 4 "
                tmp_sql = tmp_sql & " 		 END AS [物件名・部屋Noの表示(不動産会社向け)]		 "
                tmp_sql = tmp_sql & " 		 ,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061011') AS [地図表示] "
                tmp_sql = tmp_sql & " 		,2 AS [番地以降の表示]		/*要デフォルト値確認*/ "
                tmp_sql = tmp_sql & " 		,1 AS [掲載確認日] "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/***********入金項目タブ***********/ "
                tmp_sql = tmp_sql & " 		/*,'_円' AS [賃料表示単位]*/		/*固定値のためコメントアウト*/ "
                tmp_sql = tmp_sql & " 		/*,'_円' AS [共益費表示単位]*/		/*固定値のためコメントアウト*/ "
                tmp_sql = tmp_sql & " 		/*,'_円' AS [管理費表示単位]*/		/*固定値のためコメントアウト*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061101') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061101') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [敷金表示単位] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061102') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061102') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [礼金表示単位] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061103') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061103') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [保証金表示単位] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061107') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061107') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [保証金償却表示単位] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061106') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061106') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [敷引金表示単位] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061108') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061108') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [更新料表示単位] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061109') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061109') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [仲介手数料表示単位] "
                tmp_sql = tmp_sql & " 		/*,'_円' AS [鍵交換代等]*/		/*固定値のためコメントアウト*/ "
                tmp_sql = tmp_sql & " 		/*,'_円' AS [雑費]*/			/*固定値のためコメントアウト*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061104') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061104') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [駐車場敷金表示単位] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061105') = 1 THEN '_円' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061105') = 2 THEN '_ヶ月' "
                tmp_sql = tmp_sql & " 		    ELSE '_円' "
                tmp_sql = tmp_sql & " 		 END AS [駐車場礼金表示単位]		 		 "
                tmp_sql = tmp_sql & " 		/*,'_円' AS [その他一時金]*/	/*固定値のためコメントアウト*/ "
                tmp_sql = tmp_sql & " 		/*,'_円' AS [その他月額費用]*/	/*固定値のためコメントアウト*/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061001')) AS [入金項目_賃料1] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_賃料1] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061002')) AS [入金項目_賃料2] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_賃料2] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061003')) AS [入金項目_賃料3] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_賃料3] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061004')) AS [入金項目_共益費1] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_共益費1] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061005')) AS [入金項目_共益費2] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_共益費2] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061006')) AS [入金項目_共益費3] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_共益費3] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061007')) AS [入金項目_管理費1] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_管理費1] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061008')) AS [入金項目_管理費2] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_管理費2] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061009')) AS [入金項目_管理費3] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_管理費3] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061013')) AS [入金項目_敷金] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_敷金] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061014')) AS [入金項目_礼金] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_礼金] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061015')) AS [入金項目_保証金] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_保証金] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061021')) AS [入金項目_保証金償却] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_保証金償却] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061020')) AS [入金項目_敷引金] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_敷引金] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061030')) AS [入金項目_更新料] "
                tmp_sql = tmp_sql & " 		,3 AS [入金項目区分_更新料] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061034')) AS [入金項目_仲介手数料] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_仲介手数料] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061033')) AS [入金項目_鍵交換代等] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_鍵交換代等]		 "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061010')) AS [入金項目_雑費1] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_雑費1] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061011')) AS [入金項目_雑費2] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_雑費2] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061012')) AS [入金項目_雑費3] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_雑費3]		 "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061018')) AS [入金項目_駐車場敷金] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_駐車場敷金] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061019')) AS [入金項目_駐車場礼金] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_駐車場礼金] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061022')) AS [入金項目_その他月額費用1] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用1] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061023')) AS [入金項目_その他月額費用2] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用2] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061024')) AS [入金項目_その他月額費用3] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用3] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061031')) AS [入金項目_その他月額費用4] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用4] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061032')) AS [入金項目_その他月額費用5] "
                tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用5] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061025')) AS [入金項目_その他一時金1] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金1] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061026')) AS [入金項目_その他一時金2] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金2] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061027')) AS [入金項目_その他一時金3] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金3] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061028')) AS [入金項目_その他一時金4] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金4] "
                tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061029')) AS [入金項目_その他一時金5] "
                tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金5]		 "
                tmp_sql = tmp_sql & " 		/************************************************xmlへ移行するデータ end************************************************/ "
                tmp_sql = tmp_sql & " ) AS VW "

                tmp_sql = tmp_sql.Replace("送信設定No置換文字列", (cntii + 1).ToString)
                tmp_sql = tmp_sql.Replace("V7送信設定値取得用置換文字列", cntii.ToString)

                If cntii < 9 Then
                    tmp_sql = tmp_sql & " UNION "
                End If

            Next

            '画像以外の項目でVIEWを作成
            Dim tmp_cnt As Integer = 0

            'VIEWの名称
            Dim viewname As String = PRE_VIEW_NAME & syorikomkname & "_画像以外"

            '既存VIEW削除                                                               
            Dim tmp_sqldrop As String = DBQuery.Qry_DropInfo(viewname, False)
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sqldrop, tmp_cnt)

            '新規VIEW作成
            Dim tmp_sqlview As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY & tmp_sql
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sqlview, tmp_cnt)

        End Sub

        ''' <summary>
        ''' 送信設定athome情報の画像項目のみをVIEW化 '20160829 連動中間ファイル作成処理修正 -add
        ''' </summary>
        ''' <param name="syorikomkname"></param>
        ''' <remarks></remarks>
        Public Sub Make_SubView_OnlyGazo(ByVal syorikomkname As String)

            Dim tmp_sql As String = ""

            For cntii = 0 To 9

                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & "        送信設定No置換文字列 AS [送信設定順]  "
                tmp_sql = tmp_sql & " 		,30 AS [サイトNo] "
                tmp_sql = tmp_sql & " 		/*******************************************フィールドへ移行するデータ sta*******************************************/ "
                tmp_sql = tmp_sql & " 		/***********画像設定タブ***********/ "
                tmp_sql = tmp_sql & " 		,1 AS [画像No1] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061001'),1) AS [革命側の画像種別1] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061001'),2)) AS [革命側の画像タイトルNo1] "
                tmp_sql = tmp_sql & " 		,1 AS [連動側の画像種別1]	/*固定*/ "
                tmp_sql = tmp_sql & " 		,2 AS [画像No2] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061002'),1) AS [革命側の画像種別2] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061002'),2)) AS [革命側の画像タイトルNo2] "
                tmp_sql = tmp_sql & " 		,3 AS [連動側の画像種別2]	/*固定*/ "
                tmp_sql = tmp_sql & " 		,3 AS [画像No3] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AS [革命側の画像種別3] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) AS [革命側の画像タイトルNo3] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別3] "
                tmp_sql = tmp_sql & " 		,4 AS [画像No4] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AS [革命側の画像種別4] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) AS [革命側の画像タイトルNo4] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別4] "
                tmp_sql = tmp_sql & " 		,5 AS [画像No5] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AS [革命側の画像種別5] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) AS [革命側の画像タイトルNo5] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別5] "
                tmp_sql = tmp_sql & " 		,6 AS [画像No6] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AS [革命側の画像種別6] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) AS [革命側の画像タイトルNo6] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別6]	 "
                tmp_sql = tmp_sql & " 		,7 AS [画像No7] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AS [革命側の画像種別7] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) AS [革命側の画像タイトルNo7] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別7]		 "
                tmp_sql = tmp_sql & " 		,8 AS [画像No8] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AS [革命側の画像種別8] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) AS [革命側の画像タイトルNo8] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別8] "
                tmp_sql = tmp_sql & " 		,9 AS [画像No9] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AS [革命側の画像種別9] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) AS [革命側の画像タイトルNo9] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別9] "
                tmp_sql = tmp_sql & " 		,10 AS [画像No10] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AS [革命側の画像種別10] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) AS [革命側の画像タイトルNo10] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別10] "
                tmp_sql = tmp_sql & " 		,11 AS [画像No11] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AS [革命側の画像種別11] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) AS [革命側の画像タイトルNo11] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別11] "
                tmp_sql = tmp_sql & " 		,12 AS [画像No12] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AS [革命側の画像種別12] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) AS [革命側の画像タイトルNo12] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別12] "
                tmp_sql = tmp_sql & " 		,13 AS [画像No13] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AS [革命側の画像種別13] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) AS [革命側の画像タイトルNo13] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別13] "
                tmp_sql = tmp_sql & " 		,14 AS [画像No14] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AS [革命側の画像種別14] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) AS [革命側の画像タイトルNo14] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別14] "
                tmp_sql = tmp_sql & " 		,15 AS [画像No15] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AS [革命側の画像種別15] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) AS [革命側の画像タイトルNo15] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別15] "
                tmp_sql = tmp_sql & " 		,16 AS [画像No16] "
                tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AS [革命側の画像種別16] "
                tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) AS [革命側の画像タイトルNo16] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 0 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 3 THEN 14 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 4 THEN 20 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 5 THEN 18 "
                tmp_sql = tmp_sql & "     		ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別16] "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		 /***********連動項目設定タブ***********/ "
                tmp_sql = tmp_sql & " 		  "
                tmp_sql = tmp_sql & " 		/************************************************xmlへ移行するデータ end************************************************/ "
                tmp_sql = tmp_sql & " ) AS VW "

                tmp_sql = tmp_sql.Replace("送信設定No置換文字列", (cntii + 1).ToString)
                tmp_sql = tmp_sql.Replace("V7送信設定値取得用置換文字列", cntii.ToString)

                If cntii < 9 Then
                    tmp_sql = tmp_sql & " UNION "
                End If

            Next

            '画像以外の項目でVIEWを作成
            Dim tmp_cnt As Integer = 0

            'VIEWの名称
            Dim viewname As String = PRE_VIEW_NAME & syorikomkname & "_画像"

            '既存VIEW削除                                                               
            Dim tmp_sqldrop As String = DBQuery.Qry_DropInfo(viewname, False)
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sqldrop, tmp_cnt)

            '新規VIEW作成
            Dim tmp_sqlview As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY & tmp_sql
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sqlview, tmp_cnt)

        End Sub

        ''' <summary>
        ''' 移行元革命のバージョンチェック '20160829 対象革命のバージョン判定処理を追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_KakumeiVer() As String

            Dim rtn_str As String = ""
            Dim tmp_str As String = ""

            'V7バージョン情報が格納されているレジストリを取得
            Dim regkey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\VB and VBA Program Settings\Fkanri50\Ver", False)

            'バージョン情報を取得
            If regkey Is Nothing Then
                Return rtn_str
            Else
                '値を取得 (存在しない場合は空文字)
                tmp_str = DirectCast(regkey.GetValue("ExeVersion", ""), String)
            End If

            '必要な情報のみ取得
            If tmp_str = "" Then
                Return rtn_str
            Else
                rtn_str = tmp_str.Substring(0, 1)
            End If

            Return rtn_str

        End Function

#End Region

#Region "※汎用コンバート用処理"

        ''' <summary>
        ''' 汎用/既存用中間ファイルの照合と中間ファイル作成 '20160812 汎用コンバート対応
        ''' </summary>
        ''' <param name="list_cv"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_BaseMidFileToExistingMidFile(ByVal list_cv As List(Of String), ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim tmp_cnt As Integer = 0
            Dim tmptmpstr As String = ""                                      '20161009 ログ出力処理追加 -add
            Dim tmptmpcnt As Long                                             '20161009 ログ出力処理追加 -add
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
            Dim list_basecvtotal As New List(Of String)

            '----- 全体用プログレスバー初期化 -----
            Dim obj_pgb As New ProgressBarManager
            Dim pgbcnt As Integer = 0
            Dim pgbtotalcnt As Integer = 7                                    '7工程で進捗表示を行う                            
         

            '---------------------------------------------------------------
            ' 汎用用中間ファイルと既存用中間ファイルの照合用仮テーブル作成
            '---------------------------------------------------------------
            '仮テーブル初期化
            Dim midfileinfo_dropqry As String = DBQuery.Qry_DropInfo(MIDFILEINFO_DBNAME, True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, midfileinfo_dropqry, tmp_cnt)

            '仮テーブル新規作成
            Dim midfileinfo_createqry As String = MidFileInfoModule.Get_MidFileInfo_CreateQry()
            DBExec.Exec_NonQuery(Me.sqlcnnv10, midfileinfo_createqry, tmp_cnt)

            '挿入
            Dim normalflg As Boolean = True
            Dim midfileinfo_insertqry As String = MidFileInfoModule.Get_MidFileInfo_InsertQry()
            If midfileinfo_insertqry <> "" Then
                normalflg = DBExec.Exec_NonQuery(Me.sqlcnnv10, midfileinfo_insertqry, tmp_cnt)
            End If
            If normalflg = False Then
                errstr = MSG_ERR_MAKE_CVDB  '20160913_2 エラー時の処理を追加する修正 -add
                rtn = False
                Return rtn
            End If


            '20160905 中間ファイルコピー処理改善 -chg sta
            ''-----------------------------------------------------------------------
            ''作成した照合用仮テーブルからデータを取得してオブジェクトへ格納
            ''-----------------------------------------------------------------------
            'Call Me.Set_MidInfoToObj()
            '-----------------------------------------------------------------------
            ' 作成した照合用仮テーブルからヘッダーに関する情報を取得してオブジェクトへ格納
            '-----------------------------------------------------------------------
            Call Me.Set_MidHeaderInfoToObj()
            '20160905 中間ファイルコピー処理改善 -chg end

            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
            ''-----------------------------------------------------------------------
            '' 既存中間ファイル初期化 '20161007 既存中間ファイル初期化処理の追加 -add
            ''-----------------------------------------------------------------------
            'If CancelFlg = False And MidChkCancelFlg = False Then
            '    normalflg = Me.Init_ExistMidFile(errstr, list_cv)
            '    If normalflg = False Then
            '        rtn = normalflg
            '        Return rtn
            '    End If
            'End If
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

            '-----------------------------------------------------------------------
            ' 仮テーブルから紐付情報を取得して既存用中間ファイルを作成する
            ' →処理が遅いため要速度改善
            '-----------------------------------------------------------------------
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
            '-----------------------------------------------------------------------
            ' 中間仮テーブル作成
            '-----------------------------------------------------------------------
            '移行する全項目をオブジェクトへ格納
            list_basecvtotal = Me.Get_BaseCvItem()

            '仮テーブル作成処理
            For Each basecvitem In list_basecvtotal

                Dim tmp_str() As String = basecvitem.Split("-")
                Dim tmptblname As String = PRE_TBL_NAME & tmp_str(1)

                '初期化
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(tmptblname, True)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_drop, tmp_cnt)

                '作成
                Dim tmp_replacetblname As String = "replacetmptblname"
                Dim tmp_sql_create As String = MidTableModule.MakeMidTable(basecvitem.Replace("-", STR_SPLIT_1), tmp_replacetblname)
                tmp_sql_create = tmp_sql_create.Replace(tmp_replacetblname, tmptblname)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_create, tmp_cnt)

            Next
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end
            
            '**************************************************
            ' 通常コピー
            '**************************************************
            If CancelFlg = False And MidChkCancelFlg = False Then               '20160927 キャンセル処理を追加 CancelFlg = False の条件追加 -add   '20160929 データチェックのキャンセル処理を追加 And MidChkCancelFlg = False の条件も追加 -chg
                Me.lblTotalSituation.Text = SITUATION_MID_READ & " (1/4)"       '----- 全体用プログレスバーラベル表示 -----
                Me.lblCheckSituation.Text = SITUATION_MID_READ & " (1/4)"       '----- 中間用プログレスバーラベル表示 -----
                Me.Refresh()
                normalflg = Me.Set_BaseMidFile_Copy(errstr, list_cv)
                If normalflg = False Then
                    rtn = normalflg
                    Return rtn
                End If
            End If
            If MidChkCancelFlg Then
                Return rtn                                                      '中断時移行の処理をスキップ
            End If

            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
            '通常コピーへの統合に伴いコメントアウト
            ''**************************************************
            '' 変換コピー
            ''**************************************************
            'If CancelFlg = False And MidChkCancelFlg = False Then   '20160927 キャンセル処理を追加 CancelFlg = False の条件追加 -add   '20160929 データチェックのキャンセル処理を追加 And MidChkCancelFlg = False の条件も追加 -chg
            '    Me.lblTotalSituation.Text = SITUATION_MID_READ & " (2/5)"       '----- 全体用プログレスバーラベル表示 -----
            '    Me.lblCheckSituation.Text = SITUATION_MID_READ & " (2/5)"       '----- 中間用プログレスバーラベル表示 -----
            '    Me.Refresh()
            '    normalflg = Me.Set_BaseMidFile_ChgCopyMain(errstr, list_cv)
            '    If normalflg = False Then
            '        rtn = normalflg
            '        Return rtn
            '    End If
            'End If

            'If MidChkCancelFlg Then
            '    Return rtn                                                      '中断時移行の処理をスキップ
            'End If
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end           
            '**************************************************
            ' 自動設定処理 契約No等に「1」を自動設定                            '20160920 プログラム自動設定箇所の追加 -add
            '**************************************************
            If CancelFlg = False And MidChkCancelFlg = False Then   '20160927 キャンセル処理を追加 CancelFlg = False の条件追加 -add   '20160929 データチェックのキャンセル処理を追加 And MidChkCancelFlg = False の条件も追加 -chg
                Me.lblTotalSituation.Text = SITUATION_DATACONDITION & " (2/4)"  '----- 全体用プログレスバーラベル表示 -----
                Me.lblCheckSituation.Text = SITUATION_DATACONDITION & " (2/4)"  '----- 中間用プログレスバーラベル表示 -----
                Me.Refresh()

                '----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                Console.WriteLine(Now & "調整①：管理Noを「1」に設定")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整①"), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
                
                normalflg = Me.Set_ExistMidFile_DefCopyMain(errstr)
                If normalflg = False Then
                    rtn = normalflg
                    Return rtn
                End If
            End If
            If MidChkCancelFlg Then
                Return rtn                                                      '中断時移行の処理をスキップ
            End If

            '******************************************************
            ' ファイル名集約                                                    '20161011 汎用→既存書込時の不要処理除去修正 -add
            '******************************************************
            '----- ログ出力 -----                                               '20161009 ログ出力処理追加 -add
            Console.WriteLine(Now & "調整②：ファイル名集約")
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整②"), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

            Dim list_filename As New List(Of String)
            list_filename = Me.Set_FileNameIntensive(list_cv)

            '******************************************************
            ' 自動設定処理 所有形態に一棟/区分の設定値を自動設定                 '20160920 プログラム自動設定箇所の追加 -add
            '******************************************************
            If CancelFlg = False And MidChkCancelFlg = False Then
                '20161011 汎用→既存書込時の不要処理除去修正 -chg sta
                'normalflg = Me.Set_ExistMidFile_SyoyuCopyMain(errstr)
                Me.lblTotalSituation.Text = SITUATION_DATACONDITION & " (3/4)"  '----- 全体用プログレスバーラベル表示 -----
                Me.lblCheckSituation.Text = SITUATION_DATACONDITION & " (3/4)"  '----- 中間用プログレスバーラベル表示 -----
                Me.Refresh()

                '----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                Console.WriteLine(Now & "調整③：所有形態自動設定")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整③"), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

                normalflg = Me.Set_ExistMidFile_SyoyuCopyMain(errstr, list_filename)
                '20161011 汎用→既存書込時の不要処理除去修正 -chg end
                If normalflg = False Then
                    rtn = normalflg
                    Return rtn
                End If
            End If
            If MidChkCancelFlg Then
                Return rtn                                                      '中断時移行の処理をスキップ
            End If

            '******************************************************
            ' 適用開始年月自動設定
            '******************************************************
            '20161004 適用開始年月の自動設定 -add sta
            If CancelFlg = False And MidChkCancelFlg = False Then
                '20161011 汎用→既存書込時の不要処理除去修正 -chg sta
                'normalflg = Me.Set_ExistMidFile_SoruleTekiyoYmdCopyMain(errstr)
                Me.lblTotalSituation.Text = SITUATION_DATACONDITION & " (4/4)"  '----- 全体用プログレスバーラベル表示 -----
                Me.lblCheckSituation.Text = SITUATION_DATACONDITION & " (4/4)"  '----- 中間用プログレスバーラベル表示 -----
                Me.Refresh()

                '----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                Console.WriteLine(Now & "調整④：適用開始年月自動設定")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整④"), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
               
                normalflg = Me.Set_ExistMidFile_SoruleTekiyoYmdCopyMain(errstr, list_filename)
                '20161011 汎用→既存書込時の不要処理除去修正 -chg end
                If normalflg = False Then
                    rtn = normalflg
                    Return rtn
                End If
            End If
            '20161004 適用開始年月の自動設定 -add end
            If MidChkCancelFlg Then
                Return rtn                                                      '中断時移行の処理をスキップ
            End If

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 既存用中間ファイルのヘッダー名をオブジェクトへ格納 '20160812 汎用コンバート対応
        ''' '20160905 中間ファイルコピー処理改善 引数を追加
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_MidHeaderInfoToObj()

            Dim readtbl As New DataTable
            Dim reccnt As Integer = 0
            Dim fldname As String = ""
            Dim fldvalue As String = ""
            Dim tmp_既存中間シート名 As String = ""
            Dim tmp_既存中間フィールド名 As String = ""
            Dim tmp_汎用中間ファイル名 As String = ""
            Dim tmp_汎用中間シート名 As String = ""
            Dim tmp_汎用中間フィールド名_大分類 As String = ""
            Dim tmp_汎用中間フィールド名_小分類 As String = ""


            '--------------------------------------------------
            ' 中間ファイル情報取得
            '--------------------------------------------------
            '20160905 中間ファイルコピー処理改善 -chg sta
            'Dim tmp_sql As String = " SELECT DISTINCT * FROM " & MIDFILEINFO_DBNAME
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT DISTINCT * FROM " & MIDFILEINFO_DBNAME
            '20160905 中間ファイルコピー処理改善 -chg end
            reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl)

            '--------------------------------------------------
            ' 初期化
            '--------------------------------------------------
            List_Existmidheader.Clear()
            List_Basemidheader.Clear()

            '--------------------------------------------------
            ' 読込開始
            '--------------------------------------------------
            For cntii As Integer = 0 To readtbl.Rows.Count - 1

                '中断処理
                Application.DoEvents()
                If CancelFlg Or MidChkCancelFlg Then
                    Exit Sub
                End If

                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                    '項目名取得
                    fldname = readtbl.Columns(cntjj).ColumnName

                    '登録値取得
                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                    '各項目値→変数格納
                    Select Case fldname
                        Case "既存中間シート名"
                            tmp_既存中間シート名 = fldvalue.Trim
                        Case "既存中間フィールド名"
                            tmp_既存中間フィールド名 = fldvalue.Trim
                        Case "汎用中間シート名"
                            tmp_汎用中間シート名 = fldvalue.Trim
                        Case "汎用中間フィールド名_大分類"
                            tmp_汎用中間フィールド名_大分類 = fldvalue.Trim
                        Case "汎用中間フィールド名_小分類"
                            tmp_汎用中間フィールド名_小分類 = fldvalue.Trim
                    End Select

                Next

                '--------------------------------------------------
                ' 既存用中間ファイルと紐付けて各オブジェクトへ格納
                '--------------------------------------------------
                Dim tmp_existmidheader As String = tmp_既存中間シート名 & STR_SPLIT_1 & tmp_既存中間フィールド名
                If tmp_existmidheader = STR_SPLIT_1 Then
                    tmp_existmidheader = ""
                End If

                '既存用中間ファイルのヘッダー
                If tmp_existmidheader <> "" Then
                    List_Existmidheader.Add(tmp_existmidheader)
                End If

                '--------------------------------------------------
                ' 汎用用中間ファイルと紐付けて各オブジェクトへ格納
                '--------------------------------------------------
                Dim tmp_basemidheader As String = ""
                If tmp_汎用中間フィールド名_小分類 <> "" Then
                    '20160913_2 部屋設備移行処理の追加 -chg sta
                    'tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_小分類
                    tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_大分類 & STR_SPLIT_1 & tmp_汎用中間フィールド名_小分類
                    '20160913_2 部屋設備移行処理の追加 -chg end
                Else
                    tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_大分類
                End If
                If tmp_basemidheader = STR_SPLIT_1 Then
                    tmp_basemidheader = ""
                End If

                '汎用用中間ファイルのヘッダー
                If tmp_basemidheader <> "" And List_Basemidheader.Contains(tmp_basemidheader) = False Then
                    List_Basemidheader.Add(tmp_basemidheader)
                End If

            Next

        End Sub

        ''' <summary>
        ''' 既存中間ファイル初期化処理 '20161007 既存中間ファイル初期化処理の追加 -add
        ''' </summary>
        ''' <param name="errstr"></param>
        ''' <param name="list_cv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Init_ExistMidFile(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim list_filename As New List(Of String)


            'ファイル名を集約
            For Each cvitem In list_cv
                Dim tmp_str() As String = cvitem.Split("-")
                Dim filename As String = tmp_str(0)
                If list_filename.Contains(filename) = False Then
                    list_filename.Add(filename)
                End If
            Next

            For Each initfile In list_filename

                '中断処理
                Application.DoEvents()
                If CancelFlg Or MidChkCancelFlg Then
                    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                    Return False
                End If

                'ファイルパス設定
                Dim filepath As String = EtcMethod.Set_Path(MiddleDirPath, initfile & ".xlsx")

                '初期化処理
                Try
                    'ファイル準備
                    appli = CreateObject("Excel.Application")
                    appli.Visible = False
                    wbook = appli.Workbooks.Open(filepath)

                    For cntii = 1 To wbook.Worksheets.Count
                        'シートを取得
                        wsheet = wbook.Worksheets(cntii)
                        '最大行を取得
                        Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count

                        If maxrowcnt > 1 Then
                            '既存データ範囲を取得
                            Dim dataarea As String = EXISTMIDFILE_READWRITE_ROW.ToString & ":" & maxrowcnt.ToString
                            wsheet.Rows(dataarea).Delete()
                        End If
                    Next

                Catch ex As Exception
                    errstr = MSG_ERR_EXISTMIDFILEINIT
                    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                    Return False
                End Try

                '終了処理
                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

            Next

            Return True

        End Function

        ''' <summary>
        ''' 汎用用中間ファイル→既存用仮テーブル通常処理 '20160929 汎用→既存コピー処理改善対応
        ''' </summary>
        ''' <param name="errstr"></param>
        ''' <param name="list_cv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_BaseMidFile_Copy(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

            Dim rtn As Boolean = True
            Dim dt As New DataTable()
            Dim dr As OleDbDataReader
            Dim qry As String = Nothing

            Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            Dim tmptmpstr As String                                                         '20161009 ログ出力処理追加 -add
            Dim tmptmpcnt As Long                                                           '20161009 ログ出力処理追加 -add

            '----- 全体用・中間用プログレスバー更新 -----
            Dim obj_pgb As New ProgressBarManager
            Dim pgbcnt As Integer = 0
            Dim pgbtotalcnt_total As Integer = list_cv.Count
            Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
            Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)
            pgbtotalcnt = pgbtotalcnt_total

            '20161017 Excel接続処理の外出し修正 -add sta
            '----- 汎用用中間ファイルへの接続・オープン処理 -----
            Dim con_read As New OleDbConnection()
            Dim cmd_read As New OleDbCommand()
            Dim basemidpath As String = EtcMethod.Set_Path(BaseMidDirPath, CV_FROM_MIDDLE & ".xlsx")        '汎用中間ファイルパスを取得

            '接続文字列生成
            con_read.ConnectionString = _
                "Provider=" & tmpprovider & _
                "Data Source=" & basemidpath & ";" & _
                "Extended Properties=" & """" & tmpextend & """"
            '接続設定
            cmd_read.Connection = con_read
            con_read.Open()
            '20161017 Excel接続処理の外出し修正 -add end

            For Each cvitem In list_cv

                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
                '変換コピー処理を統合
                '自社口座、部屋設備はそのまま仮テーブルを作成する
                '既存処理は「Case Else」の中へそのまま移動
                Select Case cvitem
                    Case "部屋情報-部屋駐車場情報", "部屋情報-部屋特約情報", "契約情報-契約契約者情報", "契約情報-契約保証人情報", "契約情報-契約特約事項情報"
                        '変換コピー
                        rtn = Me.Set_BaseMidFile_ChgCopyMain(cvitem, con_read, cmd_read, errstr)
                        If rtn = False Then
                            con_read.Close()
                            Return rtn
                        End If
                        '20161025 部屋設備取得方法の修正 -chg sta
                        'Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"
                        '    '汎用用仮テーブル作成
                        '    rtn = Me.Set_BaseMidFile_NormalCopyMain(cvitem, con_read, cmd_read, errstr)
                        '    If rtn = False Then
                        '        con_read.Close()
                        '        Return rtn
                        '    End If
                    Case "自社情報-自社口座情報"
                        '汎用用仮テーブル作成
                        rtn = Me.Set_BaseMidFile_JisyaCopyMain(cvitem, con_read, cmd_read, errstr)
                        If rtn = False Then
                            con_read.Close()
                            Return rtn
                        End If
                    Case "部屋情報-部屋設備情報"
                        '汎用用仮テーブル作成
                        rtn = Me.Set_BaseMidFile_SetubiCopyMain(cvitem, errstr)
                        If rtn = False Then
                            con_read.Close()
                            Return rtn
                        End If
                        '20161025 部屋設備取得方法の修正 -chg end
                        '20161028 物件/部屋鍵取得方法修正 -add sta
                    Case "各マスタ情報-鍵タイトルマスタ"
                        '汎用用仮テーブル作成
                        rtn = Me.Set_BaseMidFile_KagiTitleCopyMain(cvitem, con_read, cmd_read, errstr)
                        If rtn = False Then
                            con_read.Close()
                            Return rtn
                        End If
                        '20161028 物件/部屋鍵取得方法修正 -add end
                    Case Else
                        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end
                        '20161017 DoEvent処理の移動 -del sta
                        ''中断処理
                        'Application.DoEvents()
                        'If CancelFlg Or MidChkCancelFlg Then
                        '    Return rtn
                        'End If
                        '20161017 DoEvent処理の移動 -del end
                        'コピーフラグ
                        Dim copyflg As Boolean = True

                        '--------------------------------------------------
                        ' 既存データ関連
                        '--------------------------------------------------
                        'コンバート対象項目から既存ファイル名、シート名を取得
                        cvitem = cvitem.Replace("-", STR_SPLIT_1)
                        Dim tmp_existmidstr() As String = Split(cvitem, STR_SPLIT_1)
                        Dim existfilename As String = tmp_existmidstr(0)
                        Dim existsheetname As String = tmp_existmidstr(1)
                        Dim existmidpath As String = ""                                                     '既存中間ファイルパス

                        '既存中間ファイルパスを取得
                        existmidpath = EtcMethod.Set_Path(MiddleDirPath, existfilename & ".xlsx")

                        '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
                        Call Me.Set_MidDataInfoToObj(existfilename, existsheetname, False)

                        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
                        Dim tmp_cnt As Integer = 0
                        Dim tmptblname As String = PRE_TBL_NAME & existsheetname            'シート名は仮テーブルのテーブル名の一部にも使用する
                        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

                        '--------------------------------------------------
                        ' 汎用データ関連
                        '--------------------------------------------------
                        '移行項目からオープンする汎用用中間ファイル名を取得
                        Dim basemidfsname As String = ""                                                    '汎用中間ファイルのファイル名-シート名格納用
                        Dim basefilename As String = ""                                                     '汎用中間ファイル名
                        Dim basesheetname As String = ""                                                    '汎用中間シート名
                        '20161017 Excel接続処理の外出し修正 -del
                        'Dim basemidpath As String = ""                                                      '汎用中間ファイルパス

                        basemidfsname = Hash_ExistMidToBaseMid_FS(cvitem)
                        If basemidfsname <> "" Then
                            Dim tmp_basemidstr() As String = Split(basemidfsname, STR_SPLIT_1)
                            basefilename = tmp_basemidstr(0)
                            basesheetname = tmp_basemidstr(1)
                            '20161017 Excel接続処理の外出し修正 -del
                            'basemidpath = EtcMethod.Set_Path(BaseMidDirPath, basefilename & ".xlsx")        '汎用中間ファイルパスを取得
                        Else
                            copyflg = False
                        End If

                        '--------------------------------------------------
                        ' メイン処理
                        '--------------------------------------------------

                        If copyflg Then
                            '20161017 Excel接続処理の外出し修正 -del sta
                            'Dim con_read As New OleDbConnection()
                            'Dim cmd_read As New OleDbCommand()
                            '20161017 Excel接続処理の外出し修正 -del end
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                            'Dim con_write As New OleDbConnection()
                            'Dim cmd_write As New OleDbCommand()
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                            Try
                                '20161017 Excel接続処理の外出し修正 -del sta
                                ''--------------------------------------------------
                                '' ①EXCEL接続 (ファイル読込用)
                                ''--------------------------------------------------
                                ''接続文字列生成
                                'con_read.ConnectionString = _
                                '    "Provider=" & tmpprovider & _
                                '    "Data Source=" & basemidpath & ";" & _
                                '    "Extended Properties=" & """" & tmpextend & """"
                                ''接続設定
                                'cmd_read.Connection = con_read
                                '20161017 Excel接続処理の外出し修正 -del end
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                                ''--------------------------------------------------
                                '' ②EXCEL接続 (ファイル書込用)
                                ''--------------------------------------------------
                                ''接続文字列生成
                                'con_write.ConnectionString = _
                                '    "Provider=" & tmpprovider & _
                                '    "Data Source=" & existmidpath & ";" & _
                                '    "Extended Properties=" & """" & tmpextend & """"
                                ''接続設定
                                'cmd_write.Connection = con_write

                                ''--------------------------------------------------
                                '' 読込→書込
                                ''--------------------------------------------------
                                ''接続オープン処理
                                ''20161017 Excel接続処理の外出し修正 -del
                                ''con_read.Open()
                                'con_write.Open()
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                                '----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                                Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDCOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & basesheetname & ")")
                                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, LOG_SYORIKOMK_MIDFILE, basesheetname), False)
                                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

                                qry = "SELECT * FROM [" & basesheetname & "$] "
                                cmd_read = con_read.CreateCommand
                                cmd_read.CommandText = qry
                                dt = New DataTable
                                dr = cmd_read.ExecuteReader
                                dt.Load(dr)

                                Dim colcnt As Integer = dt.Columns.Count
                                Dim readrowcnt As Integer = 1
                                Dim solist_header As New SortedList(Of Integer, String)

                                'コピー項目表示
                                Me.lblCVItem.Text = existfilename & " ： " & existsheetname

                                '----- 個別プログレスバー初期化 -----
                                Dim pgbtotalcnt_part As Integer = 0                             'プログレスバー総件数初期化
                                Dim pgbbasecnt As Integer = 100                                 '実件数で表示するか否かの基準値
                                If dt.Rows.Count <= pgbbasecnt Then
                                    pgbtotalcnt_part = dt.Rows.Count
                                Else
                                    pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                                End If
                                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

                                '1行ずつ取得
                                Dim row As DataRow
                                For Each row In dt.Rows

                                    '20161017 DoEvent処理の移動 -add sta
                                    '中断処理
                                    Application.DoEvents()
                                    If CancelFlg Or MidChkCancelFlg Then
                                        Return rtn
                                    End If
                                    '20161017 DoEvent処理の移動 -add end

                                    Dim tmp_hash As New Hashtable
                                    Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                                    Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト
                                    Dim empty_gyo As Boolean = True

                                    '行データを取得して作業用オブジェクトへ格納
                                    'Excelファイルのヘッダが終わる11番目行から一行が全部空いている場合には処理しません。
                                    'tmp_solistはExcelファイルから読み込んだデータが保存されます。
                                    For cntjj = 1 To colcnt - 1
                                        tmp_solist.Add(cntjj, row(cntjj).ToString)
                                        '一行の中で""以外の情報があれば読み込むようにするフラッグです。
                                        If tmp_solist(cntjj) <> "" Then
                                            empty_gyo = False
                                        End If
                                    Next

                                    '行数から判断する
                                    If readrowcnt = 1 Then

                                        '1行目はヘッダーなのでヘッダー用のオブジェクトへ格納
                                        solist_header = tmp_solist

                                    ElseIf readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                                        If empty_gyo = False Then
                                            '実データ開始行からデータ格納用オブジェクトへ格納
                                            solist_value = tmp_solist

                                            '汎用ヘッダーから既存ヘッダーを取得し、実データを紐付けてオブジェクトへ格納する
                                            For Each existmid In Hash_ExistMidToBaseMid_SH
                                                '既存中間ファイルと汎用中間ファイルの紐付状況を取得
                                                Dim existmidshname_fromobj As String = existmid.Key
                                                Dim basemidshname_fromobj As String = existmid.Value
                                                If basemidshname_fromobj <> "" Then
                                                    For cntkk = 1 To colcnt - 1
                                                        Dim tmp_baseheader As String = solist_header(cntkk)
                                                        Dim tmp_basevalue As String = solist_value(cntkk)
                                                        Dim tmp_syogostr As String = basesheetname & STR_SPLIT_1 & tmp_baseheader
                                                        Dim tmp_existheader As String = ""
                                                        If basemidshname_fromobj = tmp_syogostr Then
                                                            Dim tmp_str() As String = Split(existmidshname_fromobj, STR_SPLIT_1)
                                                            tmp_hash.Add(tmp_str(1), tmp_basevalue)
                                                        End If
                                                    Next
                                                End If
                                            Next

                                            '挿入用に編集
                                            Dim taisyo As String = ""
                                            Dim value As String = ""

                                            '20161205 初回契約日を契約日から取得するように修正 -add sta
                                            '契約日が存在しない場合があるためチェックを行う
                                            '存在しない場合は契約開始日をセットする
                                            If cvitem = "契約情報" & STR_SPLIT_1 & "契約基本情報" Then
                                                '初回契約日チェック
                                                If tmp_hash("初回契約日") = "" Then
                                                    tmp_hash("初回契約日") = tmp_hash("契約開始日")
                                                End If
                                            End If
                                            If cvitem = "契約情報" & STR_SPLIT_1 & "契約履歴情報" Then
                                                '契約日チェック
                                                If tmp_hash("契約日") = "" Then
                                                    tmp_hash("契約日") = tmp_hash("契約開始日")
                                                End If
                                            End If
                                            '20161205 初回契約日を契約日から取得するように修正 -add end

                                            For Each item In tmp_hash
                                                Dim tmp_taisyo As String = "[" & item.Key & "]"
                                                '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg sta
                                                'Dim tmp_value As String = "'" & item.Value & "'"
                                                Dim tmp_value As String = item.Value
                                                If tmp_value Is Nothing Then
                                                    tmp_value = ""
                                                End If
                                                tmp_value = "'" & tmp_value.Replace("'", "’") & "'"
                                                '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                                                taisyo = taisyo & "," & tmp_taisyo
                                                value = value & "," & tmp_value
                                            Next
                                            taisyo = "(" & taisyo.Remove(0, 1) & ")"
                                            value = "(" & value.Remove(0, 1) & ")"

                                            '挿入
                                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                            'qry = ""
                                            'qry += " INSERT INTO [" & existsheetname & "$] "
                                            'qry += taisyo
                                            'qry += " VALUES "
                                            'qry += value
                                            'cmd_write.CommandText = qry
                                            'cmd_write.ExecuteNonQuery()
                                            Dim tmp_qry_insert As String = ""
                                            tmp_qry_insert += " INSERT INTO " & tmptblname
                                            tmp_qry_insert += taisyo
                                            tmp_qry_insert += " VALUES "
                                            tmp_qry_insert += value
                                            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_qry_insert, tmp_cnt)
                                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                                        End If

                                        '----- プログレスバー更新/進捗率表示 -----
                                        Dim tmp_pgbcnt As Integer = 0
                                        If dt.Rows.Count <= pgbbasecnt Then             '基準件数(100件)以下の場合は実件数を取得
                                            tmp_pgbcnt = readrowcnt
                                        ElseIf dt.Rows.Count > readrowcnt Then          '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                                            Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                                        ElseIf dt.Rows.Count <= readrowcnt Then         '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                                            tmp_pgbcnt = pgbtotalcnt_part
                                        End If

                                        '表示
                                        If tmp_pgbcnt <> 0 Then
                                            Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                                            Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                                        End If
                                        '20161017 進捗表示処理による速度低下の修正 -del
                                        'Me.Refresh()
                                    End If
                        readrowcnt = readrowcnt + 1
                                Next

                            Catch ex As Exception

                                'エラー処理
                                errstr = MSG_ERR_BASEMIDFILEREAD
                                rtn = False

                                'クローズ処理
                                con_read.Close()
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                                'con_write.Close()

                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
                                Return rtn

                            End Try

                            'クローズ処理
                            '20161017 Excel接続処理の外出し修正 -del
                            'con_read.Close()
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                            'con_write.Close()
                        End If

                End Select

                '----- 全体用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsettingTotal(pgbcnt)
                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt, True)
                '----- 中間用プログレスバー更新 -----
                Call obj_pgb.pgbsettingChkTotal(pgbcnt)
                Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt, True)
                Me.Refresh()

            Next

            '20161017 Excel接続処理の外出し修正 -add
            'クローズ処理
            con_read.Close()

            Return rtn

        End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
        ''' <summary>
        ''' 汎用用中間ファイル→既存用仮テーブル個別処理 
        ''' </summary>
        ''' <param name="cvitem"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_BaseMidFile_ChgCopyMain(ByVal cvitem As String, ByVal con_read As OleDbConnection, ByVal cmd_read As OleDbCommand, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim obj_rep As New Object
            Dim tmptmpstr As String = ""                                '20161009 ログ出力処理追加 -add
            Dim tmptmpcnt As Long                                       '20161009 ログ出力処理追加 -add
            Dim basemidfilesheetname As String = ""
            Dim dt As New DataTable()
            Dim dr As OleDbDataReader

            '----- 仮テーブル名取得 -----
            Dim tmp_str() As String = cvitem.Split("-")
            Dim existsheetname As String = tmp_str(1)
            Dim tmptblname As String = PRE_TBL_NAME & existsheetname            'シート名は仮テーブルのテーブル名の一部にも使用する

            '----- 各項目毎に移行用のオブジェクトを生成 -----
            Select Case cvitem
                Case "部屋情報-部屋駐車場情報"
                    obj_rep = New Njc.Repository.Hy_cyusyajo_mid_Repository
                    Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
                    basemidfilesheetname = "部屋情報"
                Case "部屋情報-部屋特約情報"
                    obj_rep = New Njc.Repository.Hy_tokuyaku_mid_Repository
                    Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
                    basemidfilesheetname = "部屋情報"
                Case "契約情報-契約契約者情報"
                    obj_rep = New Njc.Repository.Ky_kys_mid_Repository
                    Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
                    basemidfilesheetname = "契約契約者保証人情報"
                Case "契約情報-契約保証人情報"
                    obj_rep = New Njc.Repository.Ky_hosyonin_mid_Repository
                    Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
                    basemidfilesheetname = "契約契約者保証人情報"
                Case "契約情報-契約特約事項情報"
                    obj_rep = New Njc.Repository.Ky_tokuyaku_mid_Repository
                    Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
                    basemidfilesheetname = "契約特約およびメモ情報"
            End Select

            '----- ログ出力 -----                       '20161009 ログ出力処理追加 -add
            Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, LOG_SYORIKOMK_MIDFILE, cvitem), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

            '----- データ取得 -----
            Dim tmp_sql As String = "SELECT * FROM [" & basemidfilesheetname & "$] "
            cmd_read = con_read.CreateCommand
            cmd_read.CommandText = tmp_sql
            dt = New DataTable
            dr = cmd_read.ExecuteReader
            dt.Load(dr)

            '----- 項目毎に個別処理開始 -----
            rtn = obj_rep.Read_BaseMidFile(Me.sqlcnnv10, dt, tmptblname)

            '----- コピー処理失敗時は処理を抜ける -----
            If rtn = False Then

                errstr = MSG_ERR_BASEMIDFILEREAD

                '----- ログ出力 -----                   '20161009 ログ出力処理追加 -add
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
                Return rtn

            End If

            '----- 返却 -----
            Return rtn

        End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
        ''' <summary>
        ''' 汎用用中間ファイル→既存用仮テーブル個別処理_自社 
        ''' 自社口座情報は汎用中間ファイルの形でそのまま仮テーブルを作成する
        ''' </summary>
        ''' <param name="cvitem"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_BaseMidFile_JisyaCopyMain(ByVal cvitem As String, ByVal con_read As OleDbConnection, ByVal cmd_read As OleDbCommand, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim obj_rep As New Object
            Dim tmptmpstr As String = ""                                '20161009 ログ出力処理追加 -add
            Dim tmptmpcnt As Long                                       '20161009 ログ出力処理追加 -add

            Dim basesheetname As String = ""
            Dim dt As New DataTable()
            Dim dr As OleDbDataReader
            Dim obj_pgb As New ProgressBarManager

            '----- 仮テーブル名取得 -----
            Dim tmp_str() As String = cvitem.Split("-")
            Dim existsheetname As String = tmp_str(1)
            Dim tmp_cnt As Integer = 0
            Dim tmptblname As String = PRE_TBL_NAME & existsheetname            'シート名は仮テーブルのテーブル名の一部にも使用する
            '20161025 部屋設備取得方法の修正 -add sta
            basesheetname = tmp_str(1)
            Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
            '20161025 部屋設備取得方法の修正 -add end

            '20161025 部屋設備取得方法の修正 -del sta
            ''----- 各項目毎に移行用のオブジェクトを生成 -----
            'Select Case cvitem
            '    Case "自社情報-自社口座情報"
            '        basesheetname = "自社口座情報"
            '        Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
            '    Case "部屋情報-部屋設備情報"
            '        basesheetname = "部屋設備情報"
            '        Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
            'End Select
            '20161025 部屋設備取得方法の修正 -del end
            '----- ログ出力 -----             
            Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, LOG_SYORIKOMK_MIDFILE, cvitem), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

            '----- メイン処理 -----
            Try

                '----- ログ出力 -----                                          
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDCOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & basesheetname & ")")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, LOG_SYORIKOMK_MIDFILE, basesheetname), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

                Dim tmp_sql As String = "SELECT * FROM [" & basesheetname & "$] "
                cmd_read = con_read.CreateCommand
                cmd_read.CommandText = tmp_sql
                dt = New DataTable
                dr = cmd_read.ExecuteReader
                dt.Load(dr)

                Dim colcnt As Integer = dt.Columns.Count
                Dim readrowcnt As Integer = 1

                '----- 個別プログレスバー初期化 -----
                Dim pgbtotalcnt_part As Integer = 0                             'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100                                 '実件数で表示するか否かの基準値
                If dt.Rows.Count <= pgbbasecnt Then
                    pgbtotalcnt_part = dt.Rows.Count
                Else
                    pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

                '1行ずつ取得
                Dim row As DataRow
                For Each row In dt.Rows

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Or MidChkCancelFlg Then
                        Return rtn
                    End If

                    Dim tmp_hash As New Hashtable
                    Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                    Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト

                    '行データを取得して作業用オブジェクトへ格納
                    Dim insertvalue As String = ""
                    For cntjj = 1 To colcnt - 1
                        '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg sta
                        'insertvalue = insertvalue & "," & "'" & row(cntjj).ToString & "'"
                        If row(cntjj) Is Nothing Then
                            row(cntjj) = ""
                        End If
                        insertvalue = insertvalue & "," & "'" & row(cntjj).ToString.Replace("'", "’") & "'"
                        '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                    Next

                    'データ部以降から挿入処理
                    If readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                        '挿入用に成形
                        insertvalue = insertvalue.Remove(0, 1)

                        '挿入処理
                        Dim tmp_sql_insert As String = " INSERT INTO " & tmptblname & " VALUES(" & insertvalue & ")"
                        DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_insert, tmp_cnt)

                    End If

                    '----- プログレスバー更新/進捗率表示 -----
                    Dim tmp_pgbcnt As Integer = 0
                    If dt.Rows.Count <= pgbbasecnt Then             '基準件数(100件)以下の場合は実件数を取得
                        tmp_pgbcnt = readrowcnt
                    ElseIf dt.Rows.Count > readrowcnt Then          '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                    ElseIf dt.Rows.Count <= readrowcnt Then         '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        tmp_pgbcnt = pgbtotalcnt_part
                    End If

                    '表示
                    If tmp_pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                        Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                    End If

                    readrowcnt = readrowcnt + 1

                Next

            Catch ex As Exception

                'エラー処理
                rtn = False
                errstr = MSG_ERR_BASEMIDFILEREAD

                '----- ログ出力 -----              
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

                Return rtn

            End Try

            Return rtn

        End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

        '20161025 部屋設備取得方法の修正 -add sta
        ''' <summary>
        ''' 汎用用中間ファイル→既存用仮テーブル個別処理_設備
        ''' </summary>
        ''' <param name="cvitem"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_BaseMidFile_SetubiCopyMain(ByVal cvitem As String, ByRef errstr As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
            Dim rtn As Boolean = True                                       '戻り値
            Dim tmp_cnt As Integer = 0
            Dim tmptmpstr As String = ""
            Dim tmptmpcnt As Long

            '************************
            '作業準備
            '************************

            '移行項目、仮テーブル名セット
            Dim tmp_str() As String = cvitem.Split("-")
            Dim basesheetname As String = tmp_str(1)
            Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
            Dim tmptblname As String = PRE_TBL_NAME & basesheetname            'シート名は仮テーブルのテーブル名の一部にも使用する

            'Excelファイル初期設定 (紐付ファイルの読込)
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, BaseMidDirPath, CV_FROM_MIDDLE, basesheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If rtn = False Then
                errstr = MSG_ERR_BASEMIDFILEREAD
                '20161027 部屋設備ログ出力箇所修正 -add sta
                '----- ログ出力 -----               
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
                '20161027 部屋設備ログ出力箇所修正 -add end
                Return rtn
            End If

            'プログレスバー初期化
            Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
            Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

            If rowcnt <= pgbbasecnt Then
                pgbtotalcnt = rowcnt
            Else
                pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
            End If
            Call obj_pgb.pgbInitPart(pgbtotalcnt)

            '************************
            '処理開始
            '************************

            '----- ログ出力 -----                          
            Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDCOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & basesheetname & ")")
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, LOG_SYORIKOMK_MIDFILE, basesheetname), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

            With model_cvitem

                '---------------
                'データ部処理
                '---------------
                For cntii = 1 To rowcnt

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                        Return rtn
                    End If

                    '行取得
                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 2), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                    '挿入用に成形
                    Dim insertvalue As String = ""
                    For cntjj = 1 To datarowvalue.Length
                        '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg sta
                        'insertvalue = insertvalue & "," & "'" & datarowvalue(1, cntjj) & "'"
                        If datarowvalue(1, cntjj) Is Nothing Then
                            datarowvalue(1, cntjj) = ""
                        End If
                        insertvalue = insertvalue & "," & "'" & datarowvalue(1, cntjj).ToString.Replace("'", "’") & "'"
                        '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                    Next
                    If insertvalue.Trim <> "" Then
                        insertvalue = insertvalue.Remove(0, 1)
                    End If

                    '挿入処理
                    Dim tmp_sql_insert As String = " INSERT INTO " & tmptblname & " VALUES(" & insertvalue & ")"
                    DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_insert, tmp_cnt)

                    '-------------------------------
                    'プログレスバー更新/進捗率表示
                    '-------------------------------
                    '件数取得
                    Dim pgbcnt As Integer = 0
                    If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        pgbcnt = cntii
                    ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                    ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        pgbcnt = pgbtotalcnt
                    End If

                    '表示
                    If pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(pgbcnt)
                        Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                    End If

                Next

            End With

            '************************
            '終了処理
            '************************
            '20161027 部屋設備ログ出力箇所修正 -del sta
            ''----- ログ出力 -----               
            'Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
            'tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
            'DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
            '20161027 部屋設備ログ出力箇所修正 -del end
            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            '返却
            Return rtn

        End Function
        '20161025 部屋設備取得方法の修正 -add end

        '20161028 物件/部屋鍵取得方法修正 -add sta
        ''' <summary>
        ''' 汎用用中間ファイル→既存用仮テーブル個別処理_鍵タイトル
        ''' </summary>
        ''' <param name="cvitem"></param>
        ''' <param name="con_read"></param>
        ''' <param name="cmd_read"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_BaseMidFile_KagiTitleCopyMain(ByVal cvitem As String, ByVal con_read As OleDbConnection, ByVal cmd_read As OleDbCommand, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim obj_rep As New Object
            Dim tmptmpstr As String = ""
            Dim tmptmpcnt As Long

            Dim basesheetname As String() = New String() {"", "物件鍵タイトルマスタ", "部屋鍵タイトルマスタ"}   '要素数も使用するため1から値をセットする
            Dim dt As New DataTable()
            Dim dr As OleDbDataReader
            Dim obj_pgb As New ProgressBarManager

            '----- 仮テーブル名取得 -----
            Dim tmp_str() As String = cvitem.Split("-")
            Dim existsheetname As String = tmp_str(1)
            Dim tmp_cnt As Integer = 0
            Dim tmptblname As String = PRE_TBL_NAME & existsheetname            'シート名は仮テーブルのテーブル名の一部にも使用する

            '----- ログ出力 -----             
            Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, LOG_SYORIKOMK_MIDFILE, cvitem), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

            '----- メイン処理 -----
            Try

                '----- ログ出力 -----                                          
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDCOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & existsheetname & ")")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, LOG_SYORIKOMK_MIDFILE, existsheetname), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

                For cntii = 1 To UBound(basesheetname)

                    '移行項目表示
                    Me.lblCVItem.Text = basesheetname(cntii)

                    Dim tmp_sql As String = "SELECT * FROM [" & basesheetname(cntii) & "$] "
                    cmd_read = con_read.CreateCommand
                    cmd_read.CommandText = tmp_sql
                    dt = New DataTable
                    dr = cmd_read.ExecuteReader
                    dt.Load(dr)

                    Dim colcnt As Integer = dt.Columns.Count
                    Dim readrowcnt As Integer = 1

                    '----- 個別プログレスバー初期化 -----
                    Dim pgbtotalcnt_part As Integer = 0                             'プログレスバー総件数初期化
                    Dim pgbbasecnt As Integer = 100                                 '実件数で表示するか否かの基準値
                    If dt.Rows.Count <= pgbbasecnt Then
                        pgbtotalcnt_part = dt.Rows.Count
                    Else
                        pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                    End If
                    Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

                    '1行ずつ取得
                    Dim row As DataRow
                    For Each row In dt.Rows

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Or MidChkCancelFlg Then
                            Return rtn
                        End If

                        Dim tmp_hash As New Hashtable
                        Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                        Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト

                        '行データを取得して作業用オブジェクトへ格納
                        Dim insertvalue As String = ""
                        For cntjj = 1 To colcnt - 1
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg sta
                            'insertvalue = insertvalue & "," & "'" & row(cntjj).ToString & "'"
                            If row(cntjj) Is Nothing Then
                                row(cntjj) = ""
                            End If
                            insertvalue = insertvalue & "," & "'" & row(cntjj).ToString.Replace("'", "’") & "'"
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                        Next

                        'データ部以降から挿入処理
                        If readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                            '挿入用に成形
                            insertvalue = insertvalue.Remove(0, 1)
                            insertvalue = "'" & cntii.ToString & "'" & "," & insertvalue

                            '挿入処理
                            Dim tmp_sql_insert As String = " INSERT INTO " & tmptblname & " VALUES(" & insertvalue & ")"
                            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_insert, tmp_cnt)

                        End If

                        '----- プログレスバー更新/進捗率表示 -----
                        Dim tmp_pgbcnt As Integer = 0
                        If dt.Rows.Count <= pgbbasecnt Then             '基準件数(100件)以下の場合は実件数を取得
                            tmp_pgbcnt = readrowcnt
                        ElseIf dt.Rows.Count > readrowcnt Then          '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                        ElseIf dt.Rows.Count <= readrowcnt Then         '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            tmp_pgbcnt = pgbtotalcnt_part
                        End If

                        '表示
                        If tmp_pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                            Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                        End If

                        readrowcnt = readrowcnt + 1

                    Next

                Next

            Catch ex As Exception

                'エラー処理
                rtn = False
                errstr = MSG_ERR_BASEMIDFILEREAD

                '----- ログ出力 -----              
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

                Return rtn

            End Try

            Return rtn

        End Function
        '20161028 物件/部屋鍵取得方法修正 -add end

        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        '通常コピーと個別コピーの統合に伴い、もともとあった処理をコメントアウト
        ' ''' <summary>
        ' ''' 汎用用中間ファイル→既存用中間ファイル個別処理 '20160929 汎用→既存コピー処理改善対応
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <param name="list_cv"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidFile_ChgCopyMain(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim obj_rep As New Object
        '    Dim model_basemiditem As New Object
        '    Dim basemidfilename As String = ""
        '    Dim basemidfilesheetname As String = ""
        '    Dim existmidfilename As String = ""
        '    Dim existmidfilesheetname As String = ""

        '    Dim tmptmpstr As String = ""                                '20161009 ログ出力処理追加 -add
        '    Dim tmptmpcnt As Long                                       '20161009 ログ出力処理追加 -add

        '    '----- 全体用・中間用プログレスバー更新 -----
        '    Dim obj_pgb As New ProgressBarManager
        '    Dim pgbcnt As Integer = 0
        '    Dim pgbtotalcnt_total As Integer = list_cv.Count
        '    Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
        '    Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)


        '    For Each item In list_cv

        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then
        '            Return rtn
        '        End If

        '        Dim cvitem As String = item

        '        'コピー項目表示
        '        Me.lblCVItem.Text = ""

        '        '----- デバッグ用処理 ----- sta
        '        'Dim list As New List(Of String) From {"部屋情報-部屋特約情報"}
        '        'If list.Contains(cvitem) = False Then
        '        '    cvitem = ""
        '        'End If
        '        '----- デバッグ用処理 ----- end

        '        Dim normalflg As Boolean = True
        '        Dim qry As String = Nothing
        '        Dim copyflg As Boolean = True

        '        Select Case cvitem
        '            Case "部屋情報-部屋駐車場情報"
        '                obj_rep = New Njc.Repository.Hy_cyusyajo_mid_Repository
        '                pgbcnt = pgbcnt + 1
        '                Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        '            Case "部屋情報-部屋特約情報"
        '                obj_rep = New Njc.Repository.Hy_tokuyaku_mid_Repository
        '                pgbcnt = pgbcnt + 1
        '                Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        '            Case "契約情報-契約契約者情報"
        '                obj_rep = New Njc.Repository.Ky_kys_mid_Repository
        '                pgbcnt = pgbcnt + 1
        '                Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        '            Case "契約情報-契約保証人情報"
        '                obj_rep = New Njc.Repository.Ky_hosyonin_mid_Repository
        '                pgbcnt = pgbcnt + 1
        '                Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        '            Case "契約情報-契約特約事項情報"
        '                obj_rep = New Njc.Repository.Ky_tokuyaku_mid_Repository
        '                pgbcnt = pgbcnt + 1
        '                Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        '            Case Else
        '                copyflg = False
        '        End Select

        '        '----- ログ出力 -----                       '20161009 ログ出力処理追加 -add
        '        Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
        '        tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, LOG_SYORIKOMK_MIDFILE, cvitem), False)
        '        DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

        '        '項目毎に個別処理開始
        '        If copyflg Then
        '            '挿入処理
        '            rtn = obj_rep.Read_BaseMidFile()
        '        End If

        '        'コピー処理失敗時は処理を抜ける
        '        If rtn = False Then
        '            errstr = MSG_ERR_BASEMIDFILEREAD

        '            '----- ログ出力 -----                   '20161009 ログ出力処理追加 -add
        '            Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
        '            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
        '            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
        '            Return rtn
        '        End If

        '        '----- 全体用プログレスバー更新 -----
        '        Call obj_pgb.pgbsettingTotal(pgbcnt)
        '        Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        '        '----- 中間用プログレスバー更新 -----
        '        Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        '        Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        '        Me.Refresh()
        '    Next

        '    If pgbcnt < pgbtotalcnt_total Then
        '        '----- 全体用プログレスバー更新(最終調整) -----
        '        pgbcnt = pgbtotalcnt_total
        '        Call obj_pgb.pgbsettingTotal(pgbcnt)
        '        Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        '        '----- 中間用プログレスバー更新(最終調整) -----
        '        Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        '        Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        '        Me.Refresh()
        '    End If

        '    '返却
        '    Return rtn

        'End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        ''' <summary>
        ''' 既存用中間ファイルのヘッダー名をオブジェクトへ格納 '20160812 汎用コンバート対応
        ''' '20160905 中間ファイルコピー処理改善 引数を追加
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="sheetname"></param>
        ''' <remarks></remarks>
        Private Sub Set_MidDataInfoToObj(ByVal filename As String, ByVal sheetname As String, ByVal getbasemidflg As Boolean)

            Dim readtbl As New DataTable
            Dim reccnt As Integer = 0
            Dim fldname As String = ""                                              '20160928 レビュー指摘修正 変数宣言時に初期化処理を行う
            Dim fldvalue As String = ""
            Dim tmp_既存中間ファイル名 As String = ""
            Dim tmp_既存中間シート名 As String = ""
            Dim tmp_既存中間フィールド名 As String = ""
            Dim tmp_汎用中間ファイル名 As String = ""
            Dim tmp_汎用中間シート名 As String = ""
            Dim tmp_汎用中間フィールド名_大分類 As String = ""
            Dim tmp_汎用中間フィールド名_小分類 As String = ""
            Dim tmp_コピーフラグ As String = ""
            Dim tmp_データ型 As String = ""
            Dim tmp_CV用最小値 As String = ""
            Dim tmp_CV用最大値 As String = ""
            Dim tmp_デフォルト値 As String = ""
            Dim tmp_CV用キーNo As String = ""
            Dim tmp_CV用必須項目フラグ As String = ""
            Dim tmp_有無参照 As String = ""
            Dim tmp_移行対象フラグ As String = ""                                   '20160926 選定した移行項目をプログラムへ反映する修正 -add


            '--------------------------------------------------
            ' 中間ファイル情報取得
            '--------------------------------------------------
            '20160905 中間ファイルコピー処理改善 -chg sta
            'Dim tmp_sql As String = " SELECT DISTINCT * FROM " & MIDFILEINFO_DBNAME
            Dim tmp_sql As String = ""
            Dim tmp_taisyofname As String = ""
            Dim tmp_taisyosname As String = ""
            Dim tmp_cvflg As String = "移行対象フラグ"                              '20160929 汎用→既存コピー処理改善対応 -add
            If getbasemidflg Then
                tmp_taisyofname = "汎用中間ファイル名"
                tmp_taisyosname = "汎用中間シート名"
            Else
                tmp_taisyofname = "既存中間ファイル名"
                tmp_taisyosname = "既存中間シート名"
            End If
            tmp_sql = tmp_sql & " SELECT DISTINCT * FROM " & MIDFILEINFO_DBNAME
            tmp_sql = tmp_sql & " WHERE RTRIM(" & tmp_taisyosname & ") = '" & sheetname & "' "
            tmp_sql = tmp_sql & " AND RTRIM(" & tmp_cvflg & ") = '●' "             '20160929 汎用→既存コピー処理改善対応 -add
            '20160905 中間ファイルコピー処理改善 -chg end
            reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl)

            '--------------------------------------------------
            ' 初期化
            '--------------------------------------------------
            'List_Existmidheader.Clear()                                             '20160812 汎用コンバート対応 -del
            Hash_ExistMiddatatype.Clear()
            Hash_ExistMiddatamin.Clear()
            Hash_ExistMiddatamax.Clear()
            Hash_ExistMiddatadef.Clear()
            Hash_ExistMiddatakey.Clear()
            List_ExistMiddatareq.Clear()
            Hash_ExistMiddataref.Clear()
            List_ExistMiddatacv.Clear()                                             '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -add
            'List_Basemidheader.Clear()                                              '20160812 汎用コンバート対応 -del
            Hash_BaseMiddatatype.Clear()
            Hash_BaseMiddatamin.Clear()
            Hash_BaseMiddatamax.Clear()
            Hash_BaseMiddatadef.Clear()
            Hash_BaseMiddatakey.Clear()
            List_BaseMiddatareq.Clear()
            Hash_BaseMiddataref.Clear()
            List_BaseMiddatacv.Clear()                                              '20160926 選定した移行項目をプログラムへ反映する修正 -add
            Hash_BaseMidToExistMid_FS.Clear()                                       '20160905 中間ファイルコピー処理改善 -add
            Hash_ExistMidToBaseMid_FS.Clear()                                       '20160905 中間ファイルコピー処理改善 -add
            Hash_BaseMidToExistMid_SH.Clear()                                       '20160905 中間ファイルコピー処理改善 -add
            Hash_ExistMidToBaseMid_SH.Clear()                                       '20160905 中間ファイルコピー処理改善 -add

            '--------------------------------------------------
            ' 読込開始
            '--------------------------------------------------
            For cntii As Integer = 0 To readtbl.Rows.Count - 1

                '中断処理
                Application.DoEvents()
                If CancelFlg Or MidChkCancelFlg Then
                    Exit Sub
                End If

                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                    '項目名取得
                    fldname = readtbl.Columns(cntjj).ColumnName

                    '登録値取得
                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                    '各項目値→変数格納
                    Select Case fldname
                        Case "既存中間ファイル名"
                            tmp_既存中間ファイル名 = fldvalue.Trim
                        Case "既存中間シート名"
                            tmp_既存中間シート名 = fldvalue.Trim
                        Case "既存中間フィールド名"
                            tmp_既存中間フィールド名 = fldvalue.Trim
                        Case "汎用中間ファイル名"
                            tmp_汎用中間ファイル名 = fldvalue.Trim
                        Case "汎用中間シート名"
                            tmp_汎用中間シート名 = fldvalue.Trim
                        Case "汎用中間フィールド名_大分類"
                            tmp_汎用中間フィールド名_大分類 = fldvalue.Trim
                        Case "汎用中間フィールド名_小分類"
                            tmp_汎用中間フィールド名_小分類 = fldvalue.Trim
                        Case "コピーフラグ"
                            tmp_コピーフラグ = fldvalue.Trim
                        Case "データ型"
                            tmp_データ型 = fldvalue.Trim
                        Case "CV用最小値"
                            tmp_CV用最小値 = fldvalue.Trim
                        Case "CV用最大値"
                            tmp_CV用最大値 = fldvalue.Trim
                        Case "デフォルト値"
                            tmp_デフォルト値 = fldvalue.Trim
                        Case "CV用キーNo"
                            tmp_CV用キーNo = fldvalue.Trim
                        Case "CV用必須項目フラグ"
                            tmp_CV用必須項目フラグ = fldvalue.Trim
                        Case "有無参照"
                            tmp_有無参照 = fldvalue.Trim
                        Case "移行対象フラグ"
                            tmp_移行対象フラグ = fldvalue.Trim
                    End Select
                Next

                '--------------------------------------------------
                ' 汎用用と既存用の紐付(ファイル名-シート名)                         '20160905 中間ファイルコピー処理改善 -add
                '--------------------------------------------------
                Dim tmp_basemiddatatotal As String = tmp_汎用中間ファイル名 & STR_SPLIT_1 & tmp_汎用中間シート名
                Dim tmp_existmiddatatotal As String = tmp_既存中間ファイル名 & STR_SPLIT_1 & tmp_既存中間シート名
                If tmp_basemiddatatotal = STR_SPLIT_1 Then
                    tmp_basemiddatatotal = ""
                End If
                If tmp_existmiddatatotal = STR_SPLIT_1 Then
                    tmp_existmiddatatotal = ""
                End If

                If tmp_basemiddatatotal <> "" And tmp_existmiddatatotal <> "" And Hash_BaseMidToExistMid_FS.Contains(tmp_basemiddatatotal) = False Then
                    Hash_BaseMidToExistMid_FS.Add(tmp_basemiddatatotal, tmp_existmiddatatotal)
                End If
                If tmp_existmiddatatotal <> "" And tmp_basemiddatatotal <> "" And Hash_ExistMidToBaseMid_FS.Contains(tmp_existmiddatatotal) = False Then
                    Hash_ExistMidToBaseMid_FS.Add(tmp_existmiddatatotal, tmp_basemiddatatotal)
                End If

                '--------------------------------------------------
                ' 既存用中間ファイルと紐付けて各オブジェクトへ格納
                '--------------------------------------------------
                Dim tmp_existmidheader As String = tmp_既存中間シート名 & STR_SPLIT_1 & tmp_既存中間フィールド名
                If tmp_existmidheader = STR_SPLIT_1 Then
                    tmp_existmidheader = ""
                End If
                '20160905 中間ファイルコピー処理改善 -del sta
                ''既存用中間ファイルのヘッダー
                'If tmp_existmidheader <> "" Then
                '    List_Existmidheader.Add(tmp_existmidheader)
                'End If
                '20160905 中間ファイルコピー処理改善 -del end
                '既存用中間ファイルデータ型
                If tmp_existmidheader <> "" Then
                    Hash_ExistMiddatatype.Add(tmp_existmidheader, tmp_データ型)
                End If

                '既存用中間ファイル最小値
                If tmp_existmidheader <> "" Then
                    Hash_ExistMiddatamin.Add(tmp_existmidheader, tmp_CV用最小値)
                End If

                '既存用中間ファイル最大値
                If tmp_existmidheader <> "" Then
                    Hash_ExistMiddatamax.Add(tmp_existmidheader, tmp_CV用最大値)
                End If

                '既存用中間ファイルデフォルト値
                If tmp_existmidheader <> "" And tmp_デフォルト値 <> "" Then
                    Hash_ExistMiddatadef.Add(tmp_existmidheader, tmp_デフォルト値)
                End If

                '既存用中間ファイルキーNo
                If tmp_existmidheader <> "" And tmp_CV用キーNo <> "" Then
                    Hash_ExistMiddatakey.Add(tmp_existmidheader, tmp_CV用キーNo)
                End If

                '既存用中間ファイル必須項目フラグ
                If tmp_existmidheader <> "" And tmp_CV用必須項目フラグ <> "" Then
                    List_ExistMiddatareq.Add(tmp_existmidheader)
                End If

                '既存用中間ファイル有無参照
                If tmp_existmidheader <> "" And tmp_有無参照 <> "" Then
                    Hash_ExistMiddataref.Add(tmp_existmidheader, tmp_有無参照)
                End If

                '移行対象フラグ                                                    '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -add
                If tmp_existmidheader <> "" And tmp_移行対象フラグ = "●" Then
                    List_ExistMiddatacv.Add(tmp_existmidheader)
                End If

                '--------------------------------------------------
                ' 汎用用中間ファイルと紐付けて各オブジェクトへ格納
                '--------------------------------------------------
                Dim tmp_basemidheader As String = ""
                If tmp_汎用中間フィールド名_小分類 <> "" Then
                    '20160913_2 部屋設備移行処理の追加 -chg sta
                    'tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_小分類
                    tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_大分類 & STR_SPLIT_1 & tmp_汎用中間フィールド名_小分類
                    '20160913_2 部屋設備移行処理の追加 -chg end
                Else
                    tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_大分類
                End If
                If tmp_basemidheader = STR_SPLIT_1 Then
                    tmp_basemidheader = ""
                End If
                '20160905 中間ファイルコピー処理改善 -del sta
                ''汎用用中間ファイルのヘッダー
                'If tmp_basemidheader <> "" And List_Basemidheader.Contains(tmp_basemidheader) = False Then
                '    List_Basemidheader.Add(tmp_basemidheader)
                'End If
                '20160905 中間ファイルコピー処理改善 -del end
                '汎用用中間ファイルデータ型
                If tmp_basemidheader <> "" And Hash_BaseMiddatatype.Contains(tmp_basemidheader) = False Then
                    Hash_BaseMiddatatype.Add(tmp_basemidheader, tmp_データ型)
                End If

                '汎用用中間ファイル最小値
                If tmp_basemidheader <> "" And Hash_BaseMiddatamin.Contains(tmp_basemidheader) = False Then
                    Hash_BaseMiddatamin.Add(tmp_basemidheader, tmp_CV用最小値)
                End If

                '汎用用中間ファイル最大値
                If tmp_basemidheader <> "" And Hash_BaseMiddatamax.Contains(tmp_basemidheader) = False Then
                    Hash_BaseMiddatamax.Add(tmp_basemidheader, tmp_CV用最大値)
                End If

                '汎用用中間ファイルデフォルト値
                If tmp_basemidheader <> "" And tmp_デフォルト値 <> "" And Hash_BaseMiddatadef.Contains(tmp_basemidheader) = False Then
                    Hash_BaseMiddatadef.Add(tmp_basemidheader, tmp_デフォルト値)
                End If

                '汎用用中間ファイルキーNo
                If tmp_basemidheader <> "" And tmp_CV用キーNo <> "" And Hash_BaseMiddatakey.Contains(tmp_basemidheader) = False Then
                    Hash_BaseMiddatakey.Add(tmp_basemidheader, tmp_CV用キーNo)
                End If

                '汎用用中間ファイル必須項目フラグ
                If tmp_basemidheader <> "" And tmp_CV用必須項目フラグ <> "" And List_BaseMiddatareq.Contains(tmp_basemidheader) = False Then
                    List_BaseMiddatareq.Add(tmp_basemidheader)
                End If

                '汎用用中間ファイル有無参照
                If tmp_basemidheader <> "" And tmp_有無参照 <> "" And Hash_BaseMiddataref.Contains(tmp_basemidheader) = False Then
                    Hash_BaseMiddataref.Add(tmp_basemidheader, tmp_有無参照)
                End If

                '移行対象フラグ    '20160926 選定した移行項目をプログラムへ反映する修正 -add
                If tmp_basemidheader <> "" And tmp_移行対象フラグ = "●" Then
                    List_BaseMiddatacv.Add(tmp_basemidheader)
                End If

                '--------------------------------------------------
                '汎用用と既存用の紐付(シート名-ヘッダー名)                         '20160905 中間ファイルコピー処理改善 -add
                '--------------------------------------------------
                If tmp_existmidheader <> "" And Hash_ExistMidToBaseMid_SH.Contains(tmp_existmidheader) = False Then
                    Hash_ExistMidToBaseMid_SH.Add(tmp_existmidheader, tmp_basemidheader)
                End If
                If tmp_basemidheader <> "" And Hash_BaseMidToExistMid_SH.Contains(tmp_basemidheader) = False Then
                    Hash_BaseMidToExistMid_SH.Add(tmp_basemidheader, tmp_existmidheader)
                End If

            Next

        End Sub

        ''' <summary>
        ''' 契約No、管理Noに「1」を自動設定する '20160929 汎用→既存コピー処理改善対応 -add
        ''' </summary>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_ExistMidFile_DefCopyMain(ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim list_file As New List(Of String) From {"契約情報", "送金ルール情報", "物件情報", "部屋情報"}
            Dim qry As String = Nothing
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
            'Dim da As New OleDbDataAdapter()
            'Dim ds As DataSet = New DataSet()
            'Dim dt As New DataTable()
            'Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            'Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            'Dim con_read As New OleDbConnection()
            'Dim cmd_read As New OleDbCommand()
            'Dim con_write As New OleDbConnection()
            'Dim cmd_write As New OleDbCommand()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
            Dim tmp_cnt As Integer = 0

            '----- 全体用・中間用プログレスバー更新 -----
            Dim obj_pgb As New ProgressBarManager
            Dim pgbcnt As Integer = 0
            Dim pgbtotalcnt_total As Integer = list_file.Count
            Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
            Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)
            Dim pgbcnt_part As Integer = 0

            '更新クエリ
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            'Dim tmpqry_ky As String()
            'qry = " UPDATE [契約基本情報$] SET [契約No] = 1"
            'qry = qry & "@#@" & " UPDATE [契約履歴情報$] SET [契約No] = 1,[契約管理レコードNo] = 1,[更新No] = 1,[改定No] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約契約者情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約入居者情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約保証人情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約車情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約保険情報$] SET [契約No] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約特約事項情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約メモ情報$] SET [契約NO] = 1,[契約管理レコードNo] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約入金項目情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            'qry = qry & "@#@" & " UPDATE [契約次回入金項目情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            'tmpqry_ky = Split(qry, "@#@")
            'Dim tmpqry_so As String()
            'qry = " UPDATE [送金ルール基本情報$] SET [送金ルール管理No] = 1 "
            'qry = qry & "@#@" & " UPDATE [送金ルール送金先情報$] SET [送金ルール管理No] = 1 "
            'qry = qry & "@#@" & " UPDATE [送金ルール入金項目情報$] SET [送金ルール管理No] = 1 "
            'qry = qry & "@#@" & " UPDATE [送金ルール控除項目情報$] SET [送金ルール管理No] = 1 "
            'tmpqry_so = Split(qry, "@#@")
            'Dim tmpqry_bk As String
            'tmpqry_bk = " UPDATE [物件所有者情報$] SET [管理No] = 1 "
            'Dim tmpqry_hy As String
            'tmpqry_hy = " UPDATE [部屋所有者情報$] SET [管理No] = 1 "
            Dim tmpqry_ky As String()
            qry = " UPDATE [CVTBL_契約基本情報] SET [契約No] = 1"
            qry = qry & "@#@" & " UPDATE [CVTBL_契約履歴情報] SET [契約No] = 1,[契約管理レコードNo] = 1,[更新No] = 1,[改定No] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約契約者情報] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約入居者情報] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約保証人情報] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約車情報] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約保険情報] SET [契約No] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約特約事項情報] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約メモ情報] SET [契約NO] = 1,[契約管理レコードNo] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約入金項目情報] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_契約次回入金項目情報] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            tmpqry_ky = Split(qry, "@#@")
            Dim tmpqry_so As String()
            qry = " UPDATE [CVTBL_送金ルール基本情報] SET [送金ルール管理No] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_送金ルール送金先情報] SET [送金ルール管理No] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_送金ルール入金項目情報] SET [送金ルール管理No] = 1 "
            qry = qry & "@#@" & " UPDATE [CVTBL_送金ルール控除項目情報] SET [送金ルール管理No] = 1 "
            tmpqry_so = Split(qry, "@#@")
            Dim tmpqry_bk As String
            tmpqry_bk = " UPDATE [CVTBL_物件所有者情報] SET [管理No] = 1 "
            Dim tmpqry_hy As String
            tmpqry_hy = " UPDATE [CVTBL_部屋所有者情報] SET [管理No] = 1 "
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
            '件数取得
            Dim pgbtotalcnt_part As Integer = tmpqry_ky.Count + tmpqry_so.Count + 1 + 1
            Call obj_pgb.pgbInitPart(pgbtotalcnt_part)
        

            Try
                For Each filename In list_file

                    Me.lblCVItem.Text = filename                                            'コピー項目表示
                    Dim existmidpath As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")

                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                    ''--------------------------------------------------
                    '' EXCEL接続 (ファイル書込用)
                    ''--------------------------------------------------
                    ''接続文字列生成
                    'con_write.ConnectionString = _
                    '    "Provider=" & tmpprovider & _
                    '    "Data Source=" & existmidpath & ";" & _
                    '    "Extended Properties=" & """" & tmpextend & """"
                    ''接続設定
                    'cmd_write.Connection = con_write

                    ''--------------------------------------------------
                    '' 読込→書込
                    ''--------------------------------------------------
                    ''接続オープン処理
                    'con_write.Open()
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                    'クエリ実行
                    Select Case filename
                        Case "契約情報"
                            For tmpcnt = 0 To UBound(tmpqry_ky)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                'cmd_write.CommandText = tmpqry_ky(tmpcnt)
                                'cmd_write.ExecuteNonQuery()
                                Dim tmp_qry As String = tmpqry_ky(tmpcnt)
                                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_qry, tmp_cnt)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                                '----- 個別用プログレスバー更新 -----
                                pgbcnt_part = pgbcnt_part + 1
                                Call obj_pgb.pgbsettingPart(pgbcnt_part)
                                Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
                                '20161017 進捗表示処理による速度低下の修正 -del
                                'Me.Refresh()
                            Next
                        Case "送金ルール情報"
                            For tmpcnt = 0 To UBound(tmpqry_so)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                'cmd_write.CommandText = tmpqry_so(tmpcnt)
                                'cmd_write.ExecuteNonQuery()
                                Dim tmp_qry As String = tmpqry_so(tmpcnt)
                                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_qry, tmp_cnt)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                                '----- 個別用プログレスバー更新 -----
                                pgbcnt_part = pgbcnt_part + 1
                                Call obj_pgb.pgbsettingPart(pgbcnt_part)
                                Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
                                '20161017 進捗表示処理による速度低下の修正 -del
                                'Me.Refresh()
                            Next
                        Case "物件情報"
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                            'cmd_write.CommandText = tmpqry_bk
                            'cmd_write.ExecuteNonQuery()
                            Dim tmp_qry As String = tmpqry_bk
                            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_qry, tmp_cnt)
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                            '----- 個別用プログレスバー更新 -----
                            pgbcnt_part = pgbcnt_part + 1
                            Call obj_pgb.pgbsettingPart(pgbcnt_part)
                            Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
                            '20161017 進捗表示処理による速度低下の修正 -del
                            'Me.Refresh()
                        Case "部屋情報"
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                            'cmd_write.CommandText = tmpqry_hy
                            'cmd_write.ExecuteNonQuery()
                            Dim tmp_qry As String = tmpqry_hy
                            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_qry, tmp_cnt)
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                            '----- 個別用プログレスバー更新 -----
                            pgbcnt_part = pgbcnt_part + 1
                            Call obj_pgb.pgbsettingPart(pgbcnt_part)
                            Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
                            '20161017 進捗表示処理による速度低下の修正 -del
                            'Me.Refresh()
                        Case Else

                    End Select

                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                    ''接続クローズ処理
                    'con_write.Close()
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                    '----- 全体用・中間用プログレスバー更新 -----
                    pgbcnt = pgbcnt + 1
                    Call obj_pgb.pgbsettingTotal(pgbcnt)
                    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
                    Call obj_pgb.pgbsettingChkTotal(pgbcnt)
                    Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
                    Me.Refresh()
                Next

            Catch ex As Exception

                'エラー処理
                errstr = MSG_ERR_BASEMIDFILEREAD
                rtn = False
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                ''クローズ処理
                'con_write.Close()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
            End Try

            '20161026 移行項目画面表示修正 -add
            Me.lblCVItem.Text = ""

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 一棟区分を所有情報から判断して自動設定する '20160929 汎用→既存コピー処理改善対応 -add
        ''' '20161011 汎用→既存書込時の不要処理除去修正 引数に「list_filename」を追加
        ''' </summary>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_ExistMidFile_SyoyuCopyMain(ByRef errstr As String, ByVal list_filename As List(Of String)) As Boolean

            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            'Dim rtn As Boolean = True
            ''20161011 一棟/区分一括調整処理改善 -chg sta
            ''Dim list_ittobkno As New List(Of String)    
            ''Dim list_kbnbkno As New List(Of String)         
            'Dim solist_kbnbkno As New SortedList(Of Integer, String)
            ''20161011 一棟/区分一括調整処理改善 -chg end

            ''20161011 汎用→既存書込時の不要処理除去修正 -add sta
            ''物件情報、部屋情報どちらも移行対象ではない場合は処理を抜ける
            'If list_filename.Contains("物件情報") = False And list_filename.Contains("部屋情報") = False Then
            '    Return rtn
            'End If
            ''20161011 汎用→既存書込時の不要処理除去修正 -add end

            ''物件No取得
            ''20161011 一棟/区分一括調整処理改善 -chg sta
            ''最初に全て一棟に設定した後で区分のみ更新するように修正するため一棟用のオブジェクトを削除
            ''rtn = Set_ExistMidFile_Syoyubkno(list_ittobkno, list_kbnbkno, errstr)
            'rtn = Set_ExistMidFile_Syoyubkno(solist_kbnbkno, errstr)
            ''20161011 一棟/区分一括調整処理改善 -chg end
            'If rtn Then
            '    '取得した物件Noから一棟/区分を設定
            '    '20161011 一棟/区分一括調整処理改善 -chg sta
            '    '最初に全て一棟に設定した後で区分のみ更新するように修正するため一棟用のオブジェクトを削除
            '    'rtn = Set_ExistMidFile_Syoyukbn(list_ittobkno, list_kbnbkno, errstr)
            '    rtn = Set_ExistMidFile_Syoyukbn(solist_kbnbkno, errstr)
            '    '20161011 一棟/区分一括調整処理改善 -chg end
            'End If

            'Return rtn

            Dim rtn As Boolean = True
            Dim tmp_sql As String = ""
            Dim tmp_cnt As Integer = 0

            '実行有無チェック
            If list_filename.Contains("物件情報") = False And list_filename.Contains("部屋情報") = False Then
                Return rtn
            End If

            Try

                '----------------------------------
                'あらかじめ一棟所有へ更新しておく
                '----------------------------------
                '物件詳細情報
                tmp_sql = " UPDATE CVTBL_物件詳細情報 SET [所有者-一棟・所有区分] = 1 "
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql, tmp_cnt)
                tmp_sql = ""

                '送金ルール基本情報
                tmp_sql = " UPDATE CVTBL_送金ルール基本情報 SET [一所有形態区分-棟/区分] = 1 "
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql, tmp_cnt)
                tmp_sql = ""

                '----------------------------------
                '区分所有を更新
                '----------------------------------
                '物件詳細情報
                tmp_sql = tmp_sql & " UPDATE CVTBL_物件詳細情報 SET [所有者-一棟・所有区分] = 2 "
                tmp_sql = tmp_sql & " WHERE [物件NO] IN "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT [物件No] FROM CVTBL_部屋所有者情報 "
                tmp_sql = tmp_sql & " 	) "
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql, tmp_cnt)
                tmp_sql = ""

                '送金ルール基本情報
                tmp_sql = tmp_sql & " UPDATE CVTBL_送金ルール基本情報 SET [一所有形態区分-棟/区分] = 2 "
                tmp_sql = tmp_sql & " WHERE [物件No] IN "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT [物件No] FROM CVTBL_部屋所有者情報 "
                tmp_sql = tmp_sql & " 	) "
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql, tmp_cnt)
                tmp_sql = ""

            Catch ex As Exception

                errstr = MSG_ERR_BASEMIDFILEREAD
                rtn = False

            End Try
            
            Return rtn
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
        End Function

        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        ' ''' <summary>
        ' ''' 中間ファイルの物件/部屋所有者情報から物件Noを取得してオブジェクトへ格納 '20160929 汎用→既存コピー処理改善対応
        ' ''' '20161011 一棟/区分一括調整処理改善
        ' ''' 引数「ByRef list_ittobkno As List(Of String)」を削除
        ' ''' 引数「ByRef list_kbnbkno As List(Of String)」を「ByRef solist_kbnbkno As SortedList(Of Integer, String)」へ変更
        ' ''' →最初に全て一棟に設定した後で区分のみ更新するように修正する
        ' ''' </summary>
        ' ''' <param name="solist_kbnbkno"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_ExistMidFile_Syoyubkno(ByRef solist_kbnbkno As SortedList(Of Integer, String), ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim da As New OleDbDataAdapter()
        '    Dim ds As DataSet = New DataSet()
        '    Dim dt As New DataTable()
        '    Dim dr As OleDbDataReader
        '    Dim qry As String = Nothing

        '    Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        '    Dim tmpextend As String = "Excel 8.0;HDR=YES;"
        '    '20161011 一棟/区分一括調整処理改善 -chg sta
        '    'Dim list_filename As New List(Of String) From {"物件情報", "部屋情報"}
        '    Dim filename As String = "部屋情報"
        '    Dim sheetname As String = "部屋所有者情報"
        '    '20161011 一棟/区分一括調整処理改善 -chg end


        '    '20161011 一棟/区分一括調整処理改善 -chg sta
        '    ''物件/部屋所有者情報から対象物件Noを取得
        '    'For Each filename In list_filename

        '    '    '中断処理
        '    '    Application.DoEvents()
        '    '    If CancelFlg Or MidChkCancelFlg Then
        '    '        Return rtn
        '    '    End If

        '    '    'ファイルパス設定
        '    '    Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")

        '    '    '------------------
        '    '    'メイン処理
        '    '    '------------------

        '    '    Dim con_read As New OleDbConnection()
        '    '    Dim cmd_read As New OleDbCommand()
        '    '    Dim con_write As New OleDbConnection()
        '    '    Dim cmd_write As New OleDbCommand()

        '    '    Try

        '    '        '================================================================================
        '    '        ' ●接続設定●
        '    '        '================================================================================
        '    '        '--------------------------------------------------
        '    '        ' ①EXCEL接続 (ファイル読込用)
        '    '        '--------------------------------------------------

        '    '        '接続文字列生成
        '    '        con_read.ConnectionString = _
        '    '            "Provider=" & tmpprovider & _
        '    '            "Data Source=" & existmidpath_read & ";" & _
        '    '            "Extended Properties=" & """" & tmpextend & """"

        '    '        '接続設定
        '    '        cmd_read.Connection = con_read

        '    '        '--------------------------------------------------
        '    '        ' 読込
        '    '        '--------------------------------------------------
        '    '        Dim sheetname As String = ""
        '    '        Select Case filename
        '    '            Case "物件情報"
        '    '                sheetname = "物件所有者情報"
        '    '            Case "部屋情報"
        '    '                sheetname = "部屋所有者情報"
        '    '        End Select

        '    '        '接続オープン処理
        '    '        con_read.Open()
        '    '        '20161011 一棟/区分一括調整処理改善 -chg sta
        '    '        'qry = "SELECT [物件No] FROM [" & sheetname & "$] "            '20161009 不具合調査：クエリが複雑すぎる。DISTINCTできない？部屋の場合に同内容の大量のデータを格納することになる
        '    '        qry = "SELECT DISTINCT [物件No] FROM [" & sheetname & "$] "
        '    '        '20161011 一棟/区分一括調整処理改善 -chg end
        '    '        cmd_read = con_read.CreateCommand
        '    '        cmd_read.CommandText = qry
        '    '        dt = New DataTable
        '    '        dr = cmd_read.ExecuteReader
        '    '        dt.Load(dr)

        '    '        Dim colcnt As Integer = dt.Columns.Count
        '    '        Dim readrowcnt As Integer = 1
        '    '        Dim solist_header As New SortedList(Of Integer, String)
        '    '        Dim tmp_list As New List(Of String)

        '    '        '1行ずつ取得
        '    '        Dim row As DataRow
        '    '        For Each row In dt.Rows
        '    '            tmp_list.Add(row(0).ToString)
        '    '        Next

        '    '        Select Case filename
        '    '            Case "物件情報"
        '    '                list_ittobkno = tmp_list
        '    '            Case "部屋情報"
        '    '                list_kbnbkno = tmp_list
        '    '        End Select

        '    '    Catch ex As Exception

        '    '        'エラー処理
        '    '        errstr = MSG_ERR_BASEMIDFILEREAD
        '    '        rtn = False

        '    '        'クローズ処理
        '    '        con_read.Close()
        '    '        con_write.Close()

        '    '    End Try

        '    '    'クローズ処理
        '    '    con_read.Close()
        '    '    con_write.Close()

        '    'Next

        '    '中断処理
        '    Application.DoEvents()
        '    If CancelFlg Or MidChkCancelFlg Then
        '        Return rtn
        '    End If

        '    'ファイルパス設定
        '    Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")

        '    '--------------------------------------------------
        '    ' メイン処理
        '    '--------------------------------------------------
        '    Dim con_read As New OleDbConnection()
        '    Dim cmd_read As New OleDbCommand()
        '    Dim con_write As New OleDbConnection()
        '    Dim cmd_write As New OleDbCommand()

        '    Try
        '        '--------------------------------------------------
        '        ' ①EXCEL接続 (ファイル読込用)
        '        '--------------------------------------------------
        '        '接続文字列生成
        '        con_read.ConnectionString = _
        '            "Provider=" & tmpprovider & _
        '            "Data Source=" & existmidpath_read & ";" & _
        '            "Extended Properties=" & """" & tmpextend & """"
        '        '接続設定
        '        cmd_read.Connection = con_read

        '        '--------------------------------------------------
        '        ' 物件/部屋所有者情報から対象物件Noを取得
        '        '--------------------------------------------------
        '        con_read.Open()
        '        '20161011 一棟/区分一括調整処理改善 -chg sta
        '        'qry = "SELECT [物件No] FROM [" & sheetname & "$] "                   '20161009 不具合調査：クエリが複雑すぎる為エラー(in連結の許容範囲オーバー)
        '        qry = "SELECT DISTINCT [物件No] FROM [" & sheetname & "$] "
        '        '20161011 一棟/区分一括調整処理改善 -chg end
        '        cmd_read = con_read.CreateCommand
        '        cmd_read.CommandText = qry
        '        dt = New DataTable
        '        dr = cmd_read.ExecuteReader
        '        dt.Load(dr)

        '        Dim colcnt As Integer = dt.Columns.Count
        '        Dim readrowcnt As Integer = 1
        '        Dim solist_header As New SortedList(Of Integer, String)
        '        Dim tmp_list As New List(Of String)

        '        '1行ずつ取得
        '        Dim row As DataRow
        '        Dim cnt As Integer = 0
        '        Dim cnttotal As Integer = 0
        '        Dim cnt_kugiri As Integer = 10000
        '        Dim tmp_str As String = ""

        '        For Each row In dt.Rows

        '            'カウンター更新
        '            cnt = cnt + 1
        '            cnttotal = cnttotal + 1

        '            'sql用の文字列へ変換
        '            tmp_str = tmp_str & "," & "'" & row(0).ToString & "'"

        '            If cnttotal = dt.Rows.Count Or cnt = cnt_kugiri Then
        '                'sql用文字列成形
        '                If tmp_str <> "" Then
        '                    tmp_str = tmp_str.Remove(0, 1)
        '                End If

        '                'オブジェクトへ格納
        '                solist_kbnbkno.Add(Math.Ceiling(cnttotal / cnt_kugiri), tmp_str)

        '                'カウンター、作業用文字列初期化
        '                cnt = 0
        '                tmp_str = ""
        '            End If
        '        Next

        '    Catch ex As Exception

        '        'エラー処理
        '        errstr = MSG_ERR_BASEMIDFILEREAD
        '        rtn = False
        '    End Try

        '    'クローズ処理
        '    con_read.Close()
        '    con_write.Close()
        '    '20161011 一棟/区分一括調整処理改善 -chg end

        '    '返却
        '    Return rtn

        'End Function

        ' ''' <summary>
        ' ''' 所有者情報から取得した物件Noを元に一棟・区分を自動設定 '20160929 汎用→既存コピー処理改善対応
        ' ''' 物件詳細情報-所有者-一棟・所有区分 と 送金ルール基本情報-一所有形態区分-棟/区分 の2フィールドへ設定する
        ' ''' '20161011 一棟/区分一括調整処理改善
        ' ''' 引数「ByVal list_ittobkno As List(Of String)」を削除
        ' ''' 引数「ByVal list_kbnbkno As List(Of String)」を「ByVal solist_kbnbkno As SortedList(Of Integer, String)」へ変更
        ' ''' →最初に全て一棟に設定した後で区分のみ更新するように修正する
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <param name="solist_kbnbkno"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_ExistMidFile_Syoyukbn(ByVal solist_kbnbkno As SortedList(Of Integer, String), ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim da As New OleDbDataAdapter()
        '    Dim ds As DataSet = New DataSet()
        '    Dim dt As New DataTable()
        '    Dim dr As OleDbDataReader
        '    Dim qry As String = Nothing

        '    Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        '    Dim tmpextend As String = "Excel 8.0;HDR=YES;"

        '    Dim filename As String() = New String() {"物件情報", "送金ルール情報"}
        '    Dim sheetname As String() = New String() {"物件詳細情報", "送金ルール基本情報"}
        '    Dim bknofldname As String() = New String() {"物件NO", "物件No"}
        '    Dim copyfldname As String() = New String() {"所有者-一棟・所有区分", "一所有形態区分-棟/区分"}

        '    '----- 全体用・中間用プログレスバー初期化 -----
        '    Dim obj_pgb As New ProgressBarManager
        '    Dim pgbcnt As Integer = 0
        '    Dim pgbtotalcnt_total As Integer = 2
        '    Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
        '    Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)


        '    '20161011 一棟/区分一括調整処理改善 -del sta
        '    '更新対象を取得する
        '    'Dim sql_ittobkno As String = ""
        '    'Dim sql_kbnbkno As String = ""
        '    'If list_ittobkno.Count <> 0 Then                         '20161009 不具合調査：クエリが複雑すぎる。分割して処理するなどを考える
        '    '    For Each ittobkno In list_ittobkno
        '    '        sql_ittobkno = sql_ittobkno & "," & "'" & ittobkno & "'"
        '    '    Next
        '    '    sql_ittobkno = sql_ittobkno.Remove(0, 1)
        '    'End If
        '    'If list_kbnbkno.Count <> 0 Then                         '20161009 不具合調査：クエリが複雑すぎる。分割して処理するなどを考える
        '    '    For Each kbnbkno In list_kbnbkno
        '    '        sql_kbnbkno = sql_kbnbkno & "," & "'" & kbnbkno & "'"
        '    '    Next
        '    '    sql_kbnbkno = sql_kbnbkno.Remove(0, 1)
        '    'End If
        '    '20161011 一棟/区分一括調整処理改善 -del end

        '    '--------------------------------------------------
        '    ' 物件/部屋所有者情報から対象物件Noを取得
        '    '--------------------------------------------------
        '    For cntii = 0 To 1

        '        '中断処理
        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then
        '            Return rtn
        '        End If

        '        'ファイルパス設定
        '        Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename(cntii) & ".xlsx")

        '        '--------------------------------------------------
        '        ' メイン処理
        '        '--------------------------------------------------
        '        Dim con_write As New OleDbConnection()
        '        Dim cmd_write As New OleDbCommand()

        '        Try
        '            '--------------------------------------------------
        '            ' EXCEL接続 (ファイル書込用)
        '            '--------------------------------------------------
        '            '接続文字列生成
        '            con_write.ConnectionString = _
        '                "Provider=" & tmpprovider & _
        '                "Data Source=" & existmidpath_read & ";" & _
        '                "Extended Properties=" & """" & tmpextend & """"
        '            '接続設定
        '            cmd_write.Connection = con_write

        '            '--------------------------------------------------
        '            ' 読込
        '            '--------------------------------------------------
        '            '接続オープン処理
        '            con_write.Open()

        '            '更新
        '            '一度全てのデータを「1」に設定しておく
        '            qry = ""
        '            '20161011 一棟/区分一括調整処理改善 -chg sta
        '            'qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 1 "
        '            qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = '1' "
        '            '20161011 一棟/区分一括調整処理改善 -chg end
        '            cmd_write.CommandText = qry
        '            cmd_write.ExecuteNonQuery()
        '            '20161011 一棟/区分一括調整処理改善 -chg sta
        '            ''一棟所有物件の更新
        '            'If sql_ittobkno <> "" Then
        '            '    qry = ""
        '            '    qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 1 "
        '            '    qry += " WHERE " & bknofldname(cntii) & " IN ( " & sql_ittobkno & ")"
        '            '    cmd_write.CommandText = qry
        '            '    cmd_write.ExecuteNonQuery()
        '            'End If

        '            ''区分所有物件の更新
        '            'If sql_kbnbkno <> "" Then
        '            '    qry = ""
        '            '    qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 2 "
        '            '    qry += " WHERE " & bknofldname(cntii) & " IN ( " & sql_kbnbkno & ")"
        '            '    cmd_write.CommandText = qry
        '            '    cmd_write.ExecuteNonQuery()
        '            'End If

        '            '区分所有物件の更新
        '            '更新対象が存在しない場合は処理を抜ける
        '            If solist_kbnbkno.Count = 0 Then
        '                Exit For
        '            End If

        '            '----- 個別用プログレスバー更新 -----
        '            Dim pgbcnt_part As Integer = 0
        '            Dim pgbtotalcnt_part As Integer = solist_kbnbkno.Count
        '            Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        '            For Each kbnbkno In solist_kbnbkno
        '                Dim tmp_sqltaisyo As String = kbnbkno.Value
        '                qry = ""
        '                qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 2 "
        '                qry += " WHERE " & bknofldname(cntii) & " IN ( " & tmp_sqltaisyo & ")"
        '                cmd_write.CommandText = qry
        '                cmd_write.ExecuteNonQuery()

        '                '----- 個別用プログレスバー更新 -----
        '                pgbcnt_part = pgbcnt_part + 1
        '                Call obj_pgb.pgbsettingPart(pgbcnt_part)
        '                Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
        '                '20161017 進捗表示処理による速度低下の修正 -del
        '                'Me.Refresh()
        '            Next
        '            '20161011 一棟/区分一括調整処理改善 -chg end

        '        Catch ex As Exception

        '            'エラー処理
        '            errstr = MSG_ERR_BASEMIDFILEREAD
        '            rtn = False
        '        End Try

        '        'クローズ処理
        '        con_write.Close()

        '        '----- 全体用・中間用プログレスバー更新 -----
        '        pgbcnt = pgbcnt + 1
        '        Call obj_pgb.pgbsettingTotal(pgbcnt)
        '        Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        '        Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        '        Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        '        Me.Refresh()
        '    Next

        '    '返却
        '    Return rtn

        'End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        ''' <summary>
        ''' 物件部屋所有情報の所有開始終了年月を送金ルールの適用開始終了年月にコピーする '20161004 適用開始年月の自動設定 -add
        ''' '20161011 汎用→既存書込時の不要処理除去修正 引数に「list_filename」を追加
        ''' </summary>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_ExistMidFile_SoruleTekiyoYmdCopyMain(ByRef errstr As String, ByVal list_filename As List(Of String)) As Boolean

            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            'Dim rtn As Boolean = True
            'Dim hash_ittosyoymd As New Hashtable                '一棟所有の所有開始終了年月格納
            'Dim hash_kbnsyoymd As New Hashtable                 '区分所有の所有開始終了年月格納

            ''20161011 汎用→既存書込時の不要処理除去修正 -add sta
            ''物件情報、部屋情報どちらも移行対象ではない場合は処理を抜ける
            'If list_filename.Contains("物件情報") = False And list_filename.Contains("部屋情報") = False Then
            '    Return rtn
            'End If
            ''20161011 汎用→既存書込時の不要処理除去修正 -add end

            ''物件Noと所有開始終了年月を取得
            'rtn = Set_ExistMidFile_SyoyuYmdbkno(hash_ittosyoymd, hash_kbnsyoymd, errstr)

            'If rtn Then
            '    '取得した情報から自動設定
            '    rtn = Set_ExistMidFile_SyoyuYmd(hash_ittosyoymd, hash_kbnsyoymd, errstr)
            'Else
            '    Return rtn
            'End If

            'Return rtn

            Dim rtn As Boolean = True
            Dim tmp_sql As String = ""
            Dim tmp_cnt As Integer = 0

            '実行有無チェック
            If list_filename.Contains("物件情報") = False And list_filename.Contains("部屋情報") = False Then
                Return rtn
            End If

            Try

                '一棟所有を更新
                tmp_sql = tmp_sql & " UPDATE CVTBL_送金ルール基本情報 SET "
                tmp_sql = tmp_sql & " 	 [送金ルール適用開始日] = RTRIM(LTRIM(BKSYO.[所有期間開始])) "
                tmp_sql = tmp_sql & " 	,[送金ルール適用終了日] = RTRIM(LTRIM(BKSYO.[所有期間終了])) "
                tmp_sql = tmp_sql & " FROM CVTBL_送金ルール基本情報 AS SO "
                tmp_sql = tmp_sql & " LEFT JOIN CVTBL_物件所有者情報 AS BKSYO "
                tmp_sql = tmp_sql & " ON RTRIM(LTRIM(SO.[物件No])) = RTRIM(LTRIM(BKSYO.[物件No])) "
                tmp_sql = tmp_sql & " WHERE BKSYO.[物件No] IS NOT NULL "
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql, tmp_cnt)
                tmp_sql = ""

                '区分所有を更新
                tmp_sql = tmp_sql & " UPDATE CVTBL_送金ルール基本情報 SET "
                tmp_sql = tmp_sql & " 	 [送金ルール適用開始日] = RTRIM(LTRIM(HYSYO.[所有期間開始])) "
                tmp_sql = tmp_sql & " 	,[送金ルール適用終了日] = RTRIM(LTRIM(HYSYO.[所有期間終了])) "
                tmp_sql = tmp_sql & " FROM CVTBL_送金ルール基本情報 AS SO "
                tmp_sql = tmp_sql & " LEFT JOIN CVTBL_部屋所有者情報 AS HYSYO "
                tmp_sql = tmp_sql & " ON  RTRIM(LTRIM(SO.[物件No])) = RTRIM(LTRIM(HYSYO.[物件No])) "
                tmp_sql = tmp_sql & " AND RTRIM(LTRIM(SO.[部屋No])) = RTRIM(LTRIM(HYSYO.[部屋No])) "
                tmp_sql = tmp_sql & " WHERE HYSYO.[物件No] IS NOT NULL "
                tmp_sql = tmp_sql & "  "
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql, tmp_cnt)
                tmp_sql = ""

            Catch ex As Exception

                errstr = MSG_ERR_BASEMIDFILEREAD
                rtn = False

            End Try

            Return rtn
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
        End Function

        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        ' ''' <summary>
        ' ''' 中間ファイルの物件/部屋所有者情報から物件Noと所有開始終了年月を紐付けて取得してオブジェクトへ格納 '20161004 適用開始年月の自動設定
        ' ''' </summary>
        ' ''' <param name="hash_ittosyoymd"></param>
        ' ''' <param name="hash_kbnsyoymd"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_ExistMidFile_SyoyuYmdbkno(ByRef hash_ittosyoymd As Hashtable, ByRef hash_kbnsyoymd As Hashtable, ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim da As New OleDbDataAdapter()
        '    Dim ds As DataSet = New DataSet()
        '    Dim dt As New DataTable()
        '    Dim dr As OleDbDataReader
        '    Dim qry As String = Nothing

        '    Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        '    Dim tmpextend As String = "Excel 8.0;HDR=YES;"
        '    Dim list_filename As New List(Of String) From {"物件情報", "部屋情報"}


        '    '--------------------------------------------------
        '    ' 物件/部屋所有者情報から対象物件Noを取得
        '    '--------------------------------------------------
        '    For Each filename In list_filename

        '        '中断処理
        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then
        '            Return rtn
        '        End If

        '        'ファイルパス設定
        '        Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")

        '        '--------------------------------------------------
        '        ' メイン処理
        '        '--------------------------------------------------
        '        Dim con_read As New OleDbConnection()
        '        Dim cmd_read As New OleDbCommand()
        '        Dim con_write As New OleDbConnection()
        '        Dim cmd_write As New OleDbCommand()

        '        Try
        '            '--------------------------------------------------
        '            ' ①EXCEL接続 (ファイル読込用)
        '            '--------------------------------------------------
        '            '接続文字列生成
        '            con_read.ConnectionString = _
        '                "Provider=" & tmpprovider & _
        '                "Data Source=" & existmidpath_read & ";" & _
        '                "Extended Properties=" & """" & tmpextend & """"
        '            '接続設定
        '            cmd_read.Connection = con_read

        '            '--------------------------------------------------
        '            ' 読込
        '            '--------------------------------------------------
        '            '接続オープン処理
        '            con_read.Open()

        '            Dim sheetname As String = ""
        '            Select Case filename
        '                Case "物件情報"
        '                    sheetname = "物件所有者情報"
        '                    qry = "SELECT [物件No],[所有期間開始],[所有期間終了] FROM [" & sheetname & "$] "
        '                Case "部屋情報"
        '                    sheetname = "部屋所有者情報"
        '                    qry = "SELECT [物件No],[部屋No],[所有期間開始],[所有期間終了] FROM [" & sheetname & "$] "
        '            End Select

        '            cmd_read = con_read.CreateCommand
        '            cmd_read.CommandText = qry
        '            dt = New DataTable
        '            dr = cmd_read.ExecuteReader
        '            dt.Load(dr)

        '            '1行ずつ取得
        '            Dim row As DataRow
        '            For Each row In dt.Rows
        '                '作業用変数
        '                Dim tmp_bkno As String = ""
        '                Dim tmp_hyno As String = ""
        '                Dim tmp_syosta As String = ""
        '                Dim tmp_syoend As String = ""

        '                For cntjj = 0 To dt.Columns.Count - 1
        '                    Dim fldname As String = dt.Columns(cntjj).ColumnName
        '                    Dim fldvalue As String = row(cntjj).ToString.Trim
        '                    Select Case fldname
        '                        Case "物件No"
        '                            tmp_bkno = fldvalue
        '                        Case "部屋No"
        '                            tmp_hyno = fldvalue
        '                        Case "所有期間開始"
        '                            tmp_syosta = fldvalue
        '                        Case "所有期間終了"
        '                            tmp_syoend = fldvalue
        '                    End Select
        '                Next

        '                'オブジェクトへ格納
        '                Dim syostaendymd As String = tmp_syosta & STR_SPLIT_1 & tmp_syoend  'オブジェクト格納用

        '                Select Case filename
        '                    Case "物件情報"
        '                        '一棟所有の所有期間開始年月
        '                        If tmp_bkno <> "" And hash_ittosyoymd.Contains(tmp_bkno) = False Then
        '                            hash_ittosyoymd.Add(tmp_bkno, syostaendymd)
        '                        End If
        '                    Case "部屋情報"
        '                        '区分所有の所有期間開始年月
        '                        Dim tmp_key As String = tmp_bkno & "-" & tmp_hyno
        '                        If tmp_key <> "" And hash_kbnsyoymd.Contains(tmp_key) = False Then
        '                            hash_kbnsyoymd.Add(tmp_key, syostaendymd)
        '                        End If
        '                End Select

        '            Next

        '        Catch ex As Exception

        '            'エラー処理
        '            errstr = MSG_ERR_BASEMIDFILEREAD
        '            rtn = False
        '        End Try

        '        'クローズ処理
        '        con_read.Close()
        '        con_write.Close()
        '    Next

        '    '返却
        '    Return rtn

        'End Function

        ' ''' <summary>
        ' ''' 所有者情報から取得した物件No、適用開始終了年月を自動設定 '20161004 適用開始年月の自動設定
        ' ''' </summary>
        ' ''' <param name="hash_ittosyoymd"></param>
        ' ''' <param name="hash_kbnsyoymd"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_ExistMidFile_SyoyuYmd(ByRef hash_ittosyoymd As Hashtable, ByRef hash_kbnsyoymd As Hashtable, ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim da As New OleDbDataAdapter()
        '    Dim ds As DataSet = New DataSet()
        '    Dim dt As New DataTable()
        '    Dim qry As String = Nothing

        '    Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        '    Dim tmpextend As String = "Excel 8.0;HDR=YES;"

        '    Dim filename As String = "送金ルール情報"
        '    Dim sheetname As String = "送金ルール基本情報"
        '    Dim fldname_staymd As String = "送金ルール適用開始日"
        '    Dim fldname_endymd As String = "送金ルール適用終了日"

        '    '----- 全体用・中間用プログレスバー更新 -----
        '    Dim obj_pgb As New ProgressBarManager
        '    Dim pgbcnt As Integer = 0
        '    Dim pgbtotalcnt_total As Integer = 2
        '    Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
        '    Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)

        '    'ファイルパス設定
        '    Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")


        '    '--------------------------------------------------
        '    ' EXCEL接続 (ファイル書込用)
        '    '--------------------------------------------------
        '    Dim con_write As New OleDbConnection()
        '    Dim cmd_write As New OleDbCommand()
        '    '接続文字列生成
        '    con_write.ConnectionString = _
        '        "Provider=" & tmpprovider & _
        '        "Data Source=" & existmidpath_read & ";" & _
        '        "Extended Properties=" & """" & tmpextend & """"
        '    '接続設定
        '    cmd_write.Connection = con_write
        '    '接続オープン処理
        '    con_write.Open()

        '    '--------------------------------------------------
        '    ' 一棟所有分を更新
        '    '--------------------------------------------------
        '    '----- 個別用プログレスバー初期化 -----
        '    Dim pgbcnt_part As Integer = 0
        '    Dim pgbtotalcnt_part As Integer = hash_ittosyoymd.Count
        '    Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        '    For Each ittoitem In hash_ittosyoymd

        '        '中断処理
        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then
        '            Return rtn
        '        End If

        '        Dim tmp_bkno As String = ittoitem.Key
        '        Dim tmp_syoymd() As String = Split(ittoitem.Value.ToString, STR_SPLIT_1)
        '        Dim tmp_staymd As String = tmp_syoymd(0)
        '        Dim tmp_endymd As String = tmp_syoymd(1)

        '        Try
        '            '一棟所有物件の更新
        '            If tmp_bkno <> "" Then
        '                qry = ""
        '                qry += " UPDATE [" & sheetname & "$] SET [" & fldname_staymd & "] = '" & tmp_staymd & "'" & "," & "[" & fldname_endymd & "] = '" & tmp_endymd & "'"
        '                qry += " WHERE [物件No] = " & "'" & tmp_bkno & "'"
        '                cmd_write.CommandText = qry
        '                cmd_write.ExecuteNonQuery()
        '            End If

        '        Catch ex As Exception

        '            'エラー処理
        '            errstr = MSG_ERR_BASEMIDFILEREAD
        '            rtn = False
        '            'クローズ処理
        '            con_write.Close()
        '            '返却
        '            Return rtn
        '        End Try

        '        '----- 個別用プログレスバー更新 -----
        '        pgbcnt_part = pgbcnt_part + 1
        '        Call obj_pgb.pgbsettingPart(pgbcnt_part)
        '        Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
        '        '20161017 進捗表示処理による速度低下の修正 -del
        '        'Me.Refresh()
        '    Next

        '    '----- 全体用・中間用プログレスバー更新 -----
        '    pgbcnt = pgbcnt + 1
        '    Call obj_pgb.pgbsettingTotal(pgbcnt)
        '    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        '    Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        '    Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        '    Me.Refresh()

        '    '--------------------------------------------------
        '    ' 区分所有分を更新
        '    '--------------------------------------------------
        '    '----- 個別用プログレスバー初期化 -----
        '    pgbcnt_part = 0
        '    '20161018 プログレスバーカウントエラー修正 -chg sta
        '    'pgbtotalcnt_part = hash_ittosyoymd.Count
        '    pgbtotalcnt_part = hash_kbnsyoymd.Count
        '    '20161018 プログレスバーカウントエラー修正 -chg end
        '    Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        '    For Each kbnitem In hash_kbnsyoymd

        '        '中断処理
        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then
        '            Return rtn
        '        End If

        '        Dim tmp_bkhyno As String = kbnitem.Key
        '        Dim tmp_syoymd() As String = Split(kbnitem.Value.ToString, STR_SPLIT_1)
        '        Dim tmp_staymd As String = tmp_syoymd(0)
        '        Dim tmp_endymd As String = tmp_syoymd(1)

        '        Try
        '            '一棟所有物件の更新
        '            If tmp_bkhyno <> "-" Then
        '                qry = ""
        '                qry += " UPDATE [" & sheetname & "$] SET [" & fldname_staymd & "] = '" & tmp_staymd & "'" & "," & "[" & fldname_endymd & "] = '" & tmp_endymd & "'"
        '                qry += " WHERE [物件No] + '-' + [部屋No] = " & "'" & tmp_bkhyno & "'"
        '                cmd_write.CommandText = qry
        '                cmd_write.ExecuteNonQuery()
        '            End If

        '        Catch ex As Exception

        '            'エラー処理
        '            errstr = MSG_ERR_BASEMIDFILEREAD
        '            rtn = False
        '            'クローズ処理
        '            con_write.Close()
        '            '返却
        '            Return rtn
        '        End Try

        '        '----- 個別用プログレスバー更新 -----
        '        pgbcnt_part = pgbcnt_part + 1
        '        Call obj_pgb.pgbsettingPart(pgbcnt_part)
        '        Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
        '        '20161017 進捗表示処理による速度低下の修正 -del
        '        'Me.Refresh()
        '    Next

        '    'クローズ処理
        '    con_write.Close()

        '    '----- 全体用・中間用プログレスバー更新 -----
        '    pgbcnt = pgbcnt + 1
        '    Call obj_pgb.pgbsettingTotal(pgbcnt)
        '    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        '    Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        '    Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        '    Me.Refresh()

        '    '返却
        '    Return rtn

        'End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        ''' <summary>
        ''' 中間ファイルから件数を取得して表示する 20160905 汎用コンバートの件数表示修正
        ''' </summary>
        ''' <remarks></remarks>
        Private Function Set_BaseCVItemCnt() As Boolean

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim hash_lbltosheet As New Hashtable
            Dim hash_lbltocnt As New Hashtable
            Dim basemidfilepath As String = Me.txtMidDirPath.Text
            Dim middirpath As String = ""
            '20161012 中間ファイル件数表示速度改善 -del sta
            'Dim rtn As Boolean = True
            'Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            'Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            'Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            'Dim maxrowcnt As Integer                                        '既存データの行数
            'Dim startrow As Integer = 0
            'Dim matchflg As Boolean = False
            'Dim flg As Boolean = True
            'Dim list_cvitem As New List(Of String) From _
            '    {"バス交通マスタ", "学校区マスタ", "エリアマスタ", "保険種類マスタ", "特約マスタ", "変動費設定内容", _
            '     "自社情報", "家主情報", "仲介・管理業者情報", "修繕業者情報", "ライフライン業者情報", "保険業者情報", "家賃保証業者情報", "施設保守業者情報", "施工業者情報", _
            '     "物件情報", "部屋情報", "部屋設備情報", "契約者情報", "契約情報", "請求情報_未収", "請求情報_預り"}
            '20161012 中間ファイル件数表示速度改善 -del sta

            '20161012 中間ファイル件数表示速度改善 -add sta
            Dim normalflg As Boolean = True                                 '20161012 中間ファイル件数表示速度改善 -add
            Dim list_cvitem As New List(Of String) From _
                {"バス交通マスタ", "学校区マスタ", "エリアマスタ", "保険種類マスタ", "特約マスタ", _
                 "自社情報", "家主情報", "仲介業者情報", "修繕業者情報", "ライフライン業者情報", "保険業者情報", "家賃保証業者情報", "施設保守業者情報", "施工業者情報", _
                 "物件情報", "部屋情報", "部屋設備情報", "契約者情報", "契約情報", "家主固定控除情報"}
            '20161012 中間ファイル件数表示速度改善 -add end
            
            '20161012 中間ファイル件数表示速度改善 -chg sta
            'flg = EtcMethod.Chk_FileExist(basemidfilepath)

            'If flg Then
            '    middirpath = Path.GetDirectoryName(basemidfilepath)
            'Else
            '    MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '    Exit Sub
            'End If
            normalflg = EtcMethod.Chk_FileExist(basemidfilepath)
            If normalflg Then
                middirpath = Path.GetDirectoryName(basemidfilepath)
                CV_FROM_MIDDLE = Path.GetFileNameWithoutExtension(basemidfilepath)
            Else
                MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '20161108_2 レビュー結果戻り修正 -chg sta
                'Return True                                                     '20161108 レビュー結果：ここTRUEで問題ない？＋必要な箇所のみ"Return False"で抜ける処理を入れておき、正常Returnは最後のみ記載する(問題なければ)
                Return False
                '20161108_2 レビュー結果戻り修正 -chg end
            End If
            '20161012 中間ファイル件数表示速度改善 -chg end

            '----- 中間用プログレスバー初期化 -----
            Dim obj_pgb As New ProgressBarManager
            Dim pgbcnt As Integer = 0
            Dim pgbtotalcnt As Integer = list_cvitem.Count
            Call obj_pgb.pgbInitChkTotal(pgbtotalcnt)
            Me.lblCheckSituation.Text = SITUATION_EXTRACTION
            Me.lblPgbCheck.Text = "0 %"
            Me.pgbCheck.Value = 0
            Me.pnlPrgChk.Visible = True
            Me.tabCtrlMain.Enabled = False
            Me.Refresh()    '20161017 画面リフレッシュ機能の追加 -add

            '20161012 中間ファイル件数表示速度改善 -add sta
            'オープン処理
            Dim con_read As New OleDbConnection()
            normalflg = excelfile.ExcelFile_ReadOpenOnly(middirpath, CV_FROM_MIDDLE, con_read)

            'オープン処理失敗時は処理を抜ける
            If normalflg = False Then
                Call excelfile.ExcelFile_ReadClose(con_read)
                '20161104 中間ファイル読込エラー時の処理対応 -chg sta
                'MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                'Exit Sub
                Me.Close()
                Return False                    '20161108 レビュー結果：SUBからFunctionへ変更した？
                '                                20161108_2 レビュー結果戻り修正
                '                                  すみません、SUBからFunctionへ変更しました。ヘッダー部へ記載しておりませんでした。
                '20161104 中間ファイル読込エラー時の処理対応 -chg end
            End If
            '20161012 中間ファイル件数表示速度改善 -add end

            For Each basecvitem In list_cvitem

                '20161012 中間ファイル件数表示速度改善 -chg sta
                'Dim tmp_lbl As New Label
                'Dim tmp_sheetname As String = ""

                'Select Case basecvitem
                '    Case "バス交通マスタ" : hash_lbltosheet.Add(Me.lblKiMstBusCnt, "バス交通マスタ")
                '    Case "学校区マスタ" : hash_lbltosheet.Add(Me.lblKiMstSchoolCnt, "学校区マスタ")
                '    Case "エリアマスタ" : hash_lbltosheet.Add(Me.lblKiMstAreaCnt, "エリアマスタ")
                '    Case "保険種類マスタ" : hash_lbltosheet.Add(Me.lblKiMstHokenruiCnt, "保険種類マスタ")
                '    Case "特約マスタ" : hash_lbltosheet.Add(Me.lblKiMstTokuyakuCnt, "特約マスタ")
                '        'Case "変動費設定内容" : hash_lbltosheet.Add(Me.lblKiMstHendoCnt, "変動費設定内容")    '20160926 選定した移行項目をプログラムへ反映する修正 -del
                '    Case "自社情報" : hash_lbltosheet.Add(Me.lblKiJisyaBaseCnt, "自社情報")
                '    Case "家主情報" : hash_lbltosheet.Add(Me.lblKiOwBaseCnt, "家主情報")
                '    Case "仲介・管理業者情報" : hash_lbltosheet.Add(Me.lblKiGyCyukaiBaseCnt, "仲介業者情報")
                '    Case "修繕業者情報" : hash_lbltosheet.Add(Me.lblKiGySyuzenBaseCnt, "修繕業者情報")
                '    Case "ライフライン業者情報" : hash_lbltosheet.Add(Me.lblKiGyLifelineBaseCnt, "ライフライン業者情報")
                '    Case "保険業者情報" : hash_lbltosheet.Add(Me.lblKiGyHokenBaseCnt, "保険業者情報")
                '    Case "家賃保証業者情報" : hash_lbltosheet.Add(Me.lblKiGyYatinhosyoBaseCnt, "家賃保証業者情報")
                '    Case "施設保守業者情報" : hash_lbltosheet.Add(Me.lblKiGySisetuBaseCnt, "施設保守業者情報")
                '    Case "施工業者情報" : hash_lbltosheet.Add(Me.lblKiGySekoBaseCnt, "施工業者情報")
                '    Case "物件情報" : hash_lbltosheet.Add(Me.lblKiBkBaseCnt, "物件情報")
                '    Case "部屋情報" : hash_lbltosheet.Add(Me.lblKiHyBaseCnt, "部屋情報")
                '    Case "部屋設備情報" : hash_lbltosheet.Add(Me.lblKiHySetubiCnt, "部屋設備情報")
                '    Case "契約者情報" : hash_lbltosheet.Add(Me.lblKiKysBaseCnt, "契約者情報")
                '    Case "契約情報" : hash_lbltosheet.Add(Me.lblKiKyBaseCnt, "契約情報")
                '        'Case "請求情報_未収" : hash_lbltosheet.Add(Me.lblKiSqMiBaseCnt, "運用開始滞納金情報")  '20160926 選定した移行項目をプログラムへ反映する修正 -del
                '        'Case "請求情報_預り" : hash_lbltosheet.Add(Me.lblKiSqAzBaseCnt, "預り金情報")          '20160926 選定した移行項目をプログラムへ反映する修正 -del
                'End Select

                Dim tmp_lbl As New Label
                Select Case basecvitem
                    Case "バス交通マスタ" : tmp_lbl = Me.lblKiMstBusCnt
                    Case "学校区マスタ" : tmp_lbl = Me.lblKiMstSchoolCnt
                    Case "エリアマスタ" : tmp_lbl = Me.lblKiMstAreaCnt
                    Case "保険種類マスタ" : tmp_lbl = Me.lblKiMstHokenruiCnt
                    Case "特約マスタ" : tmp_lbl = Me.lblKiMstTokuyakuCnt
                    Case "自社情報" : tmp_lbl = Me.lblKiJisyaBaseCnt
                    Case "家主情報" : tmp_lbl = Me.lblKiOwBaseCnt
                    Case "仲介業者情報" : tmp_lbl = Me.lblKiGyCyukaiBaseCnt
                    Case "修繕業者情報" : tmp_lbl = Me.lblKiGySyuzenBaseCnt
                    Case "ライフライン業者情報" : tmp_lbl = Me.lblKiGyLifelineBaseCnt
                    Case "保険業者情報" : tmp_lbl = Me.lblKiGyHokenBaseCnt
                    Case "家賃保証業者情報" : tmp_lbl = Me.lblKiGyYatinhosyoBaseCnt
                    Case "施設保守業者情報" : tmp_lbl = Me.lblKiGySisetuBaseCnt
                    Case "施工業者情報" : tmp_lbl = Me.lblKiGySekoBaseCnt
                    Case "物件情報" : tmp_lbl = Me.lblKiBkBaseCnt
                    Case "部屋情報" : tmp_lbl = Me.lblKiHyBaseCnt
                    Case "部屋設備情報" : tmp_lbl = Me.lblKiHySetubiCnt
                    Case "契約者情報" : tmp_lbl = Me.lblKiKysBaseCnt
                    Case "契約情報" : tmp_lbl = Me.lblKiKyBaseCnt
                    Case "家主固定控除情報" : tmp_lbl = Me.lblKiSqOwKojoBaseCnt
                End Select

                'データを取得しオブジェクトへ格納
                Dim tmp_sql As String = "SELECT * FROM [" & basecvitem & "$] "
                Dim readtbl As New DataTable
                Dim cmd_read As New OleDbCommand()

                Dim dr As OleDbDataReader
                cmd_read = con_read.CreateCommand
                cmd_read.CommandText = tmp_sql
                dr = cmd_read.ExecuteReader
                readtbl.Load(dr)

                '行数取得
                Dim rowcnt As Integer = readtbl.Rows.Count - 10

                '件数表示
                If rowcnt <> 0 Then
                    tmp_lbl.Text = Int32.Parse(rowcnt).ToString("#,0") & " 件"
                Else
                    tmp_lbl.Text = "なし"
                End If

                '20161012 中間ファイル件数表示速度改善 -chg end

                '----- 中間用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsettingChkTotal(pgbcnt)
                Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt, True)
                Me.Refresh()
            Next

            'クローズ処理 '20161012 中間ファイル件数表示速度改善 -add
            excelfile.ExcelFile_ReadClose(con_read)

            '20161012 中間ファイル件数表示速度改善 -del sta
            ''Excelファイル初期設定                          
            'rtn = excelfile.Set_ExcelFile_ReadOpen_BaseMidCnt(appli, wbook, wsheet, startrow, maxrowcnt, middirpath, CV_FROM_MIDDLE, hash_lbltosheet, hash_lbltocnt)

            ''Excelファイル設定時にエラーが生じた際は処理を抜ける
            'If rtn = False Then
            '    Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
            '    MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '    Exit Sub
            'End If

            ''件数表示
            'For Each baseitem In hash_lbltocnt
            '    Dim tmp_lbl As Label = baseitem.Key
            '    Dim cnt As Integer = baseitem.Value
            '    If cnt <> 0 Then
            '        tmp_lbl.Text = Int32.Parse(cnt).ToString("#,0") & " 件"
            '    Else
            '        tmp_lbl.Text = "なし"
            '    End If

            'Next

            ''Excelファイル終了設定
            'Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            '20161012 中間ファイル件数表示速度改善 -del end

            '----- 中間用プログレスバー表示設定 -----
            Me.pnlPrgChk.Visible = False
            Me.tabCtrlMain.Enabled = True

            '20161104 中間ファイル読込エラー時の処理対応 -add
            Return True

        End Function

        ''' <summary>
        ''' 既存用→汎用用中間ファイルへコピーする処理(開発用) 20160912 既存用→汎用用中間ファイルへのコピー処理
        ''' (開発用処理)
        ''' </summary>
        ''' <param name="errstr"></param>
        ''' <param name="list_cv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_ExistMidFile_Copy(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

            Dim rtn As Boolean = True
            Dim normalflg As Boolean = True
            Dim list_existmidfileinfo As New List(Of String)


            For Each cvitem In list_cv

                'コンバート対象項目からファイル名とシート名を取得
                cvitem = cvitem.Replace("-", STR_SPLIT_1)
                Dim tmp_existmidstr() As String = Split(cvitem, STR_SPLIT_1)
                Dim existfilename As String = tmp_existmidstr(0)
                Dim existsheetname As String = tmp_existmidstr(1)

                '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
                Call Me.Set_MidDataInfoToObj(existfilename, existsheetname, False)

                '移行項目からオープンする汎用用中間ファイル名を取得
                Dim basemidfsname As String = ""

                If Hash_ExistMidToBaseMid_FS.Contains(cvitem) Then
                    basemidfsname = Hash_ExistMidToBaseMid_FS(cvitem)

                    '----- デバッグ用処理 ----- sta
                    'Dim list As New List(Of String) From {"物件情報@#@物件詳細情報"}
                    'If list.Contains(cvitem) = False Then
                    '    basemidfsname = ""
                    'End If
                    '----- デバッグ用処理 ----- end

                    If basemidfsname <> "" Then
                        Dim tmp_basemidstr() As String = Split(basemidfsname, STR_SPLIT_1)
                        Dim basefilename As String = tmp_basemidstr(0)
                        Dim basesheetname As String = tmp_basemidstr(1)

                        '既存用中間ファイルからデータ取得
                        Dim middata As New Object
                        Dim hash_middata As New Hashtable
                        Dim writerowcnt As Integer = 0
                        normalflg = Me.Set_ExistMidDataToObj(existfilename, existsheetname, middata, hash_middata, writerowcnt, basesheetname, errstr)

                        '汎用用中間ファイルへコピー
                        If normalflg Then

                            '開くファイル名とシート名を元に初期化フラグを設定(一度開いたシート名をオブジェクトへ退避させておく)
                            Dim iniflg As Boolean = False
                            If list_existmidfileinfo.Contains(basefilename & STR_SPLIT_1 & basesheetname) = False Then
                                list_existmidfileinfo.Add(basefilename & STR_SPLIT_1 & basesheetname)
                                iniflg = True
                            End If

                            'コピー処理
                            Call Me.Set_ExistMidDataToBaseMidFile(basefilename, basesheetname, hash_middata, writerowcnt, iniflg)
                        End If
                    End If
                End If
            Next

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 既存用中間ファイルからデータを取得してオブジェクトへ格納 20160912 既存用→汎用用中間ファイルへのコピー処理
        ''' (開発用処理)
        ''' </summary>
        ''' <param name="existmidfilename"></param>
        ''' <param name="existmidsheetname"></param>
        ''' <param name="middata"></param>
        ''' <param name="hash_middata"></param>
        ''' <param name="writerowcnt"></param>
        ''' <param name="basesheetname"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_ExistMidDataToObj(ByVal existmidfilename As String, _
                                               ByVal existmidsheetname As String, _
                                               ByRef middata As Object, _
                                               ByRef hash_middata As Hashtable, _
                                               ByRef writerowcnt As Integer, _
                                               ByVal basesheetname As String, _
                                               ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim matchflg As Boolean = False
            Dim tmp_middirpath As String = Me.txtExistMidToBaseMid.Text

            'Excelファイル初期設定                          
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, tmp_middirpath, existmidfilename, existmidsheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If rtn = False Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            'データが存在しない場合は処理を抜ける
            If rowcnt = 0 Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            'ヘッダー取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '対象ヘッダー検索
            For cntjj = 1 To columncnt
                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                Dim keysfname As String = existmidsheetname & STR_SPLIT_1 & fldname

                '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -chg sta
                ''既存用中間ファイルヘッダー存在有無確認
                'If List_Existmidheader.Contains(keysfname) Then

                '    '紐付く汎用中間ファイルヘッダー有無確認
                '    If Hash_ExistMidToBaseMid_SH.Contains(keysfname) Then
                '        Dim tmp_existstr() As String = Split(Hash_ExistMidToBaseMid_SH(keysfname).ToString, STR_SPLIT_1)
                '        If basesheetname = tmp_existstr(0) Then
                '            middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
                '            hash_middata.Add(keysfname, middata)
                '        End If
                '    End If

                'End If

                '移行対象フラグ確認
                If List_ExistMiddatacv.Contains(keysfname) Then
                    '既存用中間ファイルヘッダー存在有無確認
                    If List_Existmidheader.Contains(keysfname) Then
                        '紐付く汎用中間ファイルヘッダー有無確認
                        If Hash_ExistMidToBaseMid_SH.Contains(keysfname) Then
                            Dim tmp_existstr() As String = Split(Hash_ExistMidToBaseMid_SH(keysfname).ToString, STR_SPLIT_1)
                            If basesheetname = tmp_existstr(0) Then
                                middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
                                hash_middata.Add(keysfname, middata)
                            End If
                        End If
                    End If
                End If
                '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -chg end
            Next

            '書込行数を返却
            writerowcnt = rowcnt
            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 取得した既存用中間ファイルデータを汎用用中間ファイルへコピー '20160912 既存用→汎用用中間ファイルへのコピー処理
        ''' (開発用処理)
        ''' </summary>
        ''' <param name="basemidfilename"></param>
        ''' <param name="basemidsheetname"></param>
        ''' <param name="hash_existmiddata"></param>
        ''' <param name="writerowcnt"></param>
        ''' <param name="iniflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_ExistMidDataToBaseMidFile(ByVal basemidfilename As String, _
                                                       ByVal basemidsheetname As String, _
                                                       ByVal hash_existmiddata As Hashtable, _
                                                       ByVal writerowcnt As Integer, _
                                                       ByVal iniflg As Boolean) As Boolean

            Dim rtn As Boolean = True
            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim headerrow As Integer = 2
            Dim tmp_middirpath As String = Me.txtExistMidToBaseMid.Text


            '**************************************************
            ' 作業準備
            '**************************************************
            'Excelファイル初期設定(汎用用中間ファイルの初期化有無で処理を分岐させる)
            If iniflg Then
                rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, tmp_middirpath, basemidfilename, basemidsheetname)
            Else
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, tmp_middirpath, basemidfilename, basemidsheetname)
            End If

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If rtn = False Then
                Return rtn
            End If

            '既存用中間ファイルが空の場合は処理を抜ける
            If hash_existmiddata.Count = 0 Then
                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
                Return rtn
            End If

            '列数を退避
            Dim writecol As Integer = columncnt

            'ヘッダー取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value

            '既存用中間ファイルから取得したデータを元に貼付け先を取得
            For Each existmiddata In hash_existmiddata

                Dim existmidfldname As String = existmiddata.Key
                Dim existmidcoldata As Object = existmiddata.Value
                Dim fldmatchflg As Boolean = False

                '照合→コピー処理
                If Hash_ExistMidToBaseMid_SH.Contains(existmidfldname) Then

                    Dim basemiddata() As String = Split(Hash_ExistMidToBaseMid_SH(existmidfldname).ToString, STR_SPLIT_1)
                    Dim basemidfldname As String = basemiddata(1)

                    '20160915 開発用_貼付け処理改善 -chg sta
                    ''取得した列名を元にコピー処理
                    ''一致するフィールドが存在する場合
                    'For cntjj = 1 To columncnt
                    '    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
                    '    If fldname = basemidfldname Then
                    '        wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(startrow + writerowcnt - 1, cntjj)).Value = existmidcoldata
                    '        fldmatchflg = True
                    '        Exit For
                    '    End If
                    'Next

                    ''一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
                    'If fldmatchflg = False Then
                    '    writecol = writecol + 1
                    '    wsheet.Cells(headerrow, writecol).Value = existmidfldname
                    '    wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = existmidcoldata
                    'End If
                    If basemidfldname <> "共有No" Then
                        '取得した列名を元にコピー処理
                        '一致するフィールドが存在する場合
                        For cntjj = 1 To columncnt
                            Dim fldname As String = headervalue(1, cntjj).ToString.Trim
                            If fldname = basemidfldname Then
                                wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(startrow + writerowcnt - 1, cntjj)).Value = existmidcoldata
                                fldmatchflg = True
                                Exit For
                            End If
                        Next

                        '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -del sta
                        '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
                        'If fldmatchflg = False Then
                        '    writecol = writecol + 1
                        '    wsheet.Cells(headerrow, writecol).Value = existmidfldname
                        '    wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = existmidcoldata
                        'End If
                        '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -del end
                    End If
                    '20160915 開発用_貼付け処理改善 -chg end
                End If
            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 移行項目のファイル名を集約する
        ''' </summary>
        ''' <param name="list_cv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_FileNameIntensive(ByVal list_cv As List(Of String)) As List(Of String)

            Dim rtn_list As New List(Of String)

            'データが存在しない場合は処理を抜ける(念の為)
            If list_cv.Count = 0 Then
                Return rtn_list
            End If

            'ファイル名を集約する
            For Each cvitem In list_cv
                Dim tmp_str() As String = cvitem.ToString.Split("-")
                If rtn_list.Contains(tmp_str(0)) = False Then
                    rtn_list.Add(tmp_str(0))
                End If
            Next

            Return rtn_list

        End Function

        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
        ''' <summary>
        ''' 汎用で移行対象となっている全項目名をオブジェクトへ格納 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_BaseCvItem() As List(Of String)

            Dim rtn_list As New List(Of String)

            rtn_list.Add("各マスタ情報-バス交通マスタ")
            rtn_list.Add("各マスタ情報-バス停マスタ")
            rtn_list.Add("各マスタ情報-学校区マスタ")
            rtn_list.Add("各マスタ情報-エリアマスタ")
            rtn_list.Add("各マスタ情報-保険種類マスタ")
            rtn_list.Add("各マスタ情報-特約マスタ")
            '20161028 物件/部屋鍵取得方法修正 -add
            rtn_list.Add("各マスタ情報-鍵タイトルマスタ")
            rtn_list.Add("業者情報-仲介業者基本情報")
            rtn_list.Add("業者情報-仲介業者口座情報")
            rtn_list.Add("業者情報-仲介業者メモ情報")
            rtn_list.Add("業者情報-保険業者基本情報")
            rtn_list.Add("業者情報-保険業者口座情報")
            rtn_list.Add("業者情報-保険業者メモ情報")
            rtn_list.Add("業者情報-家賃保証業者基本情報")
            rtn_list.Add("業者情報-家賃保証業者メモ情報")
            rtn_list.Add("業者情報-修繕業者基本情報")
            rtn_list.Add("業者情報-修繕業者口座情報")
            rtn_list.Add("業者情報-修繕業者メモ情報")
            rtn_list.Add("業者情報-ライフライン業者情報")
            rtn_list.Add("業者情報-施工業者情報")
            rtn_list.Add("業者情報-施設保守業者情報")
            rtn_list.Add("自社情報-自社基本情報")
            rtn_list.Add("自社情報-自社口座情報")
            rtn_list.Add("自社情報-自社担当者情報")
            rtn_list.Add("自社情報-自社メモ情報")
            rtn_list.Add("自社情報-振込依頼人情報")
            rtn_list.Add("自社情報-口座振替情報")
            rtn_list.Add("自社情報-家賃入金口座情報")
            rtn_list.Add("家主情報-家主基本情報")
            rtn_list.Add("家主情報-家主口座情報")
            rtn_list.Add("家主情報-家主メモ情報")
            rtn_list.Add("契約者情報-契約者基本情報")
            rtn_list.Add("契約者情報-契約者口座情報")
            rtn_list.Add("契約者情報-契約者メモ情報")
            rtn_list.Add("契約者情報-契約者照合用カナ情報")
            rtn_list.Add("契約者情報-契約者保証人情報")
            rtn_list.Add("物件情報-物件基本情報")
            rtn_list.Add("物件情報-物件詳細情報")
            rtn_list.Add("物件情報-物件所有者情報")
            rtn_list.Add("物件情報-物件交通情報")
            rtn_list.Add("物件情報-物件メモ情報")
            rtn_list.Add("物件情報-物件鍵情報")
            rtn_list.Add("物件情報-物件近隣駐車場情報")
            rtn_list.Add("部屋情報-部屋基本情報")
            rtn_list.Add("部屋情報-部屋詳細情報")
            rtn_list.Add("部屋情報-部屋所有者情報")
            rtn_list.Add("部屋情報-部屋駐車場情報")
            rtn_list.Add("部屋情報-部屋特約情報")
            rtn_list.Add("部屋情報-部屋鍵情報")
            rtn_list.Add("部屋情報-部屋間取内訳情報")
            rtn_list.Add("部屋情報-部屋設備情報")
            rtn_list.Add("部屋情報-部屋入金項目情報")
            rtn_list.Add("部屋情報-部屋メモ情報")
            rtn_list.Add("部屋情報-部屋共通セールスポイント情報")
            rtn_list.Add("送金ルール情報-送金ルール基本情報")
            rtn_list.Add("送金ルール情報-送金ルール送金先情報")
            rtn_list.Add("送金ルール情報-送金ルール入金項目情報")
            rtn_list.Add("送金ルール情報-送金ルール控除項目情報")
            rtn_list.Add("契約情報-契約基本情報")
            rtn_list.Add("契約情報-契約履歴情報")
            rtn_list.Add("契約情報-契約車情報")
            rtn_list.Add("契約情報-契約契約者情報")
            rtn_list.Add("契約情報-契約保証人情報")
            rtn_list.Add("契約情報-契約入居者情報")
            rtn_list.Add("契約情報-契約特約事項情報")
            rtn_list.Add("契約情報-契約保険情報")
            rtn_list.Add("契約情報-契約メモ情報")
            rtn_list.Add("契約情報-契約入金項目情報")
            rtn_list.Add("契約情報-契約次回入金項目情報")
            rtn_list.Add("請求情報-家主固定控除情報")

            Return rtn_list

        End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

        '20160929 汎用→既存コピー処理改善対応 -del sta
        ' ''' <summary>
        ' ''' 一棟区分を所有情報から判断して自動設定する '20160920 プログラム自動設定箇所の追加 -add
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidFile_SyoyuCopyMain(ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim list_ittobkno As New List(Of String)
        '    Dim list_kbnbkno As New List(Of String)

        '    '汎用用中間ファイルからデータを取得する
        '    rtn = Set_BaseMidFile_Syoyubkno(list_ittobkno, list_kbnbkno)

        '    If rtn = False Then
        '        errstr = MSG_ERR_BASEMIDFILEREAD
        '        Return rtn
        '    Else
        '        '取得したデータを元に一棟区分を設定する
        '        rtn = Set_ExistMidFile_Syoyubkno(list_ittobkno, list_kbnbkno)

        '        If rtn = False Then
        '            errstr = MSG_ERR_EXISTMIDFILEWRITE
        '            Return rtn
        '        End If
        '    End If

        '    Return rtn

        'End Function

        ' ''' <summary>
        ' ''' 汎用用中間ファイルの物件/部屋所有者情報から物件Noを取得してオブジェクトへ格納 '20160920 プログラム自動設定箇所の追加
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <param name="list_ittobkno"></param>
        ' ''' <param name="list_kbnbkno"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidFile_Syoyubkno(ByRef list_ittobkno As List(Of String), ByRef list_kbnbkno As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim obj_ittobkno As Object
        '    Dim obj_kbnbkno As Object
        '    Dim list_sheetname As New List(Of String) From {"物件所有者情報", "部屋所有者情報"}
        '    '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg sta
        '    'Dim filepath As String = MiddleDirPath & "\" & CV_FROM_MIDDLE & ".xlsx"
        '    Dim filepath As String = BaseMidDirPath & "\" & CV_FROM_MIDDLE & ".xlsx"
        '    '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg end
        '    Try

        '        'ファイル準備
        '        appli = CreateObject("Excel.Application")
        '        appli.Visible = False
        '        wbook = appli.Workbooks.Open(filepath)

        '        For Each sheetname In list_sheetname

        '            '中断処理 '20160927 キャンセル処理を追加 -add
        '            Application.DoEvents()
        '            If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '                Return rtn
        '            End If

        '            'シート選択
        '            wsheet = wbook.Worksheets(sheetname)

        '            '既存データの確認(ヘッダを含めた最大行の取得)
        '            Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        '            'ヘッダの列数取得
        '            Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        '            '行数取得(全行 - 読込開始行 + 1)
        '            Dim rowcnt As Integer = maxrowcnt - BASEMIDFILE_READWRITE_ROW + 1

        '            '作業用変数
        '            Dim tmp_obj As Object = Nothing

        '            'データが存在する場合はコピー処理を行う
        '            If rowcnt <> 0 Then

        '                'ヘッダー取得
        '                Dim headervalue As Object = wsheet.Range(wsheet.Cells(BASEMIDFILE_HEADER_ROW, 1), wsheet.Cells(BASEMIDFILE_HEADER_ROW, columncnt)).Value

        '                'ヘッダー照合
        '                Dim tasiyoheader As String = "物件No"
        '                Dim taisyohaederindex As Integer = 0

        '                For cntjj = 1 To columncnt
        '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        '                    If fldname = tasiyoheader Then
        '                        taisyohaederindex = cntjj
        '                        tmp_obj = wsheet.Range(wsheet.Cells(BASEMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value
        '                        Exit For
        '                    End If
        '                Next

        '            End If

        '            '取得した物件Noをリストオブジェクトへ格納
        '            If tmp_obj IsNot Nothing Then

        '                Dim tmp_list As New List(Of String)
        '                For cntii = 1 To tmp_obj.Length
        '                    tmp_list.Add(tmp_obj(cntii, 1))
        '                Next

        '                Select Case sheetname
        '                    Case "物件所有者情報"
        '                        list_ittobkno = tmp_list
        '                    Case "部屋所有者情報"
        '                        list_kbnbkno = tmp_list
        '                End Select

        '            End If

        '        Next

        '        'Excelファイル終了設定
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    Catch ex As Exception

        '        rtn = False
        '        'Excelファイル終了設定
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    End Try

        '    Return rtn

        'End Function

        ' ''' <summary>
        ' ''' 所有者情報から取得した物件Noを元に一棟・区分を自動設定 '20160920 プログラム自動設定箇所の追加
        ' ''' 物件詳細情報-所有者-一棟・所有区分 と 送金ルール基本情報-一所有形態区分-棟/区分 の2フィールドへ設定する
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <param name="list_ittobkno"></param>
        ' ''' <param name="list_kbnbkno"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_ExistMidFile_Syoyubkno(ByVal list_ittobkno As List(Of String), ByVal list_kbnbkno As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim matchflg As Boolean = False

        '    Dim filename As String() = New String() {"物件情報", "送金ルール情報"}
        '    Dim sheetname As String() = New String() {"物件詳細情報", "送金ルール基本情報"}
        '    Dim bknofldname As String() = New String() {"物件NO", "物件No"}
        '    Dim copyfldname As String() = New String() {"所有者-一棟・所有区分", "一所有形態区分-棟/区分"}

        '    For cntfile = 0 To UBound(filename)

        '        'Excelファイル初期設定                          
        '        rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename(cntfile), sheetname(cntfile))

        '        '中断処理 '20160927 キャンセル処理を追加 -add
        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '            Return rtn
        '        End If

        '        'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '        If rtn = False Then
        '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '            Return rtn
        '        End If

        '        'データが存在しない場合は処理を抜ける
        '        If rowcnt = 0 Then
        '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '            Return rtn
        '        End If

        '        'ヘッダー取得
        '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        '        '作業用変数
        '        Dim copycolindex As Integer = 0
        '        Dim obj_bknocol As Object = Nothing

        '        '対象ヘッダー検索
        '        For cntjj = 1 To columncnt - 1

        '            Dim fldname As String = headervalue(1, cntjj).ToString.Trim

        '            '物件No列と一棟区分設定列から必要な情報を取得
        '            Select Case fldname
        '                Case bknofldname(cntfile)
        '                    obj_bknocol = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        '                Case copyfldname(cntfile)
        '                    copycolindex = cntjj
        '            End Select

        '        Next

        '        '取得した物件Noから一棟区分を設定
        '        For cntkk = 1 To obj_bknocol.Length

        '            Dim tmp_bkno As String = obj_bknocol(cntkk, 1)
        '            If list_ittobkno.Contains(tmp_bkno) Then
        '                wsheet.Cells(cntkk + EXISTMIDFILE_READWRITE_ROW - 1, copycolindex).Value = "1"
        '            ElseIf list_kbnbkno.Contains(tmp_bkno) Then
        '                wsheet.Cells(cntkk + EXISTMIDFILE_READWRITE_ROW - 1, copycolindex).Value = "2"
        '            Else
        '                wsheet.Cells(cntkk + EXISTMIDFILE_READWRITE_ROW - 1, copycolindex).Value = "1"
        '            End If

        '        Next

        '        '終了処理
        '        Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        '    Next

        '    Return rtn

        'End Function
        '20160929 汎用→既存コピー処理改善対応 -del end

        '20160929 汎用→既存コピー処理改善対応 -del sta
        ' ''' <summary>
        ' ''' 契約No、管理Noに「1」を自動設定する '20160920 プログラム自動設定箇所の追加 -add
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidFile_DefCopyMain(ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

        '    Dim list_file As New List(Of String) From {"契約情報", "送金ルール情報", "物件情報", "部屋情報"}   '20160921 プログラム自動設定箇所の追加2 -add 
        '    Dim list_sheet_ky As New List(Of String) From _
        '        {"契約メモ情報", "契約基本情報", "契約契約者情報", "契約次回入金項目情報", "契約車情報", "契約特約事項情報", "契約入居者情報", _
        '         "契約入金項目情報", "契約変動費各戸メーター情報", "契約保険情報", "契約保証人情報", "契約履歴情報"}
        '    Dim list_sheet_so As New List(Of String) From _
        '        {"送金ルール基本情報", "送金ルール送金先情報", "送金ルール入金項目情報", "送金ルール控除項目情報"}
        '    Dim list_fldname_ky As New List(Of String) From _
        '        {"契約NO", "契約No", "契約管理レコードNo", "更新No", "改定No"}
        '    Dim fldname_so As String = "送金ルール管理No"

        '    '20160921 プログラム自動設定箇所の追加2 -add sta
        '    Dim sheetname_bksyo As String = "物件所有者情報"
        '    Dim sheetname_hysyo As String = "部屋所有者情報"
        '    Dim fldname_comsyo As String = "管理No"
        '    '20160921 プログラム自動設定箇所の追加2 -add end

        '    Try

        '        For Each filename In list_file

        '            '中断処理 '20160927 キャンセル処理を追加 -add
        '            Application.DoEvents()
        '            If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
        '                Return rtn
        '            End If

        '            Dim filepath As String = MiddleDirPath & "\" & filename & ".xlsx"

        '            'ファイル準備
        '            appli = CreateObject("Excel.Application")
        '            appli.Visible = False
        '            wbook = appli.Workbooks.Open(filepath)

        '            Select Case filename
        '                Case "契約情報"

        '                    For Each sheetname In list_sheet_ky

        '                        'シート選択
        '                        wsheet = wbook.Worksheets(sheetname)

        '                        '既存データの確認(ヘッダを含めた最大行の取得)
        '                        Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        '                        'ヘッダの列数取得
        '                        Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        '                        '行数取得(全行 - 読込開始行 + 1)
        '                        Dim rowcnt As Integer = maxrowcnt - EXISTMIDFILE_READWRITE_ROW + 1

        '                        'データが存在する場合はコピー処理を行う
        '                        If rowcnt <> 0 Then

        '                            'ヘッダー取得
        '                            Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        '                            'ヘッダー照合
        '                            For Each header In list_fldname_ky

        '                                Dim tasiyoheader As String = header
        '                                Dim taisyohaederindex As Integer = 0

        '                                For cntjj = 1 To columncnt
        '                                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        '                                    If fldname = tasiyoheader Then

        '                                        taisyohaederindex = cntjj

        '                                        '自動設定
        '                                        wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value = "1"

        '                                        Exit For

        '                                    End If
        '                                Next

        '                            Next

        '                        End If

        '                    Next

        '                Case "送金ルール情報"

        '                    For Each sheetname In list_sheet_so

        '                        'シート選択
        '                        wsheet = wbook.Worksheets(sheetname)

        '                        '既存データの確認(ヘッダを含めた最大行の取得)
        '                        Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        '                        'ヘッダの列数取得
        '                        Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        '                        '行数取得(全行 - 読込開始行 + 1)
        '                        Dim rowcnt As Integer = maxrowcnt - EXISTMIDFILE_READWRITE_ROW + 1

        '                        'データが存在する場合はコピー処理を行う
        '                        If rowcnt <> 0 Then

        '                            'ヘッダー取得
        '                            Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        '                            'ヘッダー照合
        '                            Dim tasiyoheader As String = fldname_so
        '                            Dim taisyohaederindex As Integer = 0

        '                            For cntjj = 1 To columncnt
        '                                Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        '                                If fldname = tasiyoheader Then
        '                                    taisyohaederindex = cntjj
        '                                    Exit For
        '                                End If
        '                            Next

        '                            '自動設定
        '                            wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value = "1"

        '                        End If

        '                    Next

        '                Case "物件情報", "部屋情報"     '20160921 プログラム自動設定箇所の追加2 -add

        '                    Dim tmp_sheetname As String = ""
        '                    Select Case filename
        '                        Case "物件情報"
        '                            tmp_sheetname = sheetname_bksyo
        '                        Case "部屋情報"
        '                            tmp_sheetname = sheetname_hysyo
        '                    End Select

        '                    'シート選択
        '                    wsheet = wbook.Worksheets(tmp_sheetname)

        '                    '既存データの確認(ヘッダを含めた最大行の取得)
        '                    Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        '                    'ヘッダの列数取得
        '                    Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        '                    '行数取得(全行 - 読込開始行 + 1)
        '                    Dim rowcnt As Integer = maxrowcnt - EXISTMIDFILE_READWRITE_ROW + 1

        '                    'データが存在する場合はコピー処理を行う
        '                    If rowcnt <> 0 Then

        '                        'ヘッダー取得
        '                        Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        '                        'ヘッダー照合
        '                        Dim tasiyoheader As String = fldname_comsyo
        '                        Dim taisyohaederindex As Integer = 0

        '                        For cntjj = 1 To columncnt
        '                            Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        '                            If fldname = tasiyoheader Then
        '                                taisyohaederindex = cntjj
        '                                Exit For
        '                            End If
        '                        Next

        '                        '自動設定
        '                        wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value = "1"

        '                    End If

        '            End Select

        '            'Excelファイル終了設定
        '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        '        Next

        '    Catch ex As Exception

        '        rtn = False
        '        errstr = MSG_ERR_EXISTMIDFILEWRITE
        '        'Excelファイル終了設定
        '        Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        '    End Try

        '    Return rtn

        'End Function
        '20160929 汎用→既存コピー処理改善対応 -del end

        '20160929 汎用→既存コピー処理改善対応 -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイル→既存用中間ファイル個別処理 '20160812 汎用コンバート対応
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <param name="list_cv"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidFile_ChgCopyMain(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim obj_rep As New Object
        '    Dim model_basemiditem As New Object
        '    Dim basemidfilename As String = ""
        '    Dim basemidfilesheetname As String = ""
        '    Dim existmidfilename As String = ""
        '    Dim existmidfilesheetname As String = ""

        '    For Each item In list_cv

        '        '中断処理 '20160927 キャンセル処理を追加 -add
        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        '            Return rtn
        '        End If

        '        Dim cvitem As String = item

        '        'test
        '        'Dim list As New List(Of String) From {"部屋情報-部屋特約情報"}
        '        'If list.Contains(cvitem) = False Then
        '        '    cvitem = ""
        '        'End If

        '        Dim normalflg As Boolean = True     '20160913_2 エラー時の処理を追加する修正 -add

        '        Select Case cvitem
        '            Case "部屋情報-部屋駐車場情報"
        '                obj_rep = New Njc.Repository.Hy_cyusyajo_mid_Repository
        '                basemidfilename = "中間ファイル"
        '                basemidfilesheetname = "部屋情報"
        '                existmidfilename = "部屋情報"
        '                existmidfilesheetname = "部屋駐車場情報"
        '            Case "部屋情報-部屋特約情報"
        '                obj_rep = New Njc.Repository.Hy_tokuyaku_mid_Repository
        '                basemidfilename = "中間ファイル"
        '                basemidfilesheetname = "部屋情報"
        '                existmidfilename = "部屋情報"
        '                existmidfilesheetname = "部屋特約情報"
        '            Case "部屋情報-部屋契約解約確認事項情報"
        '                obj_rep = New Njc.Repository.Hy_kykaikakuninjiko_mid_Repository
        '                basemidfilename = "中間ファイル"
        '                basemidfilesheetname = "部屋情報"
        '                existmidfilename = "部屋情報"
        '                existmidfilesheetname = "部屋契約解約確認事項情報"
        '            Case "契約情報-契約契約者情報"
        '                obj_rep = New Njc.Repository.Ky_kys_mid_Repository
        '                basemidfilename = "中間ファイル"
        '                basemidfilesheetname = "契約契約者保証人情報"
        '                existmidfilename = "契約情報"
        '                existmidfilesheetname = "契約契約者情報"
        '            Case "契約情報-契約保証人情報"
        '                obj_rep = New Njc.Repository.Ky_hosyonin_mid_Repository
        '                basemidfilename = "中間ファイル"
        '                basemidfilesheetname = "契約契約者保証人情報"
        '                existmidfilename = "契約情報"
        '                existmidfilesheetname = "契約保証人情報"

        '            Case "契約情報-契約特約事項情報"
        '                obj_rep = New Njc.Repository.Ky_tokuyaku_mid_Repository
        '                basemidfilename = "中間ファイル"
        '                basemidfilesheetname = "契約特約およびメモ情報"
        '                existmidfilename = "契約情報"
        '                existmidfilesheetname = "契約特約事項情報"

        '            Case Else
        '                obj_rep = Nothing
        '        End Select

        '        If obj_rep IsNot Nothing Then
        '            '20160913_2 エラー時の処理を追加する修正 -chg sta
        '            ''汎用用中間ファイル読込
        '            'obj_rep.Read_BaseMidFile(basemidfilename, basemidfilesheetname, model_basemiditem)

        '            ''既存用中間ファイル書込
        '            'obj_rep.Set_ExistMidFile(existmidfilename, existmidfilesheetname, model_basemiditem)

        '            '汎用用中間ファイル読込
        '            normalflg = obj_rep.Read_BaseMidFile(basemidfilename, basemidfilesheetname, model_basemiditem)

        '            If normalflg Then

        '                '既存用中間ファイル書込
        '                normalflg = obj_rep.Set_ExistMidFile(existmidfilename, existmidfilesheetname, model_basemiditem)

        '                If normalflg = False Then
        '                    errstr = MSG_ERR_EXISTMIDFILEWRITE
        '                    rtn = False
        '                    Return rtn
        '                End If
        '            Else
        '                errstr = MSG_ERR_BASEMIDFILEREAD
        '                rtn = False
        '                Return rtn
        '            End If
        '            '20160913_2 エラー時の処理を追加する修正 -chg end
        '        End If

        '    Next

        '    Return rtn

        'End Function
        '20160929 汎用→既存コピー処理改善対応 -del end

        '20160929 汎用→既存コピー処理改善対応 -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイルからデータを取得してオブジェクトへ格納 20160905 中間ファイルコピー処理改善
        ' ''' 20160913_2 エラー時の処理を追加する修正 引数「errstr」を削除
        ' ''' </summary>
        ' ''' <param name="midfilename"></param>
        ' ''' <param name="midsheetname"></param>
        ' ''' <param name="middata"></param>
        ' ''' <param name="hash_middata"></param>
        ' ''' <param name="writerowcnt"></param>
        ' ''' <param name="existsheetname"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidDataToObj(ByVal midfilename As String, ByVal midsheetname As String, ByRef middata As Object, ByRef hash_middata As Hashtable, ByRef writerowcnt As Integer, _
        '                                      ByVal existsheetname As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim matchflg As Boolean = False
        '    Dim headerrow As Integer = 2                                    '20160912 中間ファイル作成に伴うプログラム修正 -add

        '    'Excelファイル初期設定
        '    '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg sta
        '    'rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, midfilename, midsheetname)
        '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, BaseMidDirPath, midfilename, midsheetname)
        '    '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg end
        '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '    If rtn = False Then
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '        Return rtn
        '    End If

        '    'データが存在しない場合は処理を抜ける
        '    If rowcnt = 0 Then
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160905 Excel終了処理の修正 -add
        '        Return rtn
        '    End If

        '    'ヘッダー取得
        '    '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        '    'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
        '    Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value
        '    '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        '    '対象ヘッダー検索
        '    For cntjj = 1 To columncnt - 1

        '        '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        '        'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim

        '        'Dim keysfname As String = midsheetname & "-" & fldname

        '        ''中間ファイルヘッダー存在有無確認
        '        'If List_Basemidheader.Contains(keysfname) Then

        '        '    '紐付く既存中間ファイルヘッダー有無確認
        '        '    If Hash_BaseMidToExistMid_SH.Contains(keysfname) Then
        '        '        Dim tmp_existstr() As String = Hash_BaseMidToExistMid_SH(keysfname).ToString.Split("-")
        '        '        If existsheetname = tmp_existstr(0) Then
        '        '            middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        '        '            hash_middata.Add(keysfname, middata)
        '        '        End If
        '        '    End If

        '        'End If
        '        Dim fldname As String = headervalue(1, cntjj).ToString.Trim

        '        If fldname <> "項目名" Then

        '            Dim keysfname As String = midsheetname & STR_SPLIT_1 & fldname            '20160913 「@#@」を共通変数にして下さい(他の箇所も同様)

        '            '20160926 選定した移行項目をプログラムへ反映する修正 -chg sta
        '            ''中間ファイルヘッダー存在有無確認
        '            'If List_Basemidheader.Contains(keysfname) Then

        '            '    '紐付く既存中間ファイルヘッダー有無確認
        '            '    If Hash_BaseMidToExistMid_SH.Contains(keysfname) Then
        '            '        Dim tmp_existstr() As String = Split(Hash_BaseMidToExistMid_SH(keysfname).ToString, STR_SPLIT_1)
        '            '        If existsheetname = tmp_existstr(0) Then
        '            '            '20160915 コピー処理のずれを修正 -chg sta
        '            '            'middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        '            '            middata = wsheet.Range(wsheet.Cells(startrow, cntjj + 1), wsheet.Cells(maxrowcnt, cntjj + 1)).Value
        '            '            '20160915 コピー処理のずれを修正 -chg end
        '            '            hash_middata.Add(keysfname, middata)
        '            '        End If
        '            '    End If

        '            'End If

        '            '移行対象フラグ確認
        '            If List_BaseMiddatacv.Contains(keysfname) Then

        '                '中間ファイルヘッダー存在有無確認
        '                If List_Basemidheader.Contains(keysfname) Then

        '                    '紐付く既存中間ファイルヘッダー有無確認
        '                    If Hash_BaseMidToExistMid_SH.Contains(keysfname) Then
        '                        Dim tmp_existstr() As String = Split(Hash_BaseMidToExistMid_SH(keysfname).ToString, STR_SPLIT_1)
        '                        If existsheetname = tmp_existstr(0) Then
        '                            '20160915 コピー処理のずれを修正 -chg sta
        '                            'middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        '                            middata = wsheet.Range(wsheet.Cells(startrow, cntjj + 1), wsheet.Cells(maxrowcnt, cntjj + 1)).Value
        '                            '20160915 コピー処理のずれを修正 -chg end
        '                            hash_middata.Add(keysfname, middata)
        '                        End If
        '                    End If

        '                End If

        '            End If
        '            '20160926 選定した移行項目をプログラムへ反映する修正 -chg end
        '        End If
        '        '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        '    Next

        '    '書込行数を返却
        '    '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        '    'writerowcnt = maxrowcnt
        '    writerowcnt = rowcnt
        '    '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    Return rtn

        'End Function

        ' ''' <summary>
        ' ''' 取得した汎用用中間ファイルデータを既存用中間ファイルへコピー '20160905 中間ファイルコピー処理改善
        ' ''' </summary>
        ' ''' <param name="existmidfilename"></param>
        ' ''' <param name="existmidsheetname"></param>
        ' ''' <param name="hash_basemiddata"></param>
        ' ''' <param name="writerowcnt"></param>
        ' ''' <param name="iniflg"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidDataToExistMidFile(ByVal existmidfilename As String, ByVal existmidsheetname As String, ByVal hash_basemiddata As Hashtable, ByVal writerowcnt As Integer, ByVal iniflg As Boolean) As Boolean

        '    '変数説明
        '    'basemidfsname		    	… 画面上(UIチェックボックス)の移行項目からオープンする汎用用中間ファイル名を取得
        '    'List_BaseMiddatacv		    … 移行項目選定対応：汎用用に選定した項目を格納
        '    'Hash_BaseMidToExistMid_FS	… 汎用と既存のファイルとシート結合データ(中間ファイル.物件情報-物件情報.物件基本情報)
        '    'Hash_ExistMidToBaseMid_FS	… 既存と汎用のファイルとシート結合データ(物件情報.物件基本情報-中間ファイル.物件情報)
        '    'Hash_BaseMidToExistMid_SH	… 汎用と既存のシートと項目名の結合データ(物件情報.物件No-物件基本情報.物件No)
        '    '                               → 汎用用中間ファイルデータを作業用のオブジェクトへ格納 
        '    'Hash_ExistMidToBaseMid_SH	… 既存と汎用のシートと項目名の結合データ(物件基本情報.物件No-物件情報.物件No)
        '    '                               → オブジェクトへ格納した汎用用中間ファイルデータを既存U用中間ファイルへ貼り付けるためのハッシュ

        '    Dim rtn As Boolean = True
        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim fldmatchflg As Boolean = False

        '    '************************
        '    '作業準備
        '    '************************

        '    'Excelファイル初期設定(既存用中間ファイルの初期化有無で処理を分岐させる)
        '    If iniflg Then
        '        rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        '    Else
        '        rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        '    End If

        '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '    If rtn = False Then                             '20160913 どこそこで「rtn」を返しているが、エラーの際、ソース上階層に移った時に何のエラーかが不明(エラーの際の処理が何もされていない？)
        '        Return rtn
        '    End If
        '    '20160905 中間ファイルが空の場合の共通コピー処理修正 -add sta
        '    If hash_basemiddata.Count = 0 Then
        '        Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
        '        Return rtn
        '    End If
        '    '20160905 中間ファイルが空の場合の共通コピー処理修正 -add end

        '    '列数を退避 '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -add
        '    Dim writecol As Integer = columncnt

        '    'ヘッダー取得
        '    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
        '    '20160926 選定した移行項目をプログラムへ反映する修正 -del sta
        '    '貼付処理を修正するためコメントアウト
        '    ''汎用用中間ファイルから取得したデータを元に貼付け先を取得
        '    'For Each basemiddata In hash_basemiddata

        '    '    Dim basemidfldname As String = basemiddata.Key
        '    '    Dim basemidcoldata As Object = basemiddata.Value

        '    '    '照合→コピー処理
        '    '    If Hash_BaseMidToExistMid_SH.Contains(basemidfldname) Then

        '    '        Dim existmiddata() As String = Split(Hash_BaseMidToExistMid_SH(basemidfldname).ToString, STR_SPLIT_1)
        '    '        Dim existmidfldname As String = existmiddata(1)

        '    '        '取得した列名を元にコピー処理
        '    '        '一致するフィールドが存在する場合
        '    '        For cntjj = 1 To columncnt
        '    '            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        '    '            If fldname = existmidfldname Then
        '    '                '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        '    '                'wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(writerowcnt, cntjj)).Value = basemidcoldata
        '    '                wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(startrow + writerowcnt - 1, cntjj)).Value = basemidcoldata     '20160913 ここで書き出さずに、オブジェクト配列とかに格納して最後に纏めて登録できない？
        '    '                '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        '    '                fldmatchflg = True
        '    '                Exit For
        '    '            End If
        '    '        Next

        '    '        '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
        '    '        If fldmatchflg = False Then
        '    '            '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg sta
        '    '            'wsheet.Cells(startrow - 1, columncnt + 1).Value = existmidfldname
        '    '            'wsheet.Range(wsheet.Cells(startrow, columncnt + 1), wsheet.Cells(writerowcnt, columncnt + 1)).Value = basemidcoldata
        '    '            writecol = writecol + 1
        '    '            wsheet.Cells(startrow - 1, writecol).Value = existmidfldname
        '    '            wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = basemidcoldata
        '    '            '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg end
        '    '        End If

        '    '    End If

        '    'Next
        '    '20160926 選定した移行項目をプログラムへ反映する修正 -del end

        '    '20160926 選定した移行項目をプログラムへ反映する修正 -add sta
        '    For Each existmid In Hash_ExistMidToBaseMid_SH

        '        '既存中間ファイルと汎用中間ファイルの紐付状況を取得
        '        Dim existmidshname As String = existmid.Key
        '        Dim basemidshname As String = existmid.Value

        '        '汎用中間ファイルから取得したデータを照合し、一致する列データを取得する
        '        If basemidshname <> "" Then

        '            Dim basemidfldname As String = ""
        '            Dim basemidcoldata As Object = Nothing

        '            For Each basemiddata In hash_basemiddata

        '                basemidfldname = basemiddata.Key
        '                basemidcoldata = basemiddata.Value

        '                If basemidfldname = basemidshname Then

        '                    '一致した列データが存在する場合は貼付け処理を行う
        '                    Dim existmiddata() As String = Split(existmidshname, STR_SPLIT_1)
        '                    Dim existmidfldname As String = existmiddata(1)

        '                    For cntjj = 1 To columncnt
        '                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        '                        If fldname = existmidfldname Then
        '                            wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(startrow + writerowcnt - 1, cntjj)).Value = basemidcoldata
        '                            fldmatchflg = True
        '                            Exit For
        '                        End If
        '                    Next

        '                    '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
        '                    If fldmatchflg Then
        '                        Exit For
        '                    Else
        '                        writecol = writecol + 1
        '                        wsheet.Cells(startrow - 1, writecol).Value = existmidfldname
        '                        wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = basemidcoldata
        '                    End If

        '                End If

        '            Next

        '        End If

        '    Next
        '    '20160926 選定した移行項目をプログラムへ反映する修正 -add end

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        '    Return rtn

        'End Function
        '20160929 汎用→既存コピー処理改善対応 -del end

        '20160905 中間ファイルコピー処理改善 -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイル→既存用中間ファイル通常処理 '20160812 汎用コンバート対応
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidFile_Copy(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim readtbl As New DataTable
        '    Dim reccnt As Integer
        '    Dim fldname As String
        '    Dim fldvalue As String
        '    Dim tmp_basemidfilename As String = ""
        '    Dim tmp_basemidfilesheetname As String = ""
        '    Dim tmp_basemidfilefldnamemain As String = ""
        '    Dim tmp_basemidfilefldnamesub As String = ""
        '    Dim tmp_existmidfilename As String = ""
        '    Dim tmp_existmidfilesheetname As String = ""
        '    Dim tmp_existmidfilefldname As String = ""
        '    Dim tmp_copyflg As String = ""
        '    Dim list_existmidfileinfo As New List(Of String)
        '    Dim basemidfilematchflg As Boolean = True

        '    '中間ファイル紐付情報取得
        '    Dim tmp_sql As String = " SELECT DISTINCT * FROM tmp_midfileinfo ORDER BY [汎用中間ファイル名],[汎用中間シート名] "
        '    reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl)

        '    '読込開始
        '    For cntii As Integer = 0 To readtbl.Rows.Count - 1

        '        '中断処理
        '        Application.DoEvents()
        '        If CancelFlg Then
        '            Exit Function
        '        End If

        '        '既存用中間ファイル初期化フラグ
        '        Dim iniflg As Boolean = True

        '        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

        '            '項目名取得
        '            fldname = readtbl.Columns(cntjj).ColumnName

        '            '登録値取得
        '            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

        '            '各項目値→変数格納
        '            Select Case fldname
        '                Case "既存中間ファイル名"
        '                    tmp_existmidfilename = fldvalue.Trim
        '                Case "既存中間シート名"
        '                    tmp_existmidfilesheetname = fldvalue.Trim
        '                Case "既存中間フィールド名"
        '                    tmp_existmidfilefldname = fldvalue.Trim
        '                Case "汎用中間ファイル名"
        '                    tmp_basemidfilename = fldvalue.Trim
        '                Case "汎用中間シート名"
        '                    tmp_basemidfilesheetname = fldvalue.Trim
        '                Case "汎用中間フィールド名_大分類"
        '                    tmp_basemidfilefldnamemain = fldvalue.Trim
        '                Case "汎用中間フィールド名_小分類"
        '                    tmp_basemidfilefldnamesub = fldvalue.Trim
        '                Case "コピーフラグ"
        '                    tmp_copyflg = fldvalue.Trim
        '            End Select

        '        Next

        '        '処理対象の場合はコピー処理を実行
        '        Dim cvtaisyo As String = tmp_existmidfilename & "-" & tmp_existmidfilesheetname

        '        If list_cv.Contains(cvtaisyo) Then

        '            'そのままコピーできるデータの処理
        '            If tmp_copyflg = "●" Then

        '                '開くファイル名とシート名を元に初期化フラグを設定(一度開いたシート名をオブジェクトへ退避させておく)
        '                Dim existmidfileinfo As String = tmp_existmidfilename & "-" & tmp_existmidfilesheetname
        '                If list_existmidfileinfo.Contains(existmidfileinfo) = False Then
        '                    iniflg = True
        '                    list_existmidfileinfo.Add(existmidfileinfo)
        '                Else
        '                    iniflg = False
        '                End If

        '                '汎用用中間ファイルからデータ取得
        '                Dim middata As New Object
        '                basemidfilematchflg = Me.Set_BaseMidDataToObj(tmp_basemidfilename, tmp_basemidfilesheetname, tmp_basemidfilefldnamemain, middata, errstr)

        '                '汎用用中間ファイルのヘッダーが一致しなかった場合(編集されている場合)は処理を抜ける
        '                If basemidfilematchflg = False Then
        '                    errstr = "中間ファイルの形式が変更された可能性があるためコンバート処理を続行することができません。"
        '                    rtn = basemidfilematchflg
        '                    Return rtn
        '                End If

        '                '既存用中間ファイルへコピー
        '                Call Me.Set_BaseMidDataToExistMidFile(tmp_existmidfilename, tmp_existmidfilesheetname, tmp_existmidfilefldname, middata, iniflg)

        '            End If

        '        End If

        '    Next

        '    Return rtn

        'End Function
        '20160905 中間ファイルコピー処理改善 -del end

        '20160905 中間ファイルコピー処理改善 -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイルからデータを取得してオブジェクトへ格納 '20160812 汎用コンバート対応
        ' ''' </summary>
        ' ''' <param name="midfilename"></param>
        ' ''' <param name="midsheetname"></param>
        ' ''' <param name="midfldname"></param>
        ' ''' <param name="middata"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidDataToObj(ByVal midfilename As String, ByVal midsheetname As String, ByVal midfldname As String, ByRef middata As Object, ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim matchflg As Boolean = False

        '    'Excelファイル初期設定                          
        '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, midfilename, midsheetname)

        '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '    If rtn = False Then
        '        Return rtn
        '    End If

        '    'ヘッダー取得
        '    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

        '    '対象ヘッダー検索
        '    For cntjj = 1 To columncnt
        '        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        '        If fldname = midfldname Then
        '            middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        '            matchflg = True
        '            Exit For
        '        End If
        '    Next

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    '照合フラグをセット
        '    If matchflg = False Then
        '        rtn = matchflg
        '    End If

        '    Return rtn

        'End Function
        '20160905 中間ファイルコピー処理改善 -del end

        '20160905 中間ファイルコピー処理改善 -del sta
        ' ''' <summary>
        ' ''' 取得した汎用用中間ファイルデータを既存用中間ファイルへコピー '20160812 汎用コンバート対応
        ' ''' </summary>
        ' ''' <param name="existmidfilename"></param>
        ' ''' <param name="existmidsheetname"></param>
        ' ''' <param name="existmidfldname"></param>
        ' ''' <param name="middata"></param>
        ' ''' <param name="iniflg"></param>
        ' ''' <param name="basemidfilename"></param>
        ' ''' <param name="basemidsheetname"></param>
        ' ''' <param name="basemidfldname"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidDataToExistMidFile(ByVal existmidfilename As String, ByVal existmidsheetname As String, ByVal existmidfldname As String, ByRef middata As Object, ByVal iniflg As Boolean) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim fldmatchflg As Boolean = False

        '    '************************
        '    '作業準備
        '    '************************

        '    'Excelファイル初期設定(既存用中間ファイルの初期化有無で処理を分岐させる)
        '    If iniflg Then
        '        rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        '    Else
        '        rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        '    End If

        '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '    If rtn = False Then
        '        Return rtn
        '    End If

        '    'ヘッダー取得
        '    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

        '    'コピー件数設定
        '    Dim writerowcnt As Integer = middata.Length

        '    '対象ヘッダー検索→一致した列にコピー
        '    For cntjj = 1 To columncnt
        '        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        '        If fldname = existmidfldname Then
        '            wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells((startrow - 1) + writerowcnt, cntjj)).Value = middata
        '            fldmatchflg = True
        '            Exit For
        '        End If
        '    Next

        '    '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
        '    If fldmatchflg = False Then
        '        wsheet.Cells(startrow - 1, columncnt + 1).Value = existmidfldname
        '        wsheet.Range(wsheet.Cells(startrow, columncnt + 1), wsheet.Cells((startrow - 1) + writerowcnt, columncnt + 1)).Value = middata
        '    End If

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        '    Return rtn

        'End Function
        '20160905 中間ファイルコピー処理改善 -del end

        '20160929 汎用→既存コピー処理改善対応 -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイル→既存用中間ファイル通常処理 '20160905 中間ファイルコピー処理改善
        ' ''' </summary>
        ' ''' <param name="errstr"></param>
        ' ''' <param name="list_cv"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Set_BaseMidFile_Copy(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim normalflg As Boolean = True
        '    Dim list_existmidfileinfo As New List(Of String)

        '    For Each cvitem In list_cv

        '        ''そのままコピー可判別用
        '        'Dim copyflg As Boolean = False

        '        '中断処理 '20160927 キャンセル処理を追加 -add
        '        Application.DoEvents()
        '        If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        '            Return rtn
        '        End If

        '        'コンバート対象項目からファイル名とシート名を取得
        '        cvitem = cvitem.Replace("-", STR_SPLIT_1)
        '        Dim tmp_existmidstr() As String = Split(cvitem, STR_SPLIT_1)
        '        Dim existfilename As String = tmp_existmidstr(0)
        '        Dim existsheetname As String = tmp_existmidstr(1)

        '        '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        '        Call Me.Set_MidDataInfoToObj(existfilename, existsheetname, False)

        '        '移行項目からオープンする汎用用中間ファイル名を取得
        '        Dim basemidfsname As String = ""
        '        If Hash_ExistMidToBaseMid_FS.Contains(cvitem) Then

        '            basemidfsname = Hash_ExistMidToBaseMid_FS(cvitem)
        '            'copyflg = True

        '            'test
        '            'Dim list As New List(Of String) From {"物件情報@#@物件交通情報"}
        '            'If list.Contains(cvitem) = False Then
        '            '    basemidfsname = ""
        '            'End If

        '            If basemidfsname <> "" Then

        '                Dim tmp_basemidstr() As String = Split(basemidfsname, STR_SPLIT_1)
        '                Dim basefilename As String = tmp_basemidstr(0)
        '                Dim basesheetname As String = tmp_basemidstr(1)

        '                '汎用用中間ファイルからデータ取得
        '                Dim middata As New Object
        '                Dim hash_middata As New Hashtable
        '                Dim writerowcnt As Integer = 0
        '                '20160913_2 エラー時の処理を追加する修正 -chg sta
        '                'normalflg = Me.Set_BaseMidDataToObj(basefilename, basesheetname, middata, hash_middata, writerowcnt, existsheetname, errstr)
        '                normalflg = Me.Set_BaseMidDataToObj(basefilename, basesheetname, middata, hash_middata, writerowcnt, existsheetname)
        '                '20160913_2 エラー時の処理を追加する修正 -chg end
        '                '既存用中間ファイルへコピー
        '                If normalflg Then

        '                    '開くファイル名とシート名を元に初期化フラグを設定(一度開いたシート名をオブジェクトへ退避させておく)
        '                    Dim iniflg As Boolean = False
        '                    If list_existmidfileinfo.Contains(cvitem) = False Then
        '                        list_existmidfileinfo.Add(cvitem)
        '                        iniflg = True
        '                    End If

        '                    'コピー処理
        '                    '20160913_2 エラー時の処理を追加する修正 -chg sta
        '                    'Call Me.Set_BaseMidDataToExistMidFile(existfilename, existsheetname, hash_middata, writerowcnt, iniflg)
        '                    normalflg = Me.Set_BaseMidDataToExistMidFile(existfilename, existsheetname, hash_middata, writerowcnt, iniflg)
        '                    If normalflg = False Then
        '                        errstr = MSG_ERR_EXISTMIDFILEWRITE
        '                        rtn = False
        '                        Return rtn
        '                    End If
        '                    '20160913_2 エラー時の処理を追加する修正 -chg end
        '                Else    '20160913_2 エラー時の処理を追加する修正 -add
        '                    errstr = MSG_ERR_BASEMIDFILEREAD
        '                    rtn = False
        '                    Return rtn

        '                End If

        '            End If

        '        End If

        '    Next

        '    Return rtn

        'End Function
        '20160929 汎用→既存コピー処理改善対応 -del end

        ''' <summary>
        ''' ダミー用の金融機関9999を作成する処理
        ''' </summary>
        ''' <returns>終了判定 True:正常終了 False:異常終了</returns>
        ''' <remarks>
        ''' 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 新規追加
        ''' </remarks>
        Private Function Set_DummyBankData() As Boolean

            Dim xmlupdatestr As String = "<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "utf-16" & """" & "?>"  '挿入用に成形するための文字列(除去用)
            Dim tmp_cnt As Integer = 0  'クエリ実行用の作業用変数
            Dim tmp_sql As String = ""
            Dim rtn As Boolean = True

            'ダミー銀行有無確認
            tmp_sql = " SELECT COUNT(*) FROM m_kinyu WHERE kinyu_no = 9999 "
            Dim tmp_kinyucnt As String = DBExec.Exec_Scalar(tmp_sql, Me.sqlcnnv10)
            tmp_sql = ""

            'ダミー銀行挿入
            If tmp_kinyucnt = "0" Then
                tmp_sql = " INSERT INTO m_kinyu VALUES(9999,'ダミー銀行','','',1,NULL," & "'" & DefHistory.Replace(xmlupdatestr, "") & "'" & ",1,NULL) "
                rtn = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
                tmp_sql = ""
                If rtn = False Then
                    Return rtn
                End If
            End If

            Return rtn

        End Function

#End Region

#Region "初期設定処理"

        ''' <summary>
        ''' コマンドライン値取得・条件設定
        ''' </summary>
        ''' <remarks>
        '''  ・コンバートタイプ判別
        '''  　第1引数…コンバートタイプ (0.汎用, 1.既存ユーザ用(V7), 2.他社システム用, 左記以外.既存ユーザ用(DEF))
        '''  　第2引数…_"njc_dev".開発用機能表示
        '''  　第3引数…ユーザ名またはシステム名 [呼出起動時または他社システム用の場合に使用]
        '''  　※引数なし…P1.既存ユーザ用＋P2.開発用処理なし
        ''' 
        ''' ・パラメータ渡し例
        ''' 　1 njc_dev  … 既存ユーザ用＋開発用機能表示
        ''' 　0          … 汎用＋開発用機能なし
        ''' 　2 0 nissin … 他社システム用＋開発用機能なし＋日新不動産様用
        ''' </remarks>
        Private Sub IniCmdline()

            Dim cmdline() As String = System.Environment.GetCommandLineArgs()
            Dim cmdlinecnt As Integer = UBound(cmdline)

            For cntii = 0 To UBound(cmdline)
                Select Case cntii
                    Case 0
                        'なし
                    Case 1
                        'コンバートタイプ取得
                        '20160926 コンバーター起動時パラメータの仕様変更修正 -del sta
                        'Dim tmp_cnvnostr As String = cmdline(cntii).ToString
                        'Dim tmp_cnvnoint As Integer = 0
                        'If (Int32.TryParse(tmp_cnvnostr, tmp_cnvnoint)) = False Then
                        '    tmp_cnvnoint = 1
                        'End If
                        'CNVNO = tmp_cnvnoint
                        '20160926 コンバーター起動時パラメータの仕様変更修正 -del end
                    Case 2
                        '開発用機能表示
                        If cmdline(cntii) = NJC_DEV Then
                            Dev_CVFlg = True
                        End If
                End Select
            Next

            '20160926 コンバーター起動時パラメータの仕様変更修正 -add sta
            '★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
            ' 確認メモ
            ' →既存・汎用・他社の各々で固定値をハードコーディングする
            'CNVNO = ConvertTypes._既存ユーザ用
            CNVNO = ConvertTypes._汎用
            '★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
            '20160926 コンバーター起動時パラメータの仕様変更修正 -add end

        End Sub

        ''' <summary>
        ''' メイン.開始タブ設定(1)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMStart()

            Me.lblKFirstLabel.Visible = False
            Me.lblHFirstLabel.Visible = False
            Me.lblHFirstLabel.Visible = False
            Me.lblKFirstLabel.Visible = False
            Me.grpHFirstNaiyo.Visible = False
            Me.grpKFirstNaiyo.Visible = False
            Me.lblTitleKi.Visible = False
            Me.lblTitleH.Visible = False
            Me.lblHFirstLabel.Location = New Point(Me.lblKFirstLabel.Location)
            Me.grpHFirstNaiyo.Location = New Point(Me.grpKFirstNaiyo.Location)
            Me.lblTitleH.Location = New Point(Me.lblTitleKi.Location)

            Select Case CNVNO
                Case ConvertTypes._汎用
                    Me.lblHFirstLabel.Visible = True
                    Me.grpHFirstNaiyo.Visible = True
                    Me.grpHFirstNaiyo.Visible = True
                    Me.lblTitleH.Visible = True
                Case Else
                    Me.lblKFirstLabel.Visible = True
                    Me.grpKFirstNaiyo.Visible = True
                    Me.grpKFirstNaiyo.Visible = True
                    Me.lblTitleKi.Visible = True
            End Select

        End Sub

        ''' <summary>
        ''' メイン.接続設定タブ設定(2)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMSession()

            '--------------------------------------------------
            ' 内容説明ラベル
            '--------------------------------------------------
            Me.lblHSessionDescription1.Visible = False
            Me.lblKSessionDescription1.Visible = False
            Me.lblHSessionDescription1.Location = New Point(Me.lblKSessionDescription1.Location)
            Select Case CNVNO
                Case ConvertTypes._汎用
                    Me.lblHSessionDescription1.Visible = True
                Case Else
                    Me.lblKSessionDescription1.Visible = True
            End Select


            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            '初期接続情報読込ボタン制御
            Me.btnDefConInfoRead.Visible = False

            '接続値取得(V7.レジストリ, 10.XML設定ファイル)
            If Me.Set_DBExistConInfo() = False Then
                Call Me.Set_V7DBConInfo()
                Call Me.Set_10DBConInfo()
            Else
                Me.btnDefConInfoRead.Visible = True
            End If

            '接続値セット
            Me.Set_DefConInfo_To_Control(fstmodelv7, fstmodelv10)

            'パスワード入力ボックスマスク設定
            Me.txtV7Pass.PasswordChar = "*"c
            Me.txtV10Pass.PasswordChar = "*"c

            '※ネットワークライブラリはとりあえず非表示
            Me.lblNetworklib.Visible = False
            'Me.txtV7Networklib.Visible = False
            'Me.txtV10Networklib.Visible = False

            '(開発用) タイムアウト値 
            Me.txtTimeOut.MaxLength = 3
            Me.grpTimeOut.Visible = Dev_CVFlg

            '20160711 ネットワークライブラリ対応 -add sta
            'ネットワークライブラリ
            Dim networklibrary As String() = New String() {"名前付きパイプ", "共有メモリ", "TCP/IP"}
            Me.cmbV7Networklib.Items.Add("")
            Me.cmbV7Networklib.Items.Add(networklibrary(0))
            Me.cmbV7Networklib.Items.Add(networklibrary(1))
            Me.cmbV7Networklib.Items.Add(networklibrary(2))
            Me.cmbV10Networklib.Items.Add("")
            Me.cmbV10Networklib.Items.Add(networklibrary(0))
            Me.cmbV10Networklib.Items.Add(networklibrary(1))
            Me.cmbV10Networklib.Items.Add(networklibrary(2))

            '※構築中のため非表示
            Me.cmbV7Networklib.Visible = False
            Me.cmbV10Networklib.Visible = False
            '20160711 ネットワークライブラリ対応 -add end

            '(汎用) 移行先のみ表示
            If CNVNO = ConvertTypes._汎用 Then
                grpV7ConnectInfo.Visible = False
                grp10ConnectInfo.Location = grpV7ConnectInfo.Location
            End If

        End Sub

        ''' <summary>
        ''' メイン.初期設定タブ設定(3)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMSyoki()

            '--------------------------------------------------
            ' 仮処理 (未定処理)
            '--------------------------------------------------
            '20161018 UI最終確認での修正 -chg sta
            'Me.pnlCVType.Visible = False
            Me.pnlConvertType.Visible = False
            '20161018 UI最終確認での修正 -chg end
            '--------------------------------------------------
            ' 内容説明ラベル
            '--------------------------------------------------
            Me.pnlOptSelect.Visible = False
            Me.lblHOpSeleTitle.Visible = False
            Me.lblHOpSeleNaiyo.Visible = False
            Me.lblKOpSeleTitle.Visible = False
            Me.lblKOpSeleNaiyo.Visible = False
            Me.lblSOpSeleNaiyo.Visible = False
            Me.pnlOptionSelect.Visible = False
            Me.lblHOpSeleTitle.Location = New Point(Me.lblKOpSeleTitle.Location)
            Me.lblHOpSeleNaiyo.Location = New Point(Me.lblKOpSeleNaiyo.Location)
            Select Case CNVNO
                Case ConvertTypes._汎用
                    '20161209 未使用の運用開始年月設定箇所を非表示にする修正 -chg sta
                    '↓↓↓旧srcコメントアウト↓↓↓
                    ''20160928 初期設定位置調整 -chg sta
                    ' ''Me.lblHOpSeleTitle.Visible = True                                      
                    ' ''Me.lblHOpSeleNaiyo.Visible = True                                      
                    ' ''20160921 同意内容修正 -add sta
                    ''Me.pnlLogPath.Location = New Point(Me.pnlOptSelect.Location)            
                    ''Me.pnlUserName.Location = New Point(46, 260)
                    ' ''20160921 同意内容修正 -add end
                    'Me.pnlUnyoKaisi.Location = New Point(68, 107)
                    'Me.pnlLogPath.Location = New Point(68, 215)
                    'Me.pnlUserName.Location = New Point(68, 323)
                    ''20160928 初期設定位置調整 -chg end
                    '↑↑↑旧srcコメントアウト↑↑↑
                    Me.pnlUnyoKaisi.Visible = False
                    Me.pnlLogPath.Location = New Point(68, 107)
                    Me.pnlUserName.Location = New Point(68, 215)
                    '20161209 未使用の運用開始年月設定箇所を非表示にする修正 -chg end

                Case Else
                    Me.pnlOptSelect.Visible = True											 '20160921 同意内容 -add
                    Me.lblKOpSeleTitle.Visible = True
                    Me.lblKOpSeleNaiyo.Visible = True
                    Me.lblSOpSeleNaiyo.Visible = True
                    Me.pnlOptionSelect.Visible = True
            End Select


            '--------------------------------------------------
            ' 初期値セット
            '--------------------------------------------------
            '運用開始年月
            Dim cvym As String = Now.ToString("yyyy/MM")                'コンバーター実行ファイル起動時のシステム年月を設定
            Me.txtUnyoYYYYMM.Mask = "0000/00"                           '書式設定
            Me.txtUnyoYYYYMM.Text = cvym
            '作業者名
            Me.txtRecUser.Text = DEF_REC_USER


            '--------------------------------------------------
            ' パス設定
            '--------------------------------------------------
            'コンバーターログファイルパス
            Dim tmp_deflogdirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_MAINLOG_NAME)
            Me.txtLogDirPath.Text = tmp_deflogdirpath


            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            'コンバートタイプ初期セット
            Me.optCVNew.Checked = True
            Me.optCVAdd.Checked = False

            'コンバーターログファイル
            Me.txtLogDirPath.BackColor = Color.White
            Me.txtLogDirPath.ReadOnly = True
            Me.txtLogDirPath.ForeColor = Color.Gray

            '中間ファイル関連
            Me.txtMidDirLogPath.BackColor = Color.White
            Me.txtMidDirLogPath.ReadOnly = True
            Me.txtMidDirLogPath.ForeColor = Color.Gray

            '(開発用)
            btnDevTabChange.Visible = Dev_CVFlg

            '紐付設定ファイル関連
            Me.txtRelationDirPath.BackColor = Color.White
            Me.txtRelationDirPath.ReadOnly = True
            Me.txtRelationDirPath.ForeColor = Color.Gray

        End Sub

        ''' <summary>
        ''' メイン.コンバート作業選択タブ(4)
        ''' </summary>
        ''' <remarks>
        ''' ・コンバート作業選択
        ''' </remarks>
        Private Sub IniFormtabPageMMenu()

            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            Me.btnMenuJizen.Enabled = True
            Me.btnMenuDatacv.Enabled = False
            Me.btnMenuGazocv.Enabled = True
            Me.btnMenuJigo.Enabled = True
            Me.btnMenuHojyo.Enabled = True
            btnNext.Enabled = False

            lblHMenuDatacv.Location = New Point(lblKMenuDatacv.Location)
            Me.lblHMenuDatacv.Visible = False
            Me.lblKMenuDatacv.Visible = False
            Select Case CNVNO
                Case ConvertTypes._汎用
                    '20161018 UI最終確認での修正 -add sta
                    lblHMenuJizen.Location = New Point(lblKMenuJizen.Location)
                    lblKMenuJizen.Visible = False
                    lblHMenuJizen.Visible = True
                    '20161018 UI最終確認での修正 -add end
                    Me.lblHMenuDatacv.Visible = True
                    grpMenuHojyo.Visible = False

                    '画像コンバートボタン除去
                    btnMenuGazocv.Enabled = False

                    grpMenuHojyo.Location = New Point(grpMenuJigo.Location)
                    grpMenuHojyo.Text = " ※補助機能 "
                    grpMenuJigo.Location = New Point(grpMenuGazocv.Location)
                    grpMenuJigo.Text = " 3. 事後作業 "
                    grpMenuGazocv.Visible = False
                Case Else
                    Me.lblKMenuDatacv.Visible = True
            End Select

        End Sub

        ''' <summary>
        ''' メイン.作業選択.事前調整作業タブ(5)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMJizen()

            'タブ隠し(TAB3)
            Me.lblHidden3.Width = 885
            Me.lblHidden3.Height = 55


            '--------------------------------------------------
            ' パス設定
            '--------------------------------------------------
            '事前作業での出力リスト格納パス
            Dim tmp_defjizenlistdirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_LIST_NAME)
            Me.txtJizenListPath.Text = tmp_defjizenlistdirpath


            '--------------------------------------------------
            ' 中間ファイル関連設定
            '--------------------------------------------------
            '中間ファイルパス
            Dim tmp_defmiddirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_MID_NAME)
            Select Case CNVNO
                Case ConvertTypes._汎用
                    tmp_defmiddirpath = EtcMethod.Set_Path(tmp_defmiddirpath, FILE_HMIDD_NAME)
                Case Else

            End Select
            Me.txtMidDirPath.Text = tmp_defmiddirpath
            Me.txtMidDirPath.BackColor = Color.White
            Me.txtMidDirPath.ReadOnly = True
            Me.txtMidDirPath.ForeColor = Color.Gray
            Me.grpMiddleFile.Visible = False
            If CNVNO = ConvertTypes._汎用 Or Dev_CVFlg Then
                Me.grpMiddleFile.Visible = True
            Else
                ''※位置調整
                'Me.lblHidden2.Location = New Point(6, 185)
                'Me.lblLine2.Location = New Point(31, 225)
                'Me.tabCtrlCVItem.Location = New Point(30, 200)
            End If


            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            'リスト格納先テキストボックス設定
            Me.txtJizenListPath.BackColor = Color.White
            Me.txtJizenListPath.ReadOnly = True
            Me.txtJizenListPath.ForeColor = Color.Gray
            Me.lblCntJizenKai.Text = "なし"
            Me.lblCntJizenAzu.Text = "なし"
            Me.lblCntJizenMinus.Text = "なし"
            Me.lblCntJizenBkhoureiNotDef.Text = "なし"
            Me.lblCntJizenBkhoureiNotCV.Text = "なし"
            Me.lblCntJizenHasseiOw.Text = "なし"
            Me.lblCntJizenHasseiHen.Text = "なし"
            Me.lblCntJizenYanuso.Text = "なし"
            Me.lblCntJizenSzenKyshutan.Text = "なし"
            Me.lblCntJizenHyKagi.Text = "なし"            '20160725 事前作業へ鍵情報の追加 add
            Me.lblCntJizenKyKagi.Text = "なし"            '20160725 事前作業へ鍵情報の追加 add
            Me.lblCntJizenNonJisyaKoza.Text = "なし"      '20160805 改善対応 add
            Me.lblCntJizenSimeSokin.Text = "なし"         '20160829 事前作業_送金予定日リスト出力機能を追加 add
            Me.lblCntJizenKozameigikana.Text = "なし"     '20160829 口座名義カナチェック機能の追加 add
    
            Select Case CNVNO
                Case ConvertTypes._汎用
                    Me.pnlJizenListPath.Visible = False
                    '20161018 UI最終確認での修正 -add sta
                    Me.Label355.Visible = False
                    Me.lblJizenDescription1.Text = "移行元データとなる中間ファイルを事前に作成する必要があります。" & vbCrLf & _
                                                   "(完了したらチェックボックスをONにして下さい。ONにした場合のみデータコンバートを行うことができます)"
                    'Me.lblCautionDescription.Text = ""
                    '20161018 UI最終確認での修正 -add end
                Case Else
                    Me.pnlJizenListPath.Visible = True
            End Select
           
        End Sub

        ''' <summary>
        ''' メイン.作業選択.事後調整作業タブ(6)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMJigo()

            'タブ隠し(TAB4)
            Me.lblHidden4.Width = 883
            Me.lblHidden4.Height = 26


            '--------------------------------------------------
            ' パス設定
            '--------------------------------------------------
            '事前作業での出力リスト格納パス
            Dim tmp_defjigolistdirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_LIST_NAME)
            Me.txtJigoListPath.Text = tmp_defjigolistdirpath


            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            'リスト格納先テキストボックス設定
            Me.txtJigoListPath.BackColor = Color.White
            Me.txtJigoListPath.ReadOnly = True
            Me.txtJigoListPath.ForeColor = Color.Gray
            
            '出力ファイル格納先(20160714 現時点では未使用の為、非表示とする)
            Me.pnlJigoListPath.Visible = False

            Me.lblJizenPageNum2.Visible = False
            Me.pnlHJigoCmtSyudo.Visible = False
            Me.pnlKJigoCmtSyudo.Visible = False
            Me.pnlJigoCmtSyusi.Visible = False
            Me.pnlJigoCmtCsvSyuturyoku.Visible = False
            Me.pnlJigoCmtSyosiki.Visible = False
            Me.pnlJigoCmtSqKotiku.Visible = False
            Me.pnlJigoCmtNkNyuryoku.Visible = False
            Me.pnlJigoCmtSoKotiku.Visible = False
            Me.grpJigoCmtTyuui.Visible = False

            Select Case CNVNO
                Case ConvertTypes._汎用
                    Me.pnlHJigoCmtSyudo.Visible = True
                    Me.pnlJigoCmtSqKotiku.Visible = True
                    Me.pnlJigoCmtNkNyuryoku.Visible = True
                    Me.pnlJigoCmtSoKotiku.Visible = True
                    '----- 導入時文言移動 ----- sta
                    Me.grpJigoUserSagyo.Visible = True
                    Me.grpJigoDonyuji.Visible = True
                    Me.grpJigoUserSagyo.Size = New Size(405, 342)
                    Me.grpJigoUserSagyo.Location = New Point(6, 15)
                    Call Chg_CtrlInTabLocation(tabPageJigo1, tabPageJigo2, grpJigoUserSagyo)
                    Me.grpJigoDonyuji.Size = New Size(Me.grpJigoUserSagyo.Size)
                    Me.grpJigoDonyuji.Location = New Point(436, 15)
                    '----- 導入時文言移動 ----- end
                Case Else
                    Me.lblJizenPageNum2.Visible = True
                    Me.pnlKJigoCmtSyudo.Visible = True
                    Me.pnlJigoCmtSyusi.Visible = True
                    Me.pnlJigoCmtCsvSyuturyoku.Visible = True
                    Me.pnlJigoCmtSyosiki.Visible = True
                    Me.pnlJigoCmtSqKotiku.Visible = True
                    Me.pnlJigoCmtNkNyuryoku.Visible = True
                    Me.pnlJigoCmtSoKotiku.Visible = True
                    Me.grpJigoCmtTyuui.Visible = True
            End Select

        End Sub

        ''' <summary>
        ''' メイン.作業選択.補助機能タブ(7)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMHojyo()

            'タブ隠し(TAB5)
            Me.lblHidden5.Width = 883       '20160711 タブの表示処理の追加 -chg lblHidden4 → lblHidden5
            Me.lblHidden5.Height = 26       '20160711 タブの表示処理の追加 -chg lblHidden4 → lblHidden5


            '--------------------------------------------------
            ' パス設定
            '--------------------------------------------------
            '事前作業での出力リスト格納パス
            Dim tmp_defkensyolistdirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_LIST_NAME)
            Me.txtKensyoListPath.Text = tmp_defkensyolistdirpath

            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            'リスト格納先テキストボックス設定
            Me.txtKensyoListPath.BackColor = Color.White
            Me.txtKensyoListPath.ReadOnly = True
            Me.txtKensyoListPath.ForeColor = Color.Gray

            Me.lblRelRenameClaim.Text = "なし"
            Me.lblRelRenameSzen.Text = "なし"
            Me.lblCntKiOpOw.Text = "なし"
            Me.lblCntKiOpKys.Text = "なし"
            Me.lblCntBunkatumisyu.Text = "なし"
            Me.lblCntKojyosh.Text = "なし"
            Me.lblCntSzenKysSorit.Text = "なし"
            Me.lblCntRelationSet.Text = "なし"
           
            '20160720 リネーム機能の修正 初期状態は非活性にする -add sta
            Me.btnRNRelRename.Enabled = False
            Me.Label242.Visible = False
            '20160720 リネーム機能の修正 初期状態は非活性にする -add end
        End Sub

        ''' <summary>
        ''' (仮)メイン.必須項目一括設定タブ(14)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMIkkatu()

        End Sub
        ''' <summary>
        ''' (仮)メイン.事前調整作業タブ(汎用コンバートキット用)(15)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMHanyoJizen()

        End Sub

        ''' <summary>
        ''' コンバーター.はじめにタブ(8)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageDCHajimeni()

            '20160928 同意内容修正 -chg sta
            'Me.btnDoui.Enabled = True
            ''20160725 汎用CVK対応 -add sta
            'pnlDcFstCmt09.Visible = True
            'lblHDatacvHajimeniLabel.Visible = False
            'lblKDatacvHajimeniLabel.Visible = False
            'Select Case CNVNO
            '    Case ConvertTypes._汎用
            '        lblHDatacvHajimeniLabel.Visible = True
            '        '20160921 同意内容修正 -chg sta
            '        'pnlDcFstCmt09.Location = New Size(pnlDcFstCmt08.Location)
            '        'pnlDcFstCmt08.Visible = False
            '        pnlDcFstCmtK01.Visible = False
            '        pnlDcFstCmtK02.Visible = False
            '        '20160921 同意内容修正 -chg end
            '    Case Else
            '        lblKDatacvHajimeniLabel.Visible = True
            '        '20160921 同意内容修正 -chg sta
            '        pnlDcFstCmt07.Visible = False
            '        pnlDcFstCmt10.Visible = False
            '        '20160921 同意内容修正 -chg end
            'End Select
            ''20160725 汎用CVK対応 -add end
            Me.btnDoui.Enabled = True
            lblHDatacvHajimeniLabel.Visible = False
            lblKDatacvHajimeniLabel.Visible = False
            Select Case CNVNO
                Case ConvertTypes._汎用
                    lblHDatacvHajimeniLabel.Visible = True
                    pnlDcFstCmtK02.Visible = False
                    pnlDcFstCmtK01.Visible = False
                    pnlDcFstCmtK99.Visible = False
                    pnlDcFstCmtH99.Location = New Size(455, 300)
                Case Else
                    lblKDatacvHajimeniLabel.Visible = True
                    pnlDcFstCmtH02.Visible = False
                    pnlDcFstCmtH01.Visible = False
                    pnlDcFstCmtH99.Visible = False
                    pnlDcFstCmtK02.Location = New Size(pnlDcFstCmtH02.Location)
                    pnlDcFstCmtK01.Location = New Size(pnlDcFstCmtH01.Location)
                    pnlDcFstCmtK99.Location = New Size(455, 300)
            End Select
            '20160928 同意内容修正 -chg end

        End Sub

        ''' <summary>
        ''' コンバーター. 移行項目選択タブ(9)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageDCSelect()

            '--------------------------------------------------
            ' 中間ファイル関連設定
            '--------------------------------------------------
            '中間ファイルログパス
            Dim tmp_pathstr As String = ""
            tmp_pathstr = EtcMethod.Set_Path(dcv_exedir, DIR_MID_NAME)
            Dim tmp_defmidlogdirpath As String = EtcMethod.Set_Path(tmp_pathstr, DIR_MIDLOG_NAME)
            Me.txtMidDirLogPath.Text = tmp_defmidlogdirpath
            Me.txtMidDirPath2.Text = Me.txtMidDirPath.Text
            '中間ファイル関連
            Me.txtMidDirPath2.BackColor = Color.White
            Me.txtMidDirLogPath.BackColor = Color.White
            Me.txtMidDirPath2.ReadOnly = True
            Me.txtMidDirLogPath.ReadOnly = True
            Me.txtMidDirPath2.ForeColor = Color.Gray
            Me.txtMidDirLogPath.ForeColor = Color.Gray
            Me.grpMiddleFile.Visible = False
            If CNVNO = ConvertTypes._汎用 Or Dev_CVFlg Then
                Me.grpMiddleFile.Visible = True
            Else
                ''※位置調整
                'Me.lblHidden2.Location = New Point(6, 185)
                'Me.lblLine2.Location = New Point(31, 225)
                'Me.tabCtrlCVItem.Location = New Point(30, 200)
            End If

            '20160912 既存用→汎用用中間ファイルへのコピー処理 -add sta
            If CNVNO = ConvertTypes._汎用 And Dev_CVFlg Then
                Me.txtExistMidToBaseMid.Text = Path.GetDirectoryName(Me.txtMidDirPath.Text)
            End If
            '20160912 既存用→汎用用中間ファイルへのコピー処理 -add end
            
            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            '全移行対象ブロック名取得(グループ・チェックボックス名取得)
            Call Me.Set_CVItemName()

            '全チェックボックスON(既存ユーザ用)
            Select Case CNVNO
                Case ConvertTypes._既存ユーザ用
                    Call Set_ChkboxOnOff(DC_SELTAB_K01, False, True)
                    Call Set_ChkboxOnOff(DC_SELTAB_K02, False, True)
                    Call Set_ChkboxOnOff(DC_SELTAB_K03, False, True)
                    Call Set_OptSelect()

                Case ConvertTypes._汎用
                    '20160801 タイトルマスタ統合処理 -chg sta
                    'Me.pnlKiMstBikotitle.Visible = False
                    Me.pnlKiMstTitle.Visible = False
                    '20160801 タイトルマスタ統合処理 -chg end
                    Me.pnlKiSyskanriBase.Visible = False
                    Me.grpKizonKagi.Visible = False
                    ''20160905 汎用コンバートの請求情報修正 -del
                    'Me.pnlKiSq.Visible = False
                    Me.pnlKiClaim.Visible = False
                    Me.pnlSzen.Visible = False
                    Me.pnlRendo.Visible = False
                    Me.Label359.Visible = False
                    Me.pnlKiOw.Location = New Size(Me.pnlKiJisya.Location)
                    Me.pnlKiJisya.Location = New Size(Me.pnlKiSyskanriBase.Location)
                    Me.pnlKiMstKasyoClaimrui.Visible = False                                        '20160905 汎用の場合の移行対象外チェックボックスを非表示にする -add
                    '20160725 汎用CVK対応 -add end
                    Me.pnlKiMstHendo.Visible = False    											'20160926 選定した移行項目をプログラムへ反映する修正 -add
                    '20161012 家主固定控除件数表示処理の追加 -add sta
                    Me.Label288.Text = "家主固定控除"
                    Me.Label290.Visible = False
                    Me.Label289.Visible = False
                    Me.lblKiSqMiBaseCnt.Visible = False
                    Me.Label160.Visible = False
                    Me.Label162.Visible = False
                    Me.lblKiSqAzBaseCnt.Visible = False
                    Me.Label308.Location = New Point(Me.Label290.Location)  '「( 参考件数 = 」
                    '20161017 家主固定控除の件数表示の修正 -chg sta
                    'Me.Label311.Location = New Point(Me.Label289.Location)  '「)」
                    'Me.lblKiSqOwKojoBaseCnt.Location = New Point(Me.lblKiSqMiBaseCnt.Location)
                    '契約情報のLocationを引用する
                    Me.Label311.Location = New Point(Me.Label285.Location)  '「)」
                    Me.lblKiSqOwKojoBaseCnt.Location = New Point(Me.lblKiKyBaseCnt.Location)
                    '20161017 家主固定控除の件数表示の修正 -chg end
                    Me.Label308.Visible = True
                    Me.Label311.Visible = True
                    Me.lblKiSqOwKojoBaseCnt.Visible = True
                    '20161012 家主固定控除件数表示処理の追加 -add end
                    '20161018 UI最終確認での修正 -add
                    Me.Label284.Text = "契約, 入居者"
                Case Else

            End Select

            '既存用→汎用用中間ファイルへコピーする処理(開発用) -add
            Me.grpExistMidToBaseMid.Visible = Dev_CVFlg

            'タブ隠し(TAB2)
            Me.lblHidden2.Width = 885
            Me.lblHidden2.Height = 28
            Me.lblKiMstBusCnt.Text = "なし"
            Me.lblKiMstSchoolCnt.Text = "なし"
            Me.lblKiMstAreaCnt.Text = "なし"
            Me.lblKiMstHokenruiCnt.Text = "なし"
            Me.lblKiMstTokuyakuCnt.Text = "なし"
            Me.lblKiMstKasyoClaimruiCnt.Text = "なし"
            Me.lblKiMstHendoCnt.Text = "なし"
            Me.lblKiMstKagititleCnt.Text = "なし"
            Me.lblKiMstBikotitleCnt.Text = "なし"
            Me.lblKiSyskanriBaseCnt.Text = "なし"
            Me.lblKiJisyaBaseCnt.Text = "なし"
            Me.lblKiOwBaseCnt.Text = "なし"
            Me.lblKiGyCyukaiBaseCnt.Text = "なし"
            Me.lblKiGySyuzenBaseCnt.Text = "なし"
            Me.lblKiGyLifelineBaseCnt.Text = "なし"
            Me.lblKiGyHokenBaseCnt.Text = "なし"
            Me.lblKiGyYatinhosyoBaseCnt.Text = "なし"
            Me.lblKiGySisetuBaseCnt.Text = "なし"
            Me.lblKiGySekoBaseCnt.Text = "なし"
            Me.lblKiBkBaseCnt.Text = "なし"
            Me.lblKiHyBaseCnt.Text = "なし"
            Me.lblKiHySetubiCnt.Text = "なし"
            Me.lblKiKysBaseCnt.Text = "なし"
            Me.lblKiKyBaseCnt.Text = "なし"
            Me.lblKiSqMiBaseCnt.Text = "なし"
            Me.lblKiSqAzBaseCnt.Text = "なし"
            Me.lblKiClaimBaseCnt.Text = "なし"
            Me.lblKiSzenBaseCnt.Text = "なし"
            Me.lblKiRendoBaseCnt.Text = "なし"    
            '20161108_2 レビュー結果戻り修正時に気付いた箇所の修正 -add
            '家主請求控除の件数表示文字列初期化処理を追加
            Me.lblKiSqOwKojoBaseCnt.Text = "なし"

            '--------------------------------------------------
            ' グループボックス内処理順番設定
            ' (ベースタブ→グループ→チェックボックスの順序を指定)
            '--------------------------------------------------
            '各基本情報(自社,家主,契約者)
            Call Me.Set_ControlIndex(tabPageBase130)
            'クレーム修繕情報
            Call Me.Set_ControlIndex(tabPageBase190)


            '--------------------------------------------------
            ' チェックボックス処理順番設定
            '--------------------------------------------------
            '条件(開発用、コンバートタイプ)によるチェックボックスの表示設定
            Call Me.Set_ChkBoxVisible()

            'マスタ系
            Call Me.Set_ControlIndex(grpMst)
            Call Me.Set_ChkBoxValue(grpMst, False)

            '業者情報
            Call Me.Set_ControlIndex(grpGy)
            Call Me.Set_ChkBoxValue(grpGy, False)

            '自社情報
            Call Me.Set_ControlIndex(grpJisya)
            Call Me.Set_ChkBoxValue(grpJisya, False)

            '家主情報
            Call Me.Set_ControlIndex(grpOw)
            Call Me.Set_ChkBoxValue(grpOw, False)

            '契約者情報
            Call Me.Set_ControlIndex(grpKys)
            Call Me.Set_ChkBoxValue(grpKys, False)

            '物件情報
            Call Me.Set_ControlIndex(grpBk)
            Call Me.Set_ChkBoxValue(grpBk, False)

            '部屋情報
            Call Me.Set_ControlIndex(grpHy)
            Call Me.Set_ChkBoxValue(grpHy, False)

            '送金ルール
            Call Me.Set_ControlIndex(grpSorule)
            Call Me.Set_ChkBoxValue(grpSorule, False)

            '契約情報
            Call Me.Set_ControlIndex(grpKy)
            Call Me.Set_ChkBoxValue(grpKy, False)

            '請求情報
            Call Me.Set_ControlIndex(grpSq)
            Call Me.Set_ChkBoxValue(grpSq, False)

            'クレーム情報
            Call Me.Set_ControlIndex(grpClaim)
            Call Me.Set_ChkBoxValue(grpClaim, False)

            '修繕情報
            Call Me.Set_ControlIndex(grpSzen)
            Call Me.Set_ChkBoxValue(grpSzen, False)

            '初期設定
            Call Me.Set_ControlIndex(grpSyskanri)
            Call Me.Set_ChkBoxValue(grpSyskanri, False)

            '物件データ連動情報  
            Call Me.Set_ControlIndex(grpRendo)
            Call Me.Set_ChkBoxValue(grpRendo, False)

            '(開発用) 子項目チェック制御
            If chkChildItemControl.Checked Then
                Call Me.Set_ChkBoxCondition(False)
            End If

            '移行対象外項目をオブジェクトへ格納(既存、汎用で処理を分岐)
            Select Case CNVNO
                Case ConvertTypes._既存ユーザ用
                    Call Me.Set_BaseCVItemToList()
                Case ConvertTypes._汎用
                    Call Me.Set_ExistCVItemToList()
            End Select

            '連動ID入力テキストボックスの制御
            Me.txtRendoID.MaxLength = 100

            '20161012 家主固定控除件数表示処理の追加 -del sta
            ''20160905 汎用コンバートの請求情報修正 -add
            'Me.Label288.Text = "導入時未収金, 預り金"
            '20161012 家主固定控除件数表示処理の追加 -del end

        End Sub

        ''' <summary>
        ''' コンバーター.処理実行タブ(10)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageDCJikko()

            '-----------------------------
            'プログレスバー初期設定
            '-----------------------------
            '20161017 進捗表示ラベル初期表示修正 -chg sta
            'Me.lblTotalSituation.Text = "コンバート処理準備中 ..."
            Me.lblTotalSituation.Text = SITUATION_CV_BFRUN
            '20161017 進捗表示ラベル初期表示修正 -chg end
            Me.lblCVItem.Text = ""
            Me.txtTotalSituation.ReadOnly = True
            Me.txtPartialSituation.ReadOnly = True
            Me.txtTotalSituation.BackColor = Color.White
            Me.txtPartialSituation.BackColor = Color.White
            Me.txtTotalSituation.ScrollBars = ScrollBars.Both
            Me.txtTotalSituation.WordWrap = False
            Me.txtPartialSituation.ScrollBars = ScrollBars.Both
            Me.txtPartialSituation.WordWrap = False

        End Sub

        ''' <summary>
        ''' コンバーター.終了タブ(11)(12)(13)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageDCEnd()

            Me.lblHDatacvEndLabel.Visible = False
            Me.lblKDatacvEndLabel.Visible = False
            Select Case CNVNO
                Case ConvertTypes._汎用
                    Me.lblHDatacvEndLabel.Visible = True
                    Me.lblHDatacvEndLabel.Location = New Size(Me.lblKDatacvEndLabel.Location)
                Case Else
                    Me.lblKDatacvEndLabel.Visible = True
            End Select

            '--------------------------------------------------
            ' 終了画面の不要コントロール制御
            '--------------------------------------------------
            'キャンセル画面
            Me.GroupBox22.Visible = False
            Me.GroupBox21.Visible = False

            'エラー画面
            Me.GroupBox20.Visible = False
            Me.GroupBox23.Visible = False
            Me.GroupBox19.Visible = False

        End Sub

        ''' <summary>
        ''' (隠)メイン.開発用タブ(0)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMDev()

            '--------------------------------------------------
            ' 紐付設定ファイル関連パス
            '--------------------------------------------------
            Dim tmp_pathstr As String
            tmp_pathstr = EtcMethod.Set_Path(dcv_exedir, DIR_RELEXEDIR_NAME)
            Dim tmp_defreldirpath As String = EtcMethod.Set_Path(tmp_pathstr, DIR_REL_NAME)
            Me.txtRelationDirPath.Text = tmp_defreldirpath


            '--------------------------------------------------
            ' 各チェックボックス初期設定
            '--------------------------------------------------
            Me.chkCVStart.Checked = True
            Me.chkLogTblDrop.Checked = True
            Me.chkV7ViewDrop.Checked = True
            Me.chkMiddleFile.Checked = True
            Me.chkOverWrite.Checked = False
            Me.chkRelTblDrop.Checked = True
            Me.chkMidNotStop.Checked = True
            Me.chkRelation.Checked = True

            '(汎用)
            If CNVNO = ConvertTypes._汎用 Then
                Me.chkMiddleFile.Checked = False
            End If

            '(テスト用)                                                            
            Me.chkV7ViewDrop.Checked = False
            Me.chkLogTblDrop.Checked = False
            Me.chkRelTblDrop.Checked = False

        End Sub

        ''' <summary>
        ''' メイン.初期表示フォーム設定(共通画面)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormtabPageMain()

            '--------------------------------------------------
            ' ラベル(タイトル)設定
            '--------------------------------------------------
            Me.Text = ""
            Select Case CNVNO
                Case ConvertTypes._既存ユーザ用
                    Me.Text = Me.Text & CV_KIZON_TITLE
                    Me.Text = Me.Text & ":" & CV_FROM_NAME & " → " & CV_TO_NAME
                Case ConvertTypes._他社システム用
                    Me.Text = Me.Text & CV_TASYA_TITLE
                    Me.Text = Me.Text & ":" & CV_FROM_SYSTEM & " → " & CV_TO_NAME
                Case Else       '汎用コンバートキット
                    Me.Text = Me.Text & CV_HANYO_TITLE
                    Me.Text = Me.Text & ":" & CV_FROM_MIDDLE & " → " & CV_TO_NAME
            End Select


            '--------------------------------------------------
            ' コントロール初期設定
            '--------------------------------------------------
            'サイズ設定
            '----- 要修正 ----- sta
            '※サイズ可変対応
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            '----- 要修正 ----- end

            'フォームコントロールボックス非可視設定
            Me.ControlBox = False

            'タブ隠し(TAB1)                                                  
            Me.lblHidden1.Width = 950
            Me.lblHidden1.Height = 55

            'ボタンの活性制御
            Call Chg_BtnKariStatus("mstart")

            'タブ非表示設定("1.開始タブ"以外は非表示)
            Dim tmpcntii As Integer = 0
            tabPageManager = New Njc.Common.TabPageManager(tabCtrlMain)
            Dim tmpcntjj As Integer = tabCtrlMain.TabPages.Count - 1
            For cntii As Integer = tmpcntii To tmpcntjj
                tabPageManager.ChangeTabPageVisible(cntii, False)
            Next
            tabPageManager.ChangeTabPageVisible(1, True)

            '(開発用) 開発用タブON/OFFボタン表示設定
            Me.btnDevTabChange.Visible = Dev_CVFlg

            'オプション選択全チェックON
            Call Set_ChkboxOnOff(M_SYOKITAB_OPS, False, 1)
            '中間用プログレスバー非表示(必要時のみ表示)
            Me.pnlPrgChk.Visible = False


            '--------------------------------------------------
            ' タブオブジェクト生成
            '-------------------------------------------------- 
            '共通設定値
            tmpcntii = 1
            '作業選択.事前作業
            tabPageManagerJizen = New Njc.Common.TabPageManager(tabCtrlJizen)
            tmpcntjj = tabCtrlJizen.TabPages.Count - 1
            For cntii As Integer = tmpcntii To tmpcntjj
                tabPageManagerJizen.ChangeTabPageVisible(cntii, False)
            Next
            '作業選択.コンバーター.選択項目
            tabPageManagerCvitem = New Njc.Common.TabPageManager(tabCtrlCVItem)
            tmpcntjj = tabCtrlCVItem.TabPages.Count - 1
            For cntii As Integer = tmpcntii To tmpcntjj
                tabPageManagerCvitem.ChangeTabPageVisible(cntii, False)
            Next
            '作業選択.事後作業
            tabPageManagerJigo = New Njc.Common.TabPageManager(tabCtrlJigo)
            tmpcntjj = tabCtrlJigo.TabPages.Count - 1
            For cntii As Integer = tmpcntii To tmpcntjj
                tabPageManagerJigo.ChangeTabPageVisible(cntii, False)
            Next
            '作業選択.補助機能
            tabPageManagerHojyo = New Njc.Common.TabPageManager(tabCtrlHojyo)
            tmpcntjj = tabCtrlHojyo.TabPages.Count - 1
            For cntii As Integer = tmpcntii To tmpcntjj
                tabPageManagerHojyo.ChangeTabPageVisible(cntii, False)
            Next


            '--------------------------------------------------
            ' 画面左進捗・セットフォーカス
            '--------------------------------------------------
            Call Chg_LeftProgress("mstart")
            Call Set_Forcus("next")

        End Sub

        ''' <summary>
        ''' 画面左進捗表示設定
        ''' </summary>
        ''' <param name="chgstr">状態文字列</param>
        ''' <param name="chgtype">状態タイプ…0.メイン画面, 1.作業選択画面 [op]</param>
        ''' <remarks>
        ''' ・デフォルトタイプ値=0.メイン画面
        ''' </remarks>
        Private Sub Chg_LeftProgress(ByVal chgstr As String, Optional chgtype As Integer = 0)

            Select Case chgtype

                Case 1
                    '--------------------------------------------------
                    ' コンバーター画面左進捗表示
                    '--------------------------------------------------
                    '画面左進捗表示
                    pnlMenuMain.Visible = False                 'メイン画面左進捗
                    pnlMenuDatacv.Visible = True                'コンバーター画面左進捗
                    '背景初期化
                    pnlPrgDCHajimeni.BackColor = Color.LightBlue
                    pnlPrgDCSelect.BackColor = Color.LightBlue
                    pnlPrgDCJikko.BackColor = Color.LightBlue
                    pnlPrgDCEnd.BackColor = Color.LightBlue
                    pnlPrgDCFileWrite.BackColor = Color.LightBlue
                    pnlPrgDCRelation.BackColor = Color.LightBlue
                    pnlPrgDCConv.BackColor = Color.LightBlue
                    'アイコン(青)表示
                    picIcoDCHajimeni1.Visible = True
                    picIcoDCSelect1.Visible = True
                    picIcoDCJikko1.Visible = True
                    picIcoDCEnd1.Visible = True
                    picIcoDCFileWrite1.Visible = True
                    picIcoDCRelation1.Visible = True
                    picIcoDCConv1.Visible = True
                    'アイコン(赤)表示
                    picIcoDCHajimeni2.Visible = False
                    picIcoDCSelect2.Visible = False
                    picIcoDCJikko2.Visible = False
                    picIcoDCEnd2.Visible = False
                    picIcoDCFileWrite2.Visible = False
                    picIcoDCRelation2.Visible = False
                    picIcoDCConv2.Visible = False
                    '矢印(青)表示
                    picArwDCSelect1.Visible = True
                    picArwDCJikko1.Visible = True
                    picArwDCEnd1.Visible = True
                    '矢印(赤)表示
                    picArwDCSelect2.Visible = False
                    picArwDCJikko2.Visible = False
                    picArwDCEnd2.Visible = False
                    '矢印(赤)前面表示
                    picArwDCSelect2.BringToFront()
                    picArwDCJikko2.BringToFront()
                    picArwDCEnd2.BringToFront()
                    '背景前面表示
                    pnlPrgDCHajimeni.BringToFront()
                    pnlPrgDCSelect.BringToFront()
                    pnlPrgDCJikko.BringToFront()
                    pnlPrgDCEnd.BringToFront()
                    pnlPrgDCFileWrite.BringToFront()
                    pnlPrgDCRelation.BringToFront()
                    pnlPrgDCConv.BringToFront()
                    'カレント箇所状態変更
                    Select Case chgstr
                        Case "dchajimeni"
                            pnlPrgDCHajimeni.BackColor = Color.LightPink
                            picIcoDCHajimeni2.Visible = True
                        Case "dcselect1", "dcselect2"
                            pnlPrgDCSelect.BackColor = Color.LightPink
                            picIcoDCSelect2.Visible = True
                            picArwDCSelect2.Visible = True
                        Case "dcjikko1", "dcjikko2"
                            pnlPrgDCJikko.BackColor = Color.LightPink
                            picIcoDCJikko2.Visible = True
                            picArwDCJikko2.Visible = True
                        Case "dcend"
                            pnlPrgDCEnd.BackColor = Color.LightPink
                            picIcoDCEnd2.Visible = True
                            picArwDCEnd2.Visible = True
                        Case "dcfilewrite"
                            pnlPrgDCJikko.BackColor = Color.LightPink
                            picIcoDCJikko2.Visible = True
                            picArwDCJikko2.Visible = True
                            pnlPrgDCFileWrite.BackColor = Color.LightPink
                            picIcoDCFileWrite2.Visible = True
                            picArwDCJikko2.Visible = True
                        Case "dcrelation"
                            pnlPrgDCJikko.BackColor = Color.LightPink
                            picIcoDCJikko2.Visible = True
                            picArwDCJikko2.Visible = True
                            pnlPrgDCRelation.BackColor = Color.LightPink
                            picIcoDCRelation2.Visible = True
                            picArwDCJikko2.Visible = True
                        Case "dcconv"
                            pnlPrgDCJikko.BackColor = Color.LightPink
                            picIcoDCJikko2.Visible = True
                            picArwDCJikko2.Visible = True
                            pnlPrgDCConv.BackColor = Color.LightPink
                            picIcoDCConv2.Visible = True
                            picArwDCJikko2.Visible = True
                    End Select

                Case Else
                    '--------------------------------------------------
                    ' メイン画面左進捗表示
                    '--------------------------------------------------
                    '画面左進捗表示
                    pnlMenuMain.Visible = True                      'メイン画面左進捗
                    pnlMenuDatacv.Visible = False                   'コンバーター画面左進捗
                    '背景初期化
                    pnlPrgMStart.BackColor = Color.LightBlue
                    pnlPrgMSession.BackColor = Color.LightBlue
                    pnlPrgMSyoki.BackColor = Color.LightBlue
                    pnlPrgMMenu.BackColor = Color.LightBlue
                    pnlPrgMEnd.BackColor = Color.LightBlue
                    'アイコン(青)表示
                    picIcoMStart1.Visible = True
                    picIcoMSession1.Visible = True
                    picIcoMSyoki1.Visible = True
                    picIcoMMenu1.Visible = True
                    picIcoMEnd1.Visible = True
                    'アイコン(赤)表示
                    picIcoMStart2.Visible = False
                    picIcoMSession2.Visible = False
                    picIcoMSyoki2.Visible = False
                    picIcoMMenu2.Visible = False
                    picIcoMEnd2.Visible = False
                    '矢印(青)表示
                    picArwMSession1.Visible = True
                    picArwMSyoki1.Visible = True
                    picArwMMenu1.Visible = True
                    picArwMEnd1.Visible = True
                    '矢印(赤)表示
                    picArwMSession2.Visible = False
                    picArwMSyoki2.Visible = False
                    picArwMMenu2.Visible = False
                    picArwMEnd2.Visible = False
                    '矢印(赤)前面表示
                    picArwMSession2.BringToFront()
                    picArwMSyoki2.BringToFront()
                    picArwMMenu2.BringToFront()
                    picArwMEnd2.BringToFront()
                    '背景前面表示
                    pnlPrgMStart.BringToFront()
                    pnlPrgMSession.BringToFront()
                    pnlPrgMSyoki.BringToFront()
                    pnlPrgMMenu.BringToFront()
                    pnlPrgMEnd.BringToFront()
                    '背景初期化
                    pnlPrgMStart.BackColor = Color.LightBlue
                    pnlPrgMSession.BackColor = Color.LightBlue
                    pnlPrgMSyoki.BackColor = Color.LightBlue
                    pnlPrgMMenu.BackColor = Color.LightBlue
                    pnlPrgMEnd.BackColor = Color.LightBlue
                    'カレント箇所状態変更
                    Select Case chgstr
                        Case "mstart"
                            pnlPrgMStart.BackColor = Color.LightPink
                            picIcoMStart2.Visible = True
                        Case "msession"
                            pnlPrgMSession.BackColor = Color.LightPink
                            picIcoMSession2.Visible = True
                            picArwMSession2.Visible = True
                        Case "msyoki"
                            pnlPrgMSyoki.BackColor = Color.LightPink
                            picIcoMSyoki2.Visible = True
                            picArwMSyoki2.Visible = True
                        Case "mmenu"
                            pnlPrgMMenu.BackColor = Color.LightPink
                            picIcoMMenu2.Visible = True
                            picArwMMenu2.Visible = True
                        Case "mend"
                            pnlPrgMEnd.BackColor = Color.LightPink
                            picIcoMEnd2.Visible = True
                            picArwMEnd2.Visible = True
                        Case Else

                    End Select

            End Select

        End Sub

        ''' <summary>
        ''' ハッシュテーブル初期化
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Ini_HashTable()

            Hash_TblName_JpToAlpha = New Hashtable
            Hash_TblName_AlphaToJp = New Hashtable
            Hash_FldName_JpToAlpha = New Hashtable
            Hash_FldName_AlphaToJp = New Hashtable
            Hash_FiledTypeAlpha = New Hashtable
            Hash_FiledSizeAlpha = New Hashtable
            Hash_FiledTypeJp = New Hashtable
            Hash_FiledSizeJp = New Hashtable
            Hash_DefaultValue = New Hashtable
            Hash_FiledKey = New Hashtable
            Hash_Min_Code = New Hashtable
            Hash_Max_Code = New Hashtable
            Hash_Mst_ReferenceAlpha = New Hashtable
            Hash_MidKey_FieldAlpha = New Hashtable
            Hash_MidKey_FieldJp = New Hashtable
            Hash_FldName_JpToAlpha_Mid = New Hashtable

            Hash_Kinyu = New Hashtable
            Hash_KinyuTen = New Hashtable
            Hash_KagiTitleKyoyo = New Hashtable
            Hash_KagiTitleSenyo = New Hashtable
            Hash_KagiTitle_V7to10 = New Hashtable
            Hash_SetubiMid = New Hashtable
            Hash_SetubiMst = New Hashtable
            Hash_SetubiMst_UserMake = New Hashtable 		'20160829 設備の新規挿入処理を追加

            Hash_ClaimBruiKasyo = New Hashtable
            Hash_ClaimBruiClaim = New Hashtable

            '20160711 オブジェクトインスタンス化の修正 -add sta
            '紐付データ取得用のハッシュテーブルも初期化する
            Hash_Rel_Bkrui = New Hashtable
            Hash_Rel_Kozo = New Hashtable
            Hash_Rel_Hyrui = New Hashtable
            Hash_Rel_Toritaiyo = New Hashtable
            Hash_Rel_Kozasyubetu = New Hashtable
            Hash_Rel_Nkinkomk = New Hashtable
            Hash_Rel_NkinkomkZksei = New Hashtable
            Hash_Rel_Nkinkbn = New Hashtable
            Hash_Rel_Hendometer = New Hashtable
            Hash_Rel_Setubi = New Hashtable
            Hash_Rel_Kyrui = New Hashtable
            Hash_Rel_Kyrui_Teikisyakuyakbn = New Hashtable
            Hash_Rel_Tosiyotokbn = New Hashtable
            Hash_Rel_Tosiyotono = New Hashtable
            Hash_Rel_JisyaKoza = New Hashtable
            Hash_Rel_FBInfo_KozaFurikaeFmt = New Hashtable
            Hash_Rel_FBInfo_FuriIraiFmt = New Hashtable
            Hash_Rel_FBInfo_NsSettingFmt = New Hashtable
            Hash_Rel_FBInfo_NsSettingJisya = New Hashtable
            '20160711 オブジェクトインスタンス化の修正 -add end

        End Sub

        ''' <summary>
        ''' コンバート関連フォルダ自動生成処理 '20161018 フォルダ自動生成箇所の修正 -add
        ''' </summary>
        ''' <remarks></remarks>
        Private Function IniDirSet() As Boolean

            '--------------------------------------------------
            ' フォルダ自動生成処理
            '--------------------------------------------------
            Dim tmp_middirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_MID_NAME)             'コンバーター実行ファイル格納フォルダ
            Dim tmp_reldirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_RELEXEDIR_NAME)       '紐付ツール実行ファイル格納フォルダ
            Dim tmp_gazodirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_GAZOEXEDIR_NAME)     '画像CVツール実行ファイル格納フォルダ
            Dim list_makedir As New List(Of String)

            'コンバーターフォルダ同階層
            list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_MAINLOG_NAME))
            list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_INI_NAME))
            list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_LIST_NAME))
            '中間ファイルフォルダ
            If EtcMethod.Chk_DirExist(tmp_middirpath) Then
                list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, DIR_TEMP_CSV))
                list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, DIR_MIDLOG_NAME))
            End If
            '紐付ツールフォルダ
            If EtcMethod.Chk_DirExist(tmp_reldirpath) Then
                list_makedir.Add(EtcMethod.Set_Path(tmp_reldirpath, DIR_RELLOG_NAME))
            End If
            '画像CVツールフォルダ
            '20161108_2 レビュー結果戻り修正 -chg sta
            'If EtcMethod.Chk_DirExist(tmp_gazodirpath) Then                                         '20161108 レビュー結果：汎用CVKの場合でも作成する？
            '    list_makedir.Add(EtcMethod.Set_Path(tmp_gazodirpath, DIR_INI_NAME))
            'End If
            If CNVNO = ConvertTypes._既存ユーザ用 Then
                If EtcMethod.Chk_DirExist(tmp_gazodirpath) Then
                    list_makedir.Add(EtcMethod.Set_Path(tmp_gazodirpath, DIR_INI_NAME))
                End If
            End If
            '20161108_2 レビュー結果戻り修正 -chg end
            '自動生成処理
            '20161104 フォルダ自動作成処理の修正 -chg sta
            'For Each chkdir In list_makedir
            '    If System.IO.Directory.Exists(chkdir) = False Then
            '        System.IO.Directory.CreateDirectory(chkdir)
            '    End If
            'Next

            Try

                For Each chkdir In list_makedir
                    If System.IO.Directory.Exists(chkdir) = False Then
                        System.IO.Directory.CreateDirectory(chkdir)
                    End If
                Next

                'アクセスが拒否された場合
                'Catch ex As UnauthorizedAccessException
                '    MsgResult = MessageBox.Show("フォルダ作成時にアクセスが拒否されました。" & vbCrLf & _
                '                                "コンバーターのインストール先にアクセス権限を与えるか、インストール先を変更して再度セットアップを行って下さい。", _
                '                                "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '    Return False

                '無効なフォルダパスが指定された場合
            Catch ex As DirectoryNotFoundException
                MsgResult = MessageBox.Show("無効なフォルダパスが指定されました。" & vbCrLf & _
                                            "関連フォルダが削除された可能性があります。再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False

                'その他(アクセスが拒否された場合を含む)
            Catch ex As Exception
                MsgResult = MessageBox.Show("セットアップフォルダのアクセスが拒否されました。" & vbCrLf & _
                                            "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False

            End Try

            Return True
            '20161104 フォルダ自動作成処理の修正 -chg end
        End Function

#End Region

#Region "イベント処理"

        ''' <summary>
        ''' イベント処理：メイン画面ロード
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

            '--------------------------------------------------
            ' 共通初期設定
            '--------------------------------------------------
            'コマンドライン値取得・条件設定
            Call IniCmdline()
            '20161018 フォルダ自動生成箇所の修正 -add sta
            'フォルダ自動生成処理
            '20161104 フォルダ自動作成処理の修正 -chg sta
            'Call Me.IniDirSet()
            If Me.IniDirSet() = False Then
                Me.Close()
                Exit Sub
            End If
            '20161104 フォルダ自動作成処理の修正 -chg end
            '20161018 フォルダ自動生成箇所の修正 -add end
            '--------------------------------------------------
            ' メイン画面＋各タブの初期設定
            '--------------------------------------------------
            'メインタブ設定
            Me.IniFormtabPageMStart()               '1.開始タブ
            Me.IniFormtabPageMSession()             '2.DB接続タブ
            Me.IniFormtabPageMSyoki()               '3.初期設定タブ
            Me.IniFormtabPageMMenu()                '4.作業選択タブ

            'メイン.作業選択ボタン設定
            Me.IniFormtabPageMJizen()               '5.事前調整作業タブ                      
            Me.IniFormtabPageMJigo()                '6.事後調整作業タブ                     
            Me.IniFormtabPageMHojyo()               '7.補助機能タブ

            'コンバータータブ設定
            Me.IniFormtabPageDCHajimeni()           '8.はじめにタブ
            Me.IniFormtabPageDCSelect()             '9.対象項目選択タブ                      
            Me.IniFormtabPageDCJikko()              '10.処理実行タブ
            Me.IniFormtabPageDCEnd()                '11～13.終了タブ

            '開発用・仮用意タブ設定
            Me.IniFormtabPageMIkkatu()              '14.(仮)必須項目一括設定タブ
            Me.IniFormtabPageMHanyoJizen()          '15.(仮)事前調整作業タブ(汎用CVK用)
            Me.IniFormtabPageMDev()                 '0.開発用タブ

            'メイン画面設定
            Me.IniFormtabPageMain()

            '20160719 V7オプション選択による表示設定処理追加_不足分対応 -add sta
            '※画面起動時にチェックボックスイベントに入るため起動時は入らないようにするフラグ
            frmloadflg_kysq = False
            frmloadflg_clsz = False
            '20160719 V7オプション選択による表示設定処理追加_不足分対応 -add end

        End Sub

        ''' <summary>
        ''' イベント処理：作業選択.事前作業ボタンクリック
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnMenuJizen_Click(sender As Object, e As EventArgs) Handles btnMenuJizen.Click

            tabPageManager.ChangeTabPageVisible(5, True)
            tabPageManager.ChangeTabPageVisible(4, False)

            Select Case CNVNO
                Case ConvertTypes._汎用
                    tabPageManagerJizen.ChangeTabPageVisible(0, False)
                    tabPageManagerJizen.ChangeTabPageVisible(1, False)
                    tabPageManagerJizen.ChangeTabPageVisible(2, False)
                    tabPageManagerJizen.ChangeTabPageVisible(3, False)
                    tabPageManagerJizen.ChangeTabPageVisible(4, False)
                    tabPageManagerJizen.ChangeTabPageVisible(5, True)
                    Call Chg_BtnKariStatus("jizenh")
                    Set_Forcus("jizen_tyukan")
                Case Else
                    tabPageManagerJizen.ChangeTabPageVisible(0, True)
                    tabPageManagerJizen.ChangeTabPageVisible(1, False)
                    tabPageManagerJizen.ChangeTabPageVisible(2, False)
                    tabPageManagerJizen.ChangeTabPageVisible(3, False)
                    tabPageManagerJizen.ChangeTabPageVisible(4, False)
                    tabPageManagerJizen.ChangeTabPageVisible(5, False)
                    Me.pnlRefresh.Visible = True
                    Call Me.Set_List_Cnt_Jizen()
                    Call Chg_BtnKariStatus("jizen1")
                    Set_Forcus("jizen_kai")
            End Select

        End Sub

        ''' <summary>
        ''' イベント：作業選択.データコンバートボタンクリック
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnMenuDatacv_Click(sender As Object, e As EventArgs) Handles btnMenuDatacv.Click

            tabPageManager.ChangeTabPageVisible(8, True)
            tabPageManager.ChangeTabPageVisible(4, False)
            Call Chg_BtnKariStatus("dchajimeni")
            Call Chg_LeftProgress("dchajimeni", 1)
            Set_Forcus("dcdoui")

        End Sub

        ''' <summary>
        ''' イベント：作業選択.画像コンバートボタンクリック
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・画像コンバーター呼出
        ''' </remarks>
        Private Sub btnMenuGazocv_Click(sender As Object, e As EventArgs) Handles btnMenuGazocv.Click

            Dim rowcnt As Integer = 0

            '紐付ツールパス取得(Exe実行用)
            Dim tmp_strpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_GAZOEXEDIR_NAME)
            Dim gazocvexepath As String = EtcMethod.Set_Path(tmp_strpath, DIR_GAZOEXE_NAME)

            'V7接続情報
            Dim coninfov7 As String = Me.optV7Authent1.Checked & "," & _
                                      fstmodelv7.ServerName.Trim & "," & _
                                      fstmodelv7.InitialCatalog.Trim & "," & _
                                      fstmodelv7.User.Trim & "," & _
                                      fstmodelv7.Pass.Trim

            '10接続情報
            Dim coninfov10 As String = Me.optV10Authent1.Checked & "," & _
                                       fstmodelv10.ServerName.Trim & "," & _
                                       fstmodelv10.InitialCatalog.Trim & "," & _
                                       fstmodelv10.User.Trim & "," & _
                                       fstmodelv10.Pass.Trim & "," & _
                                       fstmodelv10.TimeOut

            '画像コンバーター呼出
            '(パラメータ：P1.起動方法選択 P2.コンバートタイプ P3.開発用機能表示 P4.ユーザ名 P5.V7接続情報文字列 P6.10接続情報文字列 P7.表示位置)
            '例：0 1 False CONVUSER True,PC-USER\SQLSVR2012,fk5dtsql,sa,passwordv7 True,PC-USER\SQLSVR2012,fk8db,sa,passwordv8,15 75,75
            Dim tmp_cvno As String = CNVNO                                                  'コンバートタイプ (0.汎用, 1.既存ユーザ用, 2.他社システム用)
            Dim tmp_devflg As String = Dev_CVFlg                                            'ツール起動時の開発用フラグ
            Dim tmp_username As String = Me.txtRecUser.Text                                 'ユーザ名(作業者名)
            Dim tmp_singlestart As String = "0"                                             '画像コンバーター呼出起動(固定：0.呼出起動)
            Dim tmp_gazocvloc As String = Me.Left.ToString & "," & Me.Top.ToString          '画像コンバーター表示位置指定
            Dim cmd_fk8db_cnv As String = tmp_singlestart & " " & tmp_cvno & " " & tmp_devflg & " " & _
                                          """" & tmp_username & """" & " " & _
                                          """" & coninfov7 & """" & " " & _
                                          """" & coninfov10 & """" & " " & _
                                          tmp_gazocvloc

            '接続情報と紐付ファイル出力先フォルダパスを渡して起動                       
            Dim obj_ExeStart As System.Diagnostics.Process = System.Diagnostics.Process.Start(gazocvexepath, cmd_fk8db_cnv)
            obj_ExeStart.WaitForExit()
         
        End Sub

        ''' <summary>
        ''' イベント：作業選択.事後作業ボタンクリック
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnMenuJigo_Click(sender As Object, e As EventArgs) Handles btnMenuJigo.Click

            tabPageManager.ChangeTabPageVisible(6, True)
            tabPageManager.ChangeTabPageVisible(4, False)

            Select Case CNVNO
                Case ConvertTypes._汎用
                    tabPageManagerJigo.ChangeTabPageVisible(1, True)
                    tabPageManagerJigo.ChangeTabPageVisible(0, False)
                    Call Chg_BtnKariStatus("jigoh")
                    Set_Forcus("next")
                Case Else
                    tabPageManagerJigo.ChangeTabPageVisible(0, True)
                    tabPageManagerJigo.ChangeTabPageVisible(1, False)
                    Call Chg_BtnKariStatus("jigo1")
                    Set_Forcus("next")
            End Select

        End Sub

        ''' <summary>
        ''' イベント：作業選択.補助機能ボタンクリック
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・検証用リスト出力
        ''' </remarks>
        Private Sub btnMenuHojyo_Click(sender As Object, e As EventArgs) Handles btnMenuHojyo.Click

            tabPageManager.ChangeTabPageVisible(7, True)
            tabPageManager.ChangeTabPageVisible(4, False)
            tabPageManagerHojyo.ChangeTabPageVisible(0, True)
            tabPageManagerHojyo.ChangeTabPageVisible(1, False)
            Me.pnlRefresh.Visible = True                       '件数再読込ボタン表示設定
            '20160707 検証作業リスト出力処理の追加 -add
            Call Me.Set_List_Cnt_Hojyo()
            Call Chg_BtnKariStatus("hojyo1")
            Set_Forcus("next")

        End Sub

        ''' <summary>
        ''' イベント処理：戻るボタンクリック　
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

            Dim seltab As String = Me.tabCtrlMain.SelectedTab.Name

            Select Case seltab

                '================================================== メインタブ ================================================== -sta
                Case "tabPageStart"         'なし

                Case "tabPageSession"       '戻る [2.接続設定 -> 1.初期設定]
                    tabPageManager.ChangeTabPageVisible(1, True)
                    tabPageManager.ChangeTabPageVisible(2, False)
                    Call Chg_BtnKariStatus("mstart")
                    Call Chg_LeftProgress("mstart")
                    Set_Forcus("next")

                Case "tabPageSyoki"         '戻る [3.初期設定 -> 2.開始画面]
                    tabPageManager.ChangeTabPageVisible(2, True)
                    tabPageManager.ChangeTabPageVisible(3, False)
                    Call Chg_BtnKariStatus("msession")
                    Call Chg_LeftProgress("msession")
                    Set_Forcus("prev")

                Case "tabPageMenu"          '戻る [4.作業選択 -> 3.接続設定]
                    tabPageManager.ChangeTabPageVisible(3, True)
                    tabPageManager.ChangeTabPageVisible(4, False)
                    Call Chg_BtnKariStatus("msyoki")
                    Call Chg_LeftProgress("msyoki")
                    Set_Forcus("prev")

                    '================================================== メインタブ ================================================== -end

                    '================================================== 作業選択タブ ================================================== -sta
                Case "tabPageJizen"         '戻る [5.事前作業 -> ]

                    Select Case tabCtrlJizen.SelectedTab.Name
                        Case "tabPageJizen1"
                            tabPageManager.ChangeTabPageVisible(4, True)
                            tabPageManager.ChangeTabPageVisible(5, False)
                            Me.pnlRefresh.Visible = False                           '件数再読込ボタン表示設定
                            Call Chg_BtnKariStatus("mmenu")
                            Call Chg_LeftProgress("mmenu")
                            Set_Forcus("prev")

                        Case "tabPageJizen2"
                            tabPageManagerJizen.ChangeTabPageVisible(0, True)
                            tabPageManagerJizen.ChangeTabPageVisible(1, False)
                            Call Chg_BtnKariStatus("jizen1")
                            Set_Forcus("jizen_kai")

                        Case "tabPageJizen3"
                            tabPageManagerJizen.ChangeTabPageVisible(1, True)
                            tabPageManagerJizen.ChangeTabPageVisible(2, False)
                            Call Chg_BtnKariStatus("jizen2")
                            Set_Forcus("prev")

                        Case "tabPageJizen4"
                            tabPageManagerJizen.ChangeTabPageVisible(2, True)
                            tabPageManagerJizen.ChangeTabPageVisible(3, False)
                            Call Chg_BtnKariStatus("jizen3")
                            Set_Forcus("prev")

                        Case "tabPageJizen5"
                            tabPageManagerJizen.ChangeTabPageVisible(3, True)
                            tabPageManagerJizen.ChangeTabPageVisible(4, False)
                            Call Chg_BtnKariStatus("jizen3")
                            Set_Forcus("prev")

                        Case Else

                    End Select

                Case "tabPageSelect"        '戻る [9.項目選択 -> 8.はじめに]

                    Select Case tabCtrlCVItem.SelectedTab.Name
                        Case DC_SELTAB_K01
                            tabPageManager.ChangeTabPageVisible(8, True)
                            tabPageManager.ChangeTabPageVisible(9, False)
                            Call Chg_BtnKariStatus("dchajimeni")
                            Call Chg_LeftProgress("dchajimeni", 1)
                            Set_Forcus("prev")
                            If CNVNO = ConvertTypes._汎用 Then
                                Me.pnlRefresh.Visible = False                       '件数再読込ボタン表示設定
                            End If

                        Case DC_SELTAB_K02
                            tabPageManagerCvitem.ChangeTabPageVisible(0, True)
                            tabPageManagerCvitem.ChangeTabPageVisible(1, False)
                            Call Chg_BtnKariStatus(1)
                            Call Chg_LeftProgress("dcselect1", 1)
                            Set_Forcus("prev")

                        Case DC_SELTAB_K03
                            tabPageManagerCvitem.ChangeTabPageVisible(1, True)
                            tabPageManagerCvitem.ChangeTabPageVisible(2, False)
                            Call Chg_BtnKariStatus("dcselect2")
                            Set_Forcus("prev")

                        Case Else
                            '※上記以外のタブは表示させない
                    End Select

                Case "tabPageJikko"         'なし


                Case "tabPageEndOK"         'なし

                    '================================================== 作業選択タブ ================================================== -end

                    '================================================== コンバータータブ ================================================== -sta
                Case "tabPageJigo"          '戻る [6.事後作業 -> ]

                    Select Case tabCtrlJigo.SelectedTab.Name
                        Case "tabPageJigo1"
                            tabPageManager.ChangeTabPageVisible(4, True)
                            tabPageManager.ChangeTabPageVisible(6, False)
                            Me.pnlRefresh.Visible = False                           '件数再読込ボタン表示設定
                            Call Chg_BtnKariStatus("mmenu")
                            Call Chg_LeftProgress("mmenu")
                            Set_Forcus("prev")

                        Case "tabPageJigo2"
                            tabPageManagerJigo.ChangeTabPageVisible(0, True)
                            tabPageManagerJigo.ChangeTabPageVisible(1, False)
                            Call Chg_BtnKariStatus("jigo1")
                            Set_Forcus("next")

                        Case Else

                    End Select

                Case "tabPageHojyo"         '戻る [7.補助機能 -> ]

                    Select Case tabCtrlHojyo.SelectedTab.Name
                        Case "tabPageHojyo1"
                            tabPageManager.ChangeTabPageVisible(4, True)
                            tabPageManager.ChangeTabPageVisible(7, False)
                            Me.pnlRefresh.Visible = False                          '件数再読込ボタン表示設定
                            Call Chg_BtnKariStatus("mmenu")
                            Call Chg_LeftProgress("mmenu")
                            Set_Forcus("prev")

                        Case "tabPageHojyo2"
                            tabPageManagerHojyo.ChangeTabPageVisible(0, True)
                            tabPageManagerHojyo.ChangeTabPageVisible(1, False)
                            Call Chg_BtnKariStatus("hojyo1")
                            Set_Forcus("next")

                        Case Else

                    End Select

                Case Else

            End Select
            '================================================== コンバータータブ ================================================== -end

        End Sub

        ''' <summary>
        ''' イベント処理：次へボタンクリック　
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

            Dim seltab As String = Me.tabCtrlMain.SelectedTab.Name

            Select Case seltab

                '================================================== メインタブ ================================================== -sta
                Case "tabPageStart"     '次へ [1.メイン.開始 -> 2.メイン.接続設定]
                    tabPageManager.ChangeTabPageVisible(2, True)
                    tabPageManager.ChangeTabPageVisible(1, False)
                    Call Chg_BtnKariStatus("msession")
                    Call Chg_LeftProgress("msession")
                    Set_Forcus("next")
                    If btnNext.Enabled Then
                        Set_Forcus("next")
                    Else
                        Set_Forcus("testset")
                    End If

                Case "tabPageSession"   '次へ [2.メイン.接続設定 -> 3.メイン.初期設定]
                    tabPageManager.ChangeTabPageVisible(3, True)
                    tabPageManager.ChangeTabPageVisible(2, False)
                    Call Chg_BtnKariStatus("msyoki")
                    Call Chg_LeftProgress("msyoki")
                    Set_Forcus("next")

                Case "tabPageSyoki"     '次へ [3.メイン.初期設定 -> 4.メイン.作業選択]
                    tabPageManager.ChangeTabPageVisible(4, True)
                    tabPageManager.ChangeTabPageVisible(3, False)
                    Me.pnlRefresh.Visible = False                                                   '件数再読込ボタン表示設定
                    Call Chg_BtnKariStatus("mmenu")
                    Call Chg_LeftProgress("mmenu")
                    '※事前作業の全チェックボックスがONの場合のみ、コンバーター処理ボタンを活性化
                    If Chk_JizenChkitem(Me.tabCtrlJizen.Controls) = False Then
                        btnNext.Enabled = False
                    End If
                    Set_Forcus("menu_jizen")

                Case "tabPageMenu"      'なし [4.メイン.作業選択 -> なし]

                    '================================================== メインタブ ================================================== -end

                    '================================================== 作業選択タブ ================================================== -sta
                Case "tabPageJizen"     '次へ [5.事前作業 -> タブ選択 -> 4.作業選択]
                    Select Case tabCtrlJizen.SelectedTab.Name
                        Case "tabPageJizen1"
                            tabPageManagerJizen.ChangeTabPageVisible(5, False)
                            tabPageManagerJizen.ChangeTabPageVisible(4, False)
                            tabPageManagerJizen.ChangeTabPageVisible(3, False)
                            tabPageManagerJizen.ChangeTabPageVisible(2, False)
                            tabPageManagerJizen.ChangeTabPageVisible(1, True)
                            tabPageManagerJizen.ChangeTabPageVisible(0, False)
                            Call Chg_BtnKariStatus("jizen2")
                            Set_Forcus("jizen_yanuso")
                           
                        Case "tabPageJizen2"
                            tabPageManagerJizen.ChangeTabPageVisible(5, False)
                            tabPageManagerJizen.ChangeTabPageVisible(4, False)
                            tabPageManagerJizen.ChangeTabPageVisible(3, False)
                            tabPageManagerJizen.ChangeTabPageVisible(2, True)
                            tabPageManagerJizen.ChangeTabPageVisible(1, False)
                            tabPageManagerJizen.ChangeTabPageVisible(0, False)
                            Call Chg_BtnKariStatus("jizen3")
                            Set_Forcus("jizen_szenkyshutan")
                       
                        Case "tabPageJizen3"
                            tabPageManagerJizen.ChangeTabPageVisible(5, False)
                            tabPageManagerJizen.ChangeTabPageVisible(4, False)
                            tabPageManagerJizen.ChangeTabPageVisible(3, True)
                            tabPageManagerJizen.ChangeTabPageVisible(2, False)
                            tabPageManagerJizen.ChangeTabPageVisible(1, False)
                            tabPageManagerJizen.ChangeTabPageVisible(0, False)
                            Call Chg_BtnKariStatus("jizen3")
                            Set_Forcus("jizen_nonjisyakoza")
                          
                        Case "tabPageJizen4"
                            tabPageManagerJizen.ChangeTabPageVisible(5, False)
                            tabPageManagerJizen.ChangeTabPageVisible(4, True)
                            tabPageManagerJizen.ChangeTabPageVisible(3, False)
                            tabPageManagerJizen.ChangeTabPageVisible(2, False)
                            tabPageManagerJizen.ChangeTabPageVisible(1, False)
                            tabPageManagerJizen.ChangeTabPageVisible(0, False)
                            Call Chg_BtnKariStatus("jizen4")
                            Set_Forcus("jizen_simesokin")

                        Case "tabPageJizen5"
                            Dim jizenallchk As Boolean = Chk_JizenChkitem(Me.tabCtrlJizen.Controls)
                            tabPageManagerJizen.ChangeTabPageVisible(0, False)                      '※見た目を戻すだけの処理
                            tabPageManagerJizen.ChangeTabPageVisible(1, False)                      '※見た目を戻すだけの処理
                            tabPageManagerJizen.ChangeTabPageVisible(2, False)                      '※見た目を戻すだけの処理
                            tabPageManagerJizen.ChangeTabPageVisible(3, False)                      '※見た目を戻すだけの処理
                            tabPageManagerJizen.ChangeTabPageVisible(4, True)                       '※見た目を戻すだけの処理

                            If jizenallchk = False Then
                                '未チェックありの場合のメッセージを先に出力
                                If MessageBox.Show(MSG_CATION_JIZENCHK, "注意", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = vbYes Then
                                    Exit Sub
                                End If
                            End If
                            tabPageManager.ChangeTabPageVisible(4, True)
                            tabPageManager.ChangeTabPageVisible(5, False)
                            Me.pnlRefresh.Visible = False                                           '件数再読込ボタン表示設定

                            Call Chg_BtnKariStatus("mmenu")
                            If jizenallchk = False Then
                                Me.btnMenuDatacv.Enabled = False
                                Set_Forcus("menu_jizen")
                            Else
                                Me.btnMenuDatacv.Enabled = True
                                Set_Forcus("menu_datacv")
                            End If
                
                        Case "tabPageHJizen1"
                            Dim jizenallchk As Boolean = Chk_JizenChkitem(Me.tabCtrlJizen.Controls)
                            tabPageManagerJizen.ChangeTabPageVisible(0, False)
                            tabPageManagerJizen.ChangeTabPageVisible(1, False)
                            tabPageManagerJizen.ChangeTabPageVisible(2, False)
                            tabPageManagerJizen.ChangeTabPageVisible(3, False)
                            tabPageManagerJizen.ChangeTabPageVisible(4, False)
                            tabPageManagerJizen.ChangeTabPageVisible(5, True)
                            If jizenallchk = False Then
                                '未チェックありの場合のメッセージを先に出力
                                If MessageBox.Show(MSG_CATION_JIZENCHK, "注意", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = vbYes Then
                                    Exit Sub
                                End If
                            End If
                            tabPageManager.ChangeTabPageVisible(4, True)
                            tabPageManager.ChangeTabPageVisible(5, False)

                            Call Chg_BtnKariStatus("mmenu")
                            If jizenallchk = False Then
                                Me.btnMenuDatacv.Enabled = False
                                Set_Forcus("menu_jizen")
                            Else
                                Me.btnMenuDatacv.Enabled = True
                                Set_Forcus("menu_datacv")
                            End If

                        Case Else

                    End Select

                Case "tabPageJigo"      '次へ [6.事後作業 -> タブ選択 -> 4.作業選択] 
                    Select Case tabCtrlJigo.SelectedTab.Name
                        Case "tabPageJigo1"
                            tabPageManagerJigo.ChangeTabPageVisible(1, True)
                            tabPageManagerJigo.ChangeTabPageVisible(0, False)
                            Call Chg_BtnKariStatus("jigo2")
                            Set_Forcus("next")

                        Case "tabPageJigo2"
                            tabPageManager.ChangeTabPageVisible(4, True)
                            tabPageManager.ChangeTabPageVisible(6, False)
                            Me.pnlRefresh.Visible = False                                           '件数再読込ボタン表示設定
                            Call Chg_BtnKariStatus("mmenu")
                            Set_Forcus("menu_hojyo")

                        Case Else

                    End Select

                Case "tabPageHojyo"     '次へ [7.補助機能 -> タブ選択 -> 4.作業選択] 
                    Select Case tabCtrlHojyo.SelectedTab.Name
                        Case "tabPageHojyo1"
                            tabPageManagerHojyo.ChangeTabPageVisible(1, True)
                            tabPageManagerHojyo.ChangeTabPageVisible(0, False)
                            Call Chg_BtnKariStatus("hojyo2")
                            Set_Forcus("next")

                        Case "tabPageHojyo2"
                            tabPageManager.ChangeTabPageVisible(4, True)
                            tabPageManager.ChangeTabPageVisible(7, False)
                            Me.pnlRefresh.Visible = False                                           '件数再読込ボタン表示設定
                            Call Chg_BtnKariStatus("mmenu")
                            Set_Forcus("menu_hojyo")

                        Case Else

                    End Select
                    '================================================== 作業選択タブ ================================================== -end

                    '================================================== コンバータータブ ================================================== -sta
                Case "tabPageHajimeni"  '次へ [4-1.コンバーター.はじめに -> 4-2.対象項目選択]
                    tabPageManager.ChangeTabPageVisible(9, True)
                    tabPageManager.ChangeTabPageVisible(8, False)
                    tabPageManagerCvitem.ChangeTabPageVisible(0, True)
                    tabPageManagerCvitem.ChangeTabPageVisible(1, False)
                    tabPageManagerCvitem.ChangeTabPageVisible(2, False)
                    Call Chg_BtnKariStatus("dcselect1")
                    Call Chg_LeftProgress("dcselect1", 1)
                    pnlMenuDatacv.Refresh()       'kekeke
                    Select Case CNVNO
                        Case ConvertTypes._既存ユーザ用
                            Call Me.Set_ExistCVItemCnt()
                        Case ConvertTypes._汎用
                            Me.pnlRefresh.Visible = True                                            '件数再読込ボタン表示設定
                            '20161104 中間ファイル読込エラー時の処理対応 -chg sta
                            'Call Me.Set_BaseCVItemCnt()
                            If Me.Set_BaseCVItemCnt() = False Then
                                Exit Sub
                            End If
                            '20161104 中間ファイル読込エラー時の処理対応 -chg end
                    End Select

                    'コンバート実績保持
                    Call Me.Set_CVJisseki()

                    Set_Forcus("next")


                Case "tabPageSelect"    '次へ＋実行 [4-2.コンバーター.対象項目選択 -> タブ選択 -> 処理実行]
                    CancelFlg = False
                    MidChkCancelFlg = False                                         'チェックボタンからの中断フラグを初期化

                    Select Case tabCtrlCVItem.SelectedTab.Name
                        Case DC_SELTAB_K01
                            tabPageManagerCvitem.ChangeTabPageVisible(1, True)
                            tabPageManagerCvitem.ChangeTabPageVisible(0, False)
                            Call Chg_BtnKariStatus("dcselect2")
                            Set_Forcus("next")

                        Case DC_SELTAB_K02
                            tabPageManagerCvitem.ChangeTabPageVisible(2, True)
                            tabPageManagerCvitem.ChangeTabPageVisible(1, False)
                            Call Chg_BtnKariStatus("dcselect3")
                            Set_Forcus("next")

                        Case DC_SELTAB_K03
                            Dim list_cvitem As New List(Of String)                  '移行項目取得
                            If Not Chk_CVStartFlg(list_cvitem) Then                 '実行確認メッセージ
                                Exit Sub
                            End If

                            tabPageManager.ChangeTabPageVisible(10, True)
                            tabPageManager.ChangeTabPageVisible(9, False)
                            tabPageManager.ChangeTabPageVisible(1, False)           '開発タブ非表示            
                            btnDevTabChange.Enabled = False                         '開発用ボタン非活性状態設定
                            Me.pnlRefresh.Visible = False                           '件数再読込ボタン表示設定
                            Call Chg_BtnKariStatus("dcjikko1")                      'ボタン状態変更
                            Call Chg_LeftProgress("dcjikko1", 1)                    '画面左進捗画面設定
                            Me.tabPageJikko.Refresh()                               'コンバートタブ画面初期化
                            pnlMenuDatacv.Refresh()       'kekeke

                            '--------------------------------------------------
                            ' メイン処理
                            '--------------------------------------------------
                            If chkCVStart.Checked Then                              '(開発用) コンバート処理有無
                                Call Me.Cnv_Main(list_cvitem)                       '◎コンバートメイン処理
                            End If
                            '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -chg sta
                            'If CancelFlg Then
                            '    tabPageManager.ChangeTabPageVisible(10, False)
                            '    tabPageManager.ChangeTabPageVisible(13, True)
                            '    Call Chg_BtnKariStatus("dcend")                     'ボタン状態変更
                            '    Call Chg_LeftProgress("dcend", 1)                   '画面左進捗画面設定
                            '    Set_Forcus("next")
                            'Else        '続行
                            '    Call Chg_BtnKariStatus("dcjikko2")                  'ボタン状態変更
                            '    Set_Forcus("next")                                  'セットフォーカス
                            'End If
                            If CancelFlg Then
                                If relexeerr = 99 Then
                                    '紐付ツールでエラーが発生した場合の処理
                                    tabPageManager.ChangeTabPageVisible(10, False)
                                    tabPageManager.ChangeTabPageVisible(12, True)
                                    relexeerr = 0
                                Else
                                    tabPageManager.ChangeTabPageVisible(10, False)
                                    tabPageManager.ChangeTabPageVisible(13, True)
                                End If
                                Call Chg_BtnKariStatus("dcend")                     'ボタン状態変更
                                Call Chg_LeftProgress("dcend", 1)                   '画面左進捗画面設定
                                Set_Forcus("next")
                            Else        '続行
                                Call Chg_BtnKariStatus("dcjikko2")                  'ボタン状態変更
                                Set_Forcus("next")                                  'セットフォーカス
                            End If
                            '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -chg end
                            pnlMenuDatacv.Refresh()       'kekeke

                            Dim logdropflg As Boolean = Me.chkLogTblDrop.Checked    '(開発用) ログテーブル削除フラグ
                            Call Me.Set_LogFile(logdropflg)                         'ログ出力

                        Case Else
                            '※上記以外のタブは表示させない
                    End Select

                Case "tabPageJikko"     '次へ [10.実行処理 -> 終了]

                    tabPageManager.ChangeTabPageVisible(10, False)
                    If CancelFlg Then
                        'キャンセル終了時
                        tabPageManager.ChangeTabPageVisible(13, True)
                    Else
                        If Me.txtPartialSituation.Text.IndexOf("×") <= 0 Then
                            '正常終了時
                            tabPageManager.ChangeTabPageVisible(11, True)
                        Else
                            '異常終了時
                            tabPageManager.ChangeTabPageVisible(12, True)
                        End If
                    End If
                    Call Chg_BtnKariStatus("dcend")
                    Call Chg_LeftProgress("dcend", 1)
                    Set_Forcus("next")

                    '--------------------------------------------------
                    ' ハッシュ初期化 
                    '--------------------------------------------------
                    Call Ini_HashTable()

                Case "tabPageEndOK", "tabPageEndCancel", "tabPageEndError"          'ログを開く [終了, キャンセル, エラー]

                    System.Diagnostics.Process.Start(LogFilePath)                   'ログファイル表示

                    '--------------------------------------------------
                    ' ハッシュ初期化 
                    '--------------------------------------------------
                    Call Ini_HashTable()
                    '================================================== コンバータータブ ================================================== -end

                Case Else

            End Select

        End Sub

        ''' <summary>
        ''' イベント処理：終了ボタンクリック
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・コントロールのキャプションから処理を判定
        ''' </remarks>
        Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click

            Dim cnncloseflg As Boolean = False
            Dim frmcloseflg As Boolean = False


            '==================================================
            ' 終了・中断制御
            '==================================================
            Select Case Replace(btnEnd.Text, " ", "")
                Case "中止"
                    MsgResult = MessageBox.Show(MSG_STOP_A, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.Yes Then
                        lblKDatacvEndLabel.Text = MSG_STOP_B
                        Dim tmpcntii As Integer = 2
                        Dim tmpcntjj As Integer = tabCtrlMain.TabPages.Count - 1
                        For cntii As Integer = tmpcntii To tmpcntjj
                            tabPageManager.ChangeTabPageVisible(cntii, False)
                        Next
                        tabPageManager.ChangeTabPageVisible(4, True)
                        Me.pnlRefresh.Visible = False                               '件数再読込ボタン表示設定
                        Call Chg_BtnKariStatus("mmenu")
                        Call Chg_LeftProgress("mmenu")
                        Set_Forcus("end")
                    End If

                Case "終了"
                    Call Chg_LeftProgress("mend")

                    Dim seltab As String = Me.tabCtrlMain.SelectedTab.Name
                    MsgResult = MessageBox.Show(MSG_END_CV, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.Yes Then
                        cnncloseflg = True
                        frmcloseflg = True
                    Else
                        Select Case seltab
                            Case "tabPageStart"
                                Call Chg_LeftProgress("mstart")
                            Case "tabPageSession"
                                Call Chg_LeftProgress("msession")
                            Case "tabPageSyoki"
                                Call Chg_LeftProgress("msyoki")
                            Case "tabPageMenu"
                                Call Chg_LeftProgress("mmenu")
                        End Select
                        Set_Forcus("menu_hojyo")
                    End If

                Case "キャンセル"
                    MsgResult = MessageBox.Show(MSG_CANCEL_A, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.Yes Then
                        '処理中断
                        CancelFlg = True
                    Else
                        '処理再開
                    End If

                Case "作業選択へ"
                    'コンバート実績保持
                    Dim seltab As String = Me.tabCtrlMain.SelectedTab.Name
                    If seltab = "tabPageEndOK" Then
                        MsgResult = MessageBox.Show(MSG_END_CV_TORIREKI, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2)
                        If MsgResult = DialogResult.Yes Then
                            Call Me.MakeFile_CVJisseki()
                        Else
                            'ファイル削除
                            Call Me.DeleteFile_CVJisseki()
                        End If
                    End If

                    tabPageManager.ChangeTabPageVisible(4, True)
                    tabPageManager.ChangeTabPageVisible(11, False)
                    Me.pnlRefresh.Visible = False                                   '件数再読込ボタン表示設定
                    Call Chg_BtnKariStatus("mmenu")
                    Call Chg_LeftProgress("mmenu")
                    If seltab = "tabPageEndOK" Then
                        Select Case CNVNO
                            Case ConvertTypes._汎用
                                Set_Forcus("menu_jigo")
                            Case Else
                                Set_Forcus("menu_gazocv")
                        End Select
                    Else
                        Set_Forcus("menu_datacv")
                    End If

                Case Else

            End Select


            '==================================================
            ' 終了処理 (クローズ処理)
            '==================================================
            If cnncloseflg Then
                cnnv7.CnnClose(sqlcnnv7)
                cnnv10.CnnClose(sqlcnnv10)
            End If

            If frmcloseflg Then
                Me.Close()
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：中間ファイル参照ボタン
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnFileSeach_Click(sender As Object, e As EventArgs) Handles btnMidDirSeach.Click, btnMidDirSeach2.Click

            Dim dirdialog As New OpenFileDialog()
            dirdialog.FileName = FILE_HMIDD_NAME
            dirdialog.Title = "ファイルを指定してください。"
            dirdialog.Filter = "EXCELファイル(*.xlsx)|*.xlsx"
            dirdialog.RestoreDirectory = True

            '20160926 中間ファイル格納先参照フォルダボタンのイベント修正 -chg sta
            'Select Case True
            '    Case sender Is Me.btnLogDirSeach    'btnMidDirSeach
            '        If Me.txtMidDirPath.Text <> "" Then
            '            dirdialog.InitialDirectory = Me.txtMidDirPath.Text
            '        Else
            '            dirdialog.InitialDirectory = Environment.SpecialFolder.Desktop
            '        End If
            'End Select
            Select Case True
                Case sender Is Me.btnMidDirSeach
                    If Me.txtMidDirPath.Text <> "" Then
                        dirdialog.InitialDirectory = Path.GetDirectoryName(Me.txtMidDirPath.Text)
                    Else
                        dirdialog.InitialDirectory = Environment.SpecialFolder.Desktop
                    End If
                Case sender Is Me.btnMidDirSeach2
                    If Me.txtMidDirPath2.Text <> "" Then
                        dirdialog.InitialDirectory = Path.GetDirectoryName(Me.txtMidDirPath2.Text)
                    Else
                        dirdialog.InitialDirectory = Environment.SpecialFolder.Desktop
                    End If
            End Select
            '20160926 中間ファイル格納先参照フォルダボタンのイベント修正 -chg end

            If dirdialog.ShowDialog() = DialogResult.OK Then
                Me.txtMidDirPath.Text = dirdialog.FileName
                Me.txtMidDirPath2.Text = dirdialog.FileName
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：フォルダ参照ボタン  '20160912 既存用→汎用用中間ファイルへのコピー処理 btnExistMidToBaseMidDirSerach を追加
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnDirSeach_Click(sender As Object, e As EventArgs) Handles _
                                            btnLogDirSeach.Click, btnMidDirLogSeach.Click, btnRelDirSeach.Click, _
                                            btnJizenListDirSeach.Click, btnJigoListDirSeach.Click, btnKensyoListDirSeach.Click, _
                                            btnExistMidToBaseMidDirSerach.Click

            Dim dirdialog As New FolderBrowserDialog

            '上部に表示する説明の文字列設定
            dirdialog.Description = "フォルダを指定してください。"

            'ルートフォルダ設定(デフォルトでDesktop)
            dirdialog.RootFolder = Environment.SpecialFolder.Desktop

            '最初に選択するフォルダを指定する
            Select Case True
                Case sender Is Me.btnLogDirSeach
                    If Me.txtLogDirPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtLogDirPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnRelDirSeach
                    If Me.txtRelationDirPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtRelationDirPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnMidDirLogSeach
                    If Me.txtMidDirLogPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtMidDirLogPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnJizenListDirSeach
                    If Me.txtJizenListPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtJizenListPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnJigoListDirSeach
                    If Me.txtJigoListPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtJigoListPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnKensyoListDirSeach
                    If Me.txtKensyoListPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtKensyoListPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnExistMidToBaseMidDirSerach         '20160912 既存用→汎用用中間ファイルへのコピー処理 -add
                    If Me.txtExistMidToBaseMid.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtExistMidToBaseMid.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
            End Select

            'フォルダの新規作成可否設定(True:作成可)
            dirdialog.ShowNewFolderButton = True

            'ダイアログ表示
            If dirdialog.ShowDialog(Me) = DialogResult.OK Then
                Select Case True
                    Case sender Is Me.btnLogDirSeach
                        Me.txtLogDirPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnRelDirSeach
                        Me.txtRelationDirPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnMidDirLogSeach
                        Me.txtMidDirLogPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnJizenListDirSeach
                        Me.txtJizenListPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnJigoListDirSeach
                        Me.txtJigoListPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnKensyoListDirSeach
                        Me.txtKensyoListPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnExistMidToBaseMidDirSerach     '20160912 既存用→汎用用中間ファイルへのコピー処理 -add
                        Me.txtExistMidToBaseMid.Text = dirdialog.SelectedPath
                End Select
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：同意ボタン
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnDoui_Click(sender As Object, e As EventArgs) Handles btnDoui.Click

            Me.btnNext.Enabled = True
            Me.btnDoui.Enabled = False
            Set_Forcus("next")

        End Sub

        ''' <summary>
        ''' イベント処理：事前作業ボタン.リスト出力機能　
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・解約未処理データ処理
        ''' ・預り金データ処理
        ''' ・マイナス金額調整
        ''' ・物件法令・権利の設定値調整
        ''' ・家主固定控除の発生月調整
        ''' ・随時変動費の発生月調整
        ''' ・家主送金先未設定データの確認・調整
        ''' ・修繕対象負担者設定契約者データの確認
        ''' ・部屋鍵情報抽出
        ''' ・契約鍵情報抽出
        ''' ・口座実情報未登録データ抽出機能
        ''' </remarks>
        Private Sub btnListJizenOutput_Click(sender As Object, e As EventArgs) Handles _
                                                    btnListJizenKai.Click, btnListJizenAzu.Click, btnListJizenMinus.Click, _
                                                    btnListJizenBkhourei.Click, btnListJizenHasseiOw.Click, btnListJizenHasseiHen.Click, _
                                                    btnListJizenYanuso.Click, btnListJizenSzenKyshutan.Click, btnListJizenHyKagi.Click, _
                                                    btnListJizenKyKagi.Click, btnListJizenNonJisyaKoza.Click, btnListJizenSimeSokin.Click, _
                                                    btnListJizenKozameigikana.Click
            Dim btnname As String = sender.Name
            Call Me.Set_List_Output(btnname)

        End Sub
     
        ''' <summary>
        ''' イベント処理：変換前関連ファイルパステキストボックス '20160711 関連ファイルパスの件数表示処理
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub txtBfRelRename_LostFocus(sender As Object, e As EventArgs) Handles txtBfRelRename.LostFocus

            Call Me.Set_List_Cnt_Hojyo_Rename(Me.txtBfRelRename.Text)

        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.関連ファイルパスのリネーム '20160707 関連ファイルリネーム処理の追加
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnRNRelRename_Click(sender As Object, e As EventArgs) Handles btnRNRelRename.Click

            Call Me.Rename_RelFileName()

        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.未使用家主リスト出力 '20160711 未使用データ削除_家主
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnListKiOpOw_Click(sender As Object, e As EventArgs) Handles btnListKiOpOw.Click

            Dim btnname As String = sender.Name
            Call Me.Set_List_Output(btnname)

        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.未使用家主削除 '20160711 未使用データ削除_家主
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnKiOpOw_Click(sender As Object, e As EventArgs) Handles btnKiOpOw.Click

            Call Me.Main_NotUseData_Delete("家主")

        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.未使用契約者リスト出力 '20160711 未使用データ削除_契約者
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnListKiOpKys_Click(sender As Object, e As EventArgs) Handles btnListKiOpKys.Click

            Dim btnname As String = sender.Name
            Call Me.Set_List_Output(btnname)

        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.未使用契約者削除 '20160711 未使用データ削除_契約者
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnKiOpKys_Click(sender As Object, e As EventArgs) Handles btnKiOpKys.Click

            Call Me.Main_NotUseData_Delete("契約者")

        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.分割入金の未収分データの確認 '20160707 検証作業リスト出力処理の追加_分割入金未収分
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnListBunkatumisyu_Click(sender As Object, e As EventArgs) Handles btnListBunkatumisyu.Click

            Dim btnname As String = sender.Name
            Call Me.Set_List_Output(btnname)
            
        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.控除支払データの確認 '20160707 検証作業リスト出力処理の追加_控除支払
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnListKojyosh_Click(sender As Object, e As EventArgs) Handles btnListKojyosh.Click

            Dim btnname As String = sender.Name
            Call Me.Set_List_Output(btnname)
         
        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.修繕項目毎契約者送金率データの確認 '20160707 検証作業リスト出力処理の追加_修繕項目毎契約者送金率
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnListSzenKysSorit_Click(sender As Object, e As EventArgs) Handles btnListSzenKysSorit.Click

            Dim btnname As String = sender.Name
            Call Me.Set_List_Output(btnname)
       
        End Sub

        ''' <summary>
        ''' イベント処理：検証作業ボタン.紐付設定内容のリスト出力 '20160720 紐付設定ファイル出力機能の追加
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnListRelationSet_Click(sender As Object, e As EventArgs) Handles btnListRelationSet.Click

            '格納先フォルダ有無確認
            Dim tmp_jizenlistdir As String = Me.txtJizenListPath.Text
            If EtcMethod.Chk_DirExist(tmp_jizenlistdir) = False Then
                MsgResult = MessageBox.Show(MSG_ERR_DIR_LIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            Dim listdirpath As String = tmp_jizenlistdir

            '紐付ファイル有無確認(念の為)
            Dim reldirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_RELEXEDIR_NAME)
            Dim relfiledirpath As String = EtcMethod.Set_Path(reldirpath, DIR_REL_NAME)
            Dim relfilepath As String = EtcMethod.Set_Path(relfiledirpath, MIDFILE_RELNAME & ".xlsx")

            If EtcMethod.Chk_FileExist(relfilepath) = False Then
                MsgResult = MessageBox.Show(MSG_RELFILE_NOTEXISTERR, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            '出力確認
            MsgResult = MessageBox.Show(MSG_RELLIST_OUTSTA, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                Exit Sub
            End If

            'ファイルコピー
            Dim postfilename As String = "_" & Replace((Replace(Replace((String.Format(Now)), ":", ""), "/", "")), " ", "") & ".xlsx"
            Dim relcopyfilename As String = EtcMethod.Set_Path(listdirpath, MIDFILE_RELNAME & postfilename)
            File.Copy(relfilepath, relcopyfilename)

            '終了処理
            MsgResult = MessageBox.Show(MSG_LIST_OUTEND, "成功", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            System.Diagnostics.Process.Start(relcopyfilename)

        End Sub

        ''' <summary>
        ''' イベント処理：コンバート履歴処理(画面上のチェックボックス着色・チェックON/OFFの初期化) '20160711 コンバート履歴処理の追加
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnRekiClear_Click(sender As Object, e As EventArgs) Handles btnRekiClear.Click

            '確認メッセージ
            MsgResult = MessageBox.Show(MSG_DEL_CV_JISSEKI, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                Exit Sub
            Else
                'チェックボックスのチェックON、着色初期化
                '項目名とチェックボックスを紐付けたハッシュテーブルを生成
                Dim hash_cvitemtochkbox As New Hashtable
                hash_cvitemtochkbox = Get_Hash_CVChkitemAll()
                For Each chkbox In hash_cvitemtochkbox
                    Dim tmp_chkbox As New CheckBox
                    tmp_chkbox = chkbox.Value
                    tmp_chkbox.ForeColor = Color.Black
                    '20161018 UI最終確認での修正 -chg sta
                    'tmp_chkbox.Checked = True
                    Select Case CNVNO
                        Case ConvertTypes._汎用
                            tmp_chkbox.Checked = False
                        Case ConvertTypes._既存ユーザ用
                            tmp_chkbox.Checked = True
                    End Select
                    '20161018 UI最終確認での修正 -chg end
                Next

                '実績が記録されたファイルの削除
                Call Me.DeleteFile_CVJisseki()

                '終了メッセージ
                MsgResult = MessageBox.Show(MSG_DEL_CV_JISSEKIEND, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

                'ボタンの活性制御
                btnRekiClear.Enabled = False

            End If

            Set_Forcus("next")

        End Sub

        ''' <summary>
        ''' イベント処理：件数表示リフレッシュボタン
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

            '選択されているタブ名を取得
            Dim seltab As String = Me.tabCtrlMain.SelectedTab.Name

            '選択されているタブ毎に処理を分岐
            Select Case seltab
                Case "tabPageJizen"                                 '事前作業タブ
                    Call Me.Set_List_Cnt_Jizen()
                Case "tabPageHojyo"                                 '補助作業タブ
                    Call Me.Set_List_Cnt_Hojyo()
                Case "tabPageSelect"
                    If CNVNO = ConvertTypes._汎用 Then              '※汎用時のみ表示するボタンだが念の為
                        Call Me.Set_BaseCVItemCnt()
                    End If
            End Select

        End Sub

        ''' <summary>
        ''' イベント処理：チェックボックス値変更時挙動
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub chkJizenKai_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenKai.CheckedChanged
            Set_Forcus("jizen_azu")
        End Sub

        Private Sub chkJizenAzu_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenAzu.CheckedChanged
            Set_Forcus("jizen_hurikae")
        End Sub

        Private Sub chkJizenHurikae_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenHurikae.CheckedChanged
            Set_Forcus("next")
        End Sub

        Private Sub chkJizenSo_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenSo.CheckedChanged
            'Set_Forcus("next")
        End Sub

        Private Sub chkJizenMinus_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenMinus.CheckedChanged
            Set_Forcus("jizen_bkhourei")
        End Sub

        Private Sub chkJizenBkhourei_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenBkhourei.CheckedChanged
            Set_Forcus("jizen_hasseiow")
        End Sub

        Private Sub chkJizenHasseiOw_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenHasseiOw.CheckedChanged
            Set_Forcus("jizen_hasseihen")
        End Sub

        Private Sub chkJizenHasseiHen_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenHasseiHen.CheckedChanged
            Set_Forcus("next")
        End Sub

        Private Sub chkJizenYanuso_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenYanuso.CheckedChanged
            Set_Forcus("jizen_bkhourei")
        End Sub

        Private Sub chkJizenSzenKyshutan_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenSzenKyshutan.CheckedChanged
            Set_Forcus("jizen_kagi")
        End Sub

        Private Sub chkJizenKagi_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenKagi.CheckedChanged
            Set_Forcus("next")
        End Sub

        Private Sub chkHJizenTyukan_CheckedChanged(sender As Object, e As EventArgs) Handles chkHJizenTyukan.CheckedChanged
            Set_Forcus("next")
        End Sub

        Private Sub chkHJizenGazoKeisiki_CheckedChanged(sender As Object, e As EventArgs) Handles chkHJizenGazoKeisiki.CheckedChanged
            'Set_Forcus("next")
        End Sub

        Private Sub chkJizenNonJisyaKoza_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenNonJisyaKoza.CheckedChanged
            Set_Forcus("jizen_kozameigikana")
        End Sub

        Private Sub chkJizenSimeSokin_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenSimeSokin.CheckedChanged '20160829 事前作業_送金予定日リスト出力機能を追加 -add
            Set_Forcus("next")
        End Sub

        Private Sub chkJizenKozameigikana_CheckedChanged(sender As Object, e As EventArgs) Handles chkJizenKozameigikana.CheckedChanged
            Set_Forcus("next")
        End Sub

        ''' <summary>
        ''' イベント処理：接続テストボタン
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・接続テストが成功した場合のみ、次の処理へ進めるようにする
        ''' </remarks>
        Private Sub btnConnectTest_Click(sender As Object, e As EventArgs) Handles btnConnectTest.Click

            Dim flg_authentV7 As Boolean = optV7Authent1.Checked
            Dim flg_authentV10 As Boolean = optV10Authent1.Checked
            Dim condb As String = ""


            '20160829 対象革命のバージョン判定処理を追加 -add sta
            '--------------------------------------------------
            ' 革命バージョンチェック(V7,Regなしの場合のみCV可)
            '--------------------------------------------------
            If Not (CNVNO = ConvertTypes._汎用) Then
                Dim kakumeiver As String = Me.Chk_KakumeiVer()
                If kakumeiver = "" Then
                    'バージョン情報が存在しない場合
                    MsgResult = MessageBox.Show(MSG_ERR_VERBLANK, "確認", _
                                                MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.No Then
                        MsgResult = MessageBox.Show(MSG_ERR_VERBLANKEND, "終了", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Me.Close()
                    End If
                ElseIf Chk_KakumeiVer() <> "7" Then
                    'V7以外のバージョンの場合
                    MsgResult = MessageBox.Show(MSG_ERR_VERNOTV7, "終了", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.Close()
                End If
            End If
            '20160829 対象革命のバージョン判定処理を追加 -add end

            'コントロール設定値を接続情報へ格納
            Me.Set_Control_To_ConModel(fstmodelv7)
            Me.Set_Control_To_ConModel(fstmodelv10, 1)

            Dim flg As Boolean = True

            'V7接続処理
            If Not (CNVNO = ConvertTypes._汎用) Then
                If (fstmodelv7.ServerName = "" OrElse fstmodelv7.InitialCatalog = "") OrElse cnnv7.CnnSession(fstmodelv7, sqlcnnv7, flg_authentV7) = False Then
                    condb = CV_FROM_NAME & "DB"
                    flg = False
                End If
            End If

            'V10接続処理
            If (fstmodelv10.ServerName = "" OrElse fstmodelv10.InitialCatalog = "") OrElse cnnv10.CnnSession(fstmodelv10, sqlcnnv10, flg_authentV10) = False Then
                If condb = "" Then
                    condb = CV_TO_NAME & "DB"
                Else
                    condb = condb & "、" & CV_TO_NAME & "DB"
                End If
                flg = False
            End If

            '接続結果
            If flg Then
                '成功時
                MsgResult = MessageBox.Show(MSG_CNN_SUCCESS_A, "成功", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                btnNext.Enabled = True
                btnConnectTest.Enabled = False
                grpV7ConnectInfo.Enabled = False
                grp10ConnectInfo.Enabled = False
                txtTimeOut.Enabled = False
                btnDefConInfoRead.Enabled = False

                '接続情報を保存
                ''20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
                'Call Me.MakeFile_DBConInfo(fstmodelv7, 0)
                'Call Me.MakeFile_DBConInfo(fstmodelv10, 1)
                If Me.MakeFile_DBConInfo(fstmodelv7, 0) = False Then
                    Me.Close()
                    Exit Sub
                End If
                If Me.MakeFile_DBConInfo(fstmodelv10, 1) = False Then
                    Me.Close()
                    Exit Sub
                End If
                '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end
                Set_Forcus("next")
            Else
                '失敗時
                MsgResult = MessageBox.Show(MSG_CNN_FAILURE_A & vbCrLf & "接続先：" & condb, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                cnnv7.CnnClose(sqlcnnv7)
                cnnv10.CnnClose(sqlcnnv10)
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：移行項目全選択
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・各タブの最左上部のチェックボックスを基準にする
        ''' </remarks>
        Private Sub btnAllChk_Click(sender As Object, e As EventArgs) Handles btnAllChk.Click

            Call Set_ChkboxOnOff(Me.tabCtrlCVItem.SelectedTab.Name, True)
            Call Set_OptSelect()

        End Sub

        ''' <summary>
        ''' イベント処理：オプションボタン設定値変更
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub opt_CheckedChanged(sender As Object, e As EventArgs) Handles _
                                            optV7Authent1.CheckedChanged, optV7Authent2.CheckedChanged, _
                                            optV10Authent1.CheckedChanged, optV10Authent2.CheckedChanged

            Select Case True
                Case sender Is Me.optV7Authent1
                    Me.txtV7User.Enabled = sender.Checked
                    Me.txtV7Pass.Enabled = sender.Checked
                Case sender Is Me.optV10Authent1
                    Me.txtV10User.Enabled = sender.Checked
                    Me.txtV10Pass.Enabled = sender.Checked
            End Select

        End Sub

        ''' <summary>
        ''' イベント処理：中間ファイルチェックボタン
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnMidFileCheck_Click(sender As Object, e As EventArgs) Handles btnMidFileCheck.Click

            Dim normalflg As Boolean = True
            Dim btnflg As Boolean = True
            Dim errstr As String = ""
            Dim btntxt As String = Trim(sender.Text)


            Select Case Trim(btntxt)
                Case "キャンセル"
                    MsgResult = MessageBox.Show(MSG_CANCEL_B, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.Yes Then
                        MidChkCancelFlg = True
                        Call Chg_BtnKariStatus("midcheck")
                    End If

                Case "チェック"

                    MidChkCancelFlg = False                                                     '中断フラグを初期化
                    CancelFlg = False                                                           '実行ボタンの中断フラグを初期化
                    Dim list_cv As New List(Of String)                                          '移行項目取得

                    If Not Chk_CVStartFlg(list_cv, btnflg) Then                                 '実行確認メッセージ→実行有無
                        Exit Sub
                    End If

                    Me.btnMidFileCheck.Image = Converter10.My.Resources.Resources.x
                    Call Chg_BtnKariStatus("midstop")

                    '----- 中間用プログレスバー初期化 -----
                    If CNVNO = ConvertTypes._汎用 Then
                        Me.lblCheckSituation.Text = SITUATION_MID_READ
                        Me.lblPgbCheck.Text = "0 %"
                        Me.pgbCheck.Value = 0
                        Me.pnlPrgChk.Visible = True
                        Me.tabCtrlCVItem.Enabled = False
                        Me.btnHJizenTyukanOpen2.Enabled = False
                        Me.btnMidDirSeach2.Enabled = False
                        Me.btnMidDirLogSeach.Enabled = False
                        Me.pnlRefresh.Enabled = False
                        Me.Refresh()
                    End If

                    '--------------------------------------------------
                    ' ハッシュテーブル初期化 
                    '--------------------------------------------------
                    Call Me.Ini_HashTable()

                    '--------------------------------------------------
                    ' 中間ファイル作成 
                    '--------------------------------------------------
                    '汎用用中間ファイルから既存用中間ファイルへのコピー処理
                    normalflg = Me.Set_BaseMidFileToExistingMidFile(list_cv, errstr)
                    If normalflg = False Then
                        MsgResult = MessageBox.Show(errstr, "注意", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                        Call Chg_BtnKariStatus("midcheck")
                        Me.btnMidFileCheck.Image = Converter10.My.Resources.Resources.機能
                        Exit Sub
                    End If

                    '--------------------------------------------------
                    ' 中間ファイルチェック
                    '--------------------------------------------------
                    '20161009 改善対応：中断処理が中断しない -chg sta
                    'Call Me.Before_DBWrite(list_cv)                                             '実行前準備 
                    'normalflg = Me.Chk_MidFileChk_Main(list_cv)                                 'チェック   
                    If MidChkCancelFlg = False Then
                        Call Me.Before_DBWrite(list_cv)                                         '実行前準備 
                        normalflg = Me.Chk_MidFileChk_Main(list_cv)                             'チェック                  
                    End If
                    '20161009 改善対応：中断処理が中断しない -chg end

                    '--------------------------------------------------
                    ' エラー時対処 (20160222)
                    '--------------------------------------------------
                    Dim tmp_endmsg As String = ""
                    If MidChkCancelFlg Then                                                     '中断時
                        tmp_endmsg = MSG_STOP_MIDCHK
                    ElseIf normalflg = False Then                                               'エラー時
                        tmp_endmsg = MSG_ERR_CHK_MID
                    Else                                                                        '通常
                        tmp_endmsg = MSG_END_MIDCHK
                    End If
                    MsgResult = MessageBox.Show(tmp_endmsg, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

                    '20161009 改善対応：中断処理が中断しない -add sta
                    'System.Diagnostics.Process.Start(MiddleLogFilePath)                         'ログファイル表示
                    If MidChkCancelFlg = False Then
                        System.Diagnostics.Process.Start(MiddleLogFilePath)                     'ログファイル表示
                    End If
                    '20161009 改善対応：中断処理が中断しない -add end
                    Call Chg_BtnKariStatus("midcheck")

                    Me.pnlPrgChk.Visible = False                                                '中間用プログレスバー表示設定
                    Me.tabCtrlCVItem.Enabled = True
                    Me.btnHJizenTyukanOpen2.Enabled = True
                    Me.btnMidDirSeach2.Enabled = True
                    Me.btnMidDirLogSeach.Enabled = True
                    Me.pnlRefresh.Enabled = True
                    Me.btnMidFileCheck.Image = Converter10.My.Resources.Resources.機能          'ボタンのアイコンセット
                Case Else

            End Select

        End Sub

        ''' <summary>
        ''' イベント処理：ログ出力件数入力箇所制御
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub txtLogOutputCnt_KeyPress(sender As Object, e As KeyPressEventArgs)

            If (e.KeyChar < "0"c Or "9"c < e.KeyChar) And e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：接続情報の初期値読込
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnDefConInfoRead_Click(sender As Object, e As EventArgs) Handles btnDefConInfoRead.Click

            'デフォルト値読込
            Call Me.Set_V7DBConInfo()
            Call Me.Set_10DBConInfo()

            'コントロールへセット
            Me.Set_DefConInfo_To_Control(fstmodelv7, fstmodelv10)

            'セットフォーカス
            Set_Forcus("testset")

        End Sub

        ''' <summary>
        ''' イベント処理：DB接続情報タブのタイムアウト値入力制御 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub txtTimeOut_KeyPress(sender As Object, e As KeyPressEventArgs)

            '0～9と、バックスペース以外の時は、イベントをキャンセルする
            If (e.KeyChar < "0"c OrElse "9"c < e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：チェックボックスチェンジ
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>・既存システム用-項目1</remarks>
        Private Sub chkKiMstBus_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstBus.CheckedChanged
            chkMstBus.Checked = chkKiMstBus.Checked
            chkMstBusKotu.Checked = chkKiMstBus.Checked
        End Sub

        Private Sub chkKiMstSchool_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstSchool.CheckedChanged
            chkMstSchool.Checked = chkKiMstSchool.Checked
        End Sub

        Private Sub chkKiMstArea_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstArea.CheckedChanged
            chkMstArea.Checked = chkKiMstArea.Checked
        End Sub

        Private Sub chkKiMstHokenrui_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstHokenrui.CheckedChanged
            chkMstHokenrui.Checked = chkKiMstHokenrui.Checked
        End Sub

        '20160516 契約分類削除に伴うコメントアウト -del sta
        'Private Sub chkKiMstKeiyakurui_CheckedChanged(sender As Object, e As EventArgs)
        '    chkMstKeiyakurui.Checked = chkKiMstKeiyakurui.Checked
        'End Sub
        '20160516 契約分類削除に伴うコメントアウト -del end

        Private Sub chkKiMstTokuyaku_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstTokuyaku.CheckedChanged
            chkMstTokuyaku.Checked = chkKiMstTokuyaku.Checked
        End Sub

        Private Sub chkKiMstKasyoClaimrui_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstKasyoClaimrui.CheckedChanged
            chkMstKasyoClaimrui.Checked = chkKiMstKasyoClaimrui.Checked
        End Sub

        Private Sub chkKiMstHendo_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstHendo.CheckedChanged
            chkMstHendo.Checked = chkKiMstHendo.Checked
            chkMstHendoitiran.Checked = chkKiMstHendo.Checked
        End Sub

        Private Sub chkKiMstKagititle_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiMstTitle.CheckedChanged
            chkMstKagititle.Checked = chkKiMstTitle.Checked
            '20160801 タイトルマスタ統合処理 -add sta
            chkMstBikotitle.Checked = chkKiMstTitle.Checked
            chkMstBikolst.Checked = chkKiMstTitle.Checked
            chkMstGazotitle.Checked = chkKiMstTitle.Checked
            '20160801 タイトルマスタ統合処理 -add end
        End Sub

        '20160801 タイトルマスタ統合処理 -del sta
        'Private Sub chkKiMstBikotitle_CheckedChanged(sender As Object, e As EventArgs)
        '    chkMstBikotitle.Checked = chkKiMstBikotitle.Checked
        '    chkMstBikolst.Checked = chkKiMstBikotitle.Checked
        'End Sub
        '20160801 タイトルマスタ統合処理 -del end

        ''' <summary>
        ''' イベント処理：チェックボックスチェンジ
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub chkKiJisyaBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiJisyaBase.CheckedChanged
            Call Me.Set_ChkBoxValue(Me.grpJisya, Me.chkKiJisyaBase.Checked, 1)
        End Sub

        Private Sub chkKiOwBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiOwBase.CheckedChanged
            Call Me.Set_ChkBoxValue(Me.grpOw, Me.chkKiOwBase.Checked, 1)
        End Sub

        Private Sub chkKiBkBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiBkBase.CheckedChanged
            Call Me.Set_ChkBoxValue(Me.grpBk, Me.chkKiBkBase.Checked, 1)
        End Sub

        Private Sub chkKiHyBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiHyBase.CheckedChanged
            Call Me.Set_ChkBoxValue(Me.grpHy, Me.chkKiHyBase.Checked, 1)
            chkHySetubi.Checked = chkKiHySetubi.Checked
            Call Me.Set_ChkBoxValue(Me.grpSorule, Me.chkKiHyBase.Checked, 1)
            chkKiHySetubi.Enabled = chkKiHyBase.Checked
            chkHySetubi.Enabled = chkKiHyBase.Checked
            If chkKiHyBase.Checked = False Then
                chkKiHySetubi.Checked = chkKiHyBase.Checked
            End If
        End Sub

        Private Sub chkKiHySetubi_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiHySetubi.CheckedChanged
            chkHySetubi.Checked = chkKiHySetubi.Checked
        End Sub

        Private Sub chkKiKysBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiKysBase.CheckedChanged
            Call Me.Set_ChkBoxValue(Me.grpKys, Me.chkKiKysBase.Checked, 1)
        End Sub

        Private Sub chkKiKyBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiKyBase.CheckedChanged
            Call Me.Set_ChkBoxValue(Me.grpKy, Me.chkKiKyBase.Checked, 1)
        End Sub

        Private Sub optKiKagi_CheckedChanged(sender As Object, e As EventArgs) Handles optKiHyKagi.CheckedChanged, optKiKyKagi.CheckedChanged
            optHyKagi.Checked = optKiHyKagi.Checked
            optKyKagi.Checked = optKiKyKagi.Checked
        End Sub

        Private Sub chkKiGyCyukaiBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiGyCyukaiBase.CheckedChanged
            chkGyCyukaiBase.Checked = chkKiGyCyukaiBase.Checked
            chkGyCyukaiKoza.Checked = chkKiGyCyukaiBase.Checked
            chkGyCyukaiMemo.Checked = chkKiGyCyukaiBase.Checked
        End Sub

        Private Sub chkKiGySyuzenBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiGySyuzenBase.CheckedChanged
            chkGySyuzenBase.Checked = chkKiGySyuzenBase.Checked
            chkGySyuzenKoza.Checked = chkKiGySyuzenBase.Checked
            chkGySyuzenMemo.Checked = chkKiGySyuzenBase.Checked
        End Sub

        Private Sub chkKiGyLifelineBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiGyLifelineBase.CheckedChanged
            chkGyLifelineBase.Checked = chkKiGyLifelineBase.Checked
        End Sub

        Private Sub chkKiGyHokenBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiGyHokenBase.CheckedChanged
            chkGyHokenBase.Checked = chkKiGyHokenBase.Checked
            chkGyHokenKoza.Checked = chkKiGyHokenBase.Checked
            chkGyHokenMemo.Checked = chkKiGyHokenBase.Checked
        End Sub

        Private Sub chkKiGyYatinhosyoBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiGyYatinhosyoBase.CheckedChanged
            chkGyYatinhosyoBase.Checked = chkKiGyYatinhosyoBase.Checked
            chkGyYatinhosyoMemo.Checked = chkKiGyYatinhosyoBase.Checked
        End Sub

        Private Sub chkKiGySisetuBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiGySisetuBase.CheckedChanged
            chkGySisetuBase.Checked = chkKiGySisetuBase.Checked
        End Sub

        Private Sub chkKiGySekoBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiGySekoBase.CheckedChanged
            chkGySekoBase.Checked = chkKiGySekoBase.Checked
        End Sub

        Private Sub chkKiSqBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiSqBase.CheckedChanged                 '20160519 請求情報追加
            Call Me.Set_ChkBoxValue(Me.grpSq, Me.chkKiSqBase.Checked, 1)
        End Sub

        Private Sub chkKiClaimBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiClaimBase.CheckedChanged           '20160524 クレーム情報移行処理実装
            Call Me.Set_ChkBoxValue(Me.grpClaim, Me.chkKiClaimBase.Checked, 1)
        End Sub

        Private Sub chkKiSyskanriBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiSyskanriBase.CheckedChanged     '20160616 初期設定情報移行処理実装
            Call Me.Set_ChkBoxValue(Me.grpSyskanri, Me.chkKiSyskanriBase.Checked, 1)
        End Sub

        Private Sub chkKiSzenBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiSzenBase.CheckedChanged             '20160621 修繕関連移行処理追加
            Call Me.Set_ChkBoxValue(Me.grpSzen, Me.chkKiSzenBase.Checked, 1)
        End Sub

        Private Sub chkKiRendoBase_CheckedChanged(sender As Object, e As EventArgs) Handles chkKiRendoBase.CheckedChanged           '20160720 連動情報構築
            Call Me.Set_ChkBoxValue(Me.grpRendo, Me.chkKiRendoBase.Checked, 1)
        End Sub

        ''' <summary>
        ''' イベント処理：オプションチェックボックスON/OFFと各種作業のグレイアウトの連動
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・契約管理、請求管理、入金管理、支払管理、FB
        ''' </remarks>
        Private Sub chkOptionSelect_kysq_CheckedChanged(sender As Object, e As EventArgs) Handles _
                                                            chkOptionSelectKy.CheckedChanged, chkOptionSelectSq.CheckedChanged, chkOptionSelectNk.CheckedChanged, _
                                                            chkOptionSelectSh.CheckedChanged, chkOptionSelectFB.CheckedChanged

            '画面起動時は処理を抜ける
            If frmloadflg_kysq Then
                Exit Sub
            End If

            '--------------------------
            'グレイアウトの初期化
            '--------------------------
            '事前
            Me.pnlJizenKai.Enabled = True
            Me.pnlJizenAzu.Enabled = True
            Me.pnlJizenHurikae.Enabled = True
            Me.pnlJizenSo.Enabled = True
            '事後
            Me.Label304.Enabled = True
            Me.Label303.Enabled = True
            Me.grpJigoDonyuji.Enabled = True
            '補助
            Me.pnlBunkatumisyu.Enabled = True

            '--------------------------
            '関連する作業項目制御
            '--------------------------
            Select Case True
                Case Me.chkOptionSelectKy.Checked = False
                    '事前
                    Me.pnlJizenKai.Enabled = False                          '解約未処理データ処理
                    Me.pnlJizenAzu.Enabled = False                          '預り金データ処理
                    Me.pnlJizenHurikae.Enabled = False                      '口座振替入金処理
                    Me.pnlJizenSo.Enabled = False                           '送金処理
                    '事後
                    Me.Label304.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.Label303.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.grpJigoDonyuji.Enabled = False                            '賃貸革命導入時通常作業
                    '補助
                    Me.pnlBunkatumisyu.Enabled = False                      '分割入金未収分
                Case Me.chkOptionSelectSq.Checked = False
                    '事前
                    Me.pnlJizenHurikae.Enabled = False                      '口座振替入金処理
                    Me.pnlJizenSo.Enabled = False                           '送金処理
                    '事後
                    Me.Label304.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.Label303.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.grpJigoDonyuji.Enabled = False                            '賃貸革命導入時通常作業
                    '補助
                    Me.pnlBunkatumisyu.Enabled = False                      '分割入金未収分
                Case Me.chkOptionSelectNk.Checked = False
                    '事前
                    Me.pnlJizenHurikae.Enabled = False                      '口座振替入金処理
                    Me.pnlJizenSo.Enabled = False                           '送金処理
                    '事後
                    Me.Label304.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.Label303.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.grpJigoDonyuji.Enabled = False                            '賃貸革命導入時通常作業
                    '補助
                    Me.pnlBunkatumisyu.Enabled = False                      '分割入金未収分
                Case Me.chkOptionSelectSh.Checked = False
                    '事前
                    Me.pnlJizenSo.Enabled = False                           '送金処理
                    '事後
                    Me.Label304.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.Label303.Enabled = chkOptionSelectKy.Checked         '年間収支系帳票について
                    Me.grpJigoDonyuji.Enabled = False                            '賃貸革命導入時通常作業
                    '補助
                    Me.pnlBunkatumisyu.Enabled = False                      '分割入金未収分
                Case Me.chkOptionSelectFB.Checked = False
                    Me.chkJizenHurikae.Enabled = chkOptionSelectFB.Checked  '口座振替入金処理
                    Me.Label88.Enabled = chkOptionSelectFB.Checked          '口座振替入金処理
            End Select

        End Sub

        ''' <summary>
        ''' イベント処理：オプションチェックボックスON/OFFと各種作業のグレイアウトの連動
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・クレーム、修繕
        ''' </remarks>
        Private Sub chkOptionSelect_clsz_CheckedChanged(sender As Object, e As EventArgs) Handles chkOptionSelectClaim.CheckedChanged, chkOptionSelectReform.CheckedChanged

            '画面起動時は処理を抜ける
            If frmloadflg_clsz Then
                Exit Sub
            End If

            '--------------------------
            'グレイアウトの初期化
            '--------------------------
            '事前
            Me.pnlSzenKysSorit.Enabled = True
            '事後
            Me.Label304.Enabled = True
            Me.Label303.Enabled = True
            Me.grpJigoDonyuji.Enabled = True
            '補助
            pnlRelRename.Enabled = True
            Panel25.Enabled = True
            For Each ctrl In pnlRelRename.Controls
                ctrl.Enabled = True
            Next
            For Each ctrl In Panel25.Controls
                ctrl.Enabled = True
            Next

            '--------------------------
            '関連する作業項目制御
            '--------------------------
            Select Case True
                Case Me.chkOptionSelectClaim.Checked = False
                    '事後
                    If Me.chkOptionSelectReform.Checked = False Then        '関連ファイルリネーム
                        '修繕もチェックOFFの場合は枠からグレイアウト
                        Me.pnlRelRename.Enabled = False
                    Else
                        'クレームのみチェックOFFの場合はクレームに関する箇所をグレイアウト
                        Me.Label217.Enabled = False
                        Me.lblRelRenameClaim.Enabled = False
                        Me.Label400.Enabled = False
                        Me.chkRelRenameClaim.Enabled = False
                    End If
                Case Me.chkOptionSelectReform.Checked = False
                    '事前
                    Me.pnlSzenKysSorit.Enabled = False                      '修繕対象負担者設定契約者データの確認
                    '事後
                    If Me.chkOptionSelectClaim.Checked = False Then         '関連ファイルリネーム
                        'クレームもチェックOFFの場合は枠からグレイアウト
                        Me.pnlRelRename.Enabled = False
                    Else
                        '修繕のみチェックOFFの場合は修繕に関する箇所をグレイアウト
                        Me.Label221.Enabled = False
                        Me.lblRelRenameSzen.Enabled = False
                        Me.Label229.Enabled = False
                        Me.chkRelRenameSzen.Enabled = False
                    End If
            End Select

        End Sub

        ''' <summary>
        ''' イベント処理：関連ファイル変換後文字列入力箇所 '20160720 リネーム機能の修正
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub txtAfRelRename_LostFocus(sender As Object, e As EventArgs) Handles txtAfRelRename.LostFocus

            Dim afrelname As String = Me.txtAfRelRename.Text

            If EtcMethod.Chk_DirExist(afrelname) Then
                Me.btnRNRelRename.Enabled = True
                Me.Label242.Visible = False
            Else
                Me.btnRNRelRename.Enabled = False
                Me.Label242.Visible = True
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：運用開始入力制御 '20160707 運用開始年月入力制御処理の追加
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub txtUnyoYYYYMM_LostFocus(sender As Object, e As EventArgs) Handles txtUnyoYYYYMM.LostFocus

            Dim defdate As String = Now.ToString("yyyy/MM")
            Dim tmp_str As String = Me.txtUnyoYYYYMM.Text
            Dim tmp_date As String = tmp_str & "/01"
            Dim chg_date As Date

            '入力形式チェック
            If Date.TryParse(tmp_date, chg_date) Then
                '"yyyy/MM"形式に変換して再表示
                Me.txtUnyoYYYYMM.Text = chg_date.ToString("yyyy/MM")
            Else
                'コンバート実施年月を設定
                Me.txtUnyoYYYYMM.Text = defdate
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：中間ファイルオープン処理
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnHJizenTyukanOpenSyori_Click(sender As Object, e As EventArgs) Handles btnHJizenTyukanOpen.Click, btnHJizenTyukanOpen2.Click

            '20160829 中間ファイル有無確認処理を追加 -chg sta
            'System.Diagnostics.Process.Start(Me.txtMidDirPath.Text)
            Dim flg As Boolean = True
            Dim basemidfilepath As String = Me.txtMidDirPath.Text

            flg = EtcMethod.Chk_FileExist(basemidfilepath)

            If flg = False Then
                MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            System.Diagnostics.Process.Start(Me.txtMidDirPath.Text)
            '20160829 中間ファイル有無確認処理を追加 -chg end

        End Sub

        ''' <summary>
        ''' イベント処理：既存用→汎用用中間ファイルへのコピー処理(開発用処理)　　　　　　
        ''' 20160912 既存用→汎用用中間ファイルへのコピー処理
        ''' (開発用処理)
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnExistMidToBaseMid_Click(sender As Object, e As EventArgs) Handles btnExistMidToBaseMid.Click

            Dim tmp_cnt As Integer = 0
            Dim list_cv As New List(Of String)                                          '移行項目取得
            list_cv = Me.Get_ListCVChkitem(True)

            '移行項目チェック
            If list_cv.Count = 0 Then
                MessageBox.Show("移行項目を選択して下さい。")
                Exit Sub
            End If

            '格納先フォルダチェック
            If EtcMethod.Chk_DirExist(Me.txtExistMidToBaseMid.Text) = False Then
                MessageBox.Show("中間ファイル格納先フォルダが存在しません。再度選択して下さい。")
                Exit Sub
            End If

            MsgResult = MessageBox.Show("既存→汎用中間ファイルへのコピー処理を行います。よろしいですか？", "確認", _
                                            MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                Exit Sub
            End If

            '---------------------------------------------------------------
            '汎用用中間ファイルと既存用中間ファイルの照合用仮テーブル作成
            '---------------------------------------------------------------
            '仮テーブル初期化
            Dim midfileinfo_dropqry As String = DBQuery.Qry_DropInfo(MIDFILEINFO_DBNAME, True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, midfileinfo_dropqry, tmp_cnt)

            '仮テーブル新規作成
            Dim midfileinfo_createqry As String = MidFileInfoModule.Get_MidFileInfo_CreateQry()
            DBExec.Exec_NonQuery(Me.sqlcnnv10, midfileinfo_createqry, tmp_cnt)

            '挿入
            Dim normalflg As Boolean = True
            Dim midfileinfo_insertqry As String = MidFileInfoModule.Get_MidFileInfo_InsertQry()
            If midfileinfo_insertqry <> "" Then
                normalflg = DBExec.Exec_NonQuery(Me.sqlcnnv10, midfileinfo_insertqry, tmp_cnt)
            End If
            If normalflg = False Then
                MessageBox.Show("仮テーブル作成中にエラーが発生しました。")
                Exit Sub
            End If

            '-----------------------------------------------------------------------
            '作成した照合用仮テーブルからヘッダーに関する情報を取得してオブジェクトへ格納
            '-----------------------------------------------------------------------
            Call Me.Set_MidHeaderInfoToObj()

            '-----------------------------------------------------------------------
            '仮テーブルから紐付情報を取得して汎用用中間ファイルを作成する
            '-----------------------------------------------------------------------
            Dim errstr As String = ""
            normalflg = Me.Set_ExistMidFile_Copy(errstr, list_cv)
            If normalflg Then
                MessageBox.Show("既存用→汎用用中間ファイルへのコピー処理が完了しました。")
            Else
                MessageBox.Show("コピー処理中にエラーが発生しました。")
            End If

        End Sub

#End Region

#Region "接続情報設定"

        ''' <summary>
        ''' 接続情報：初期値→コントロール
        ''' </summary>
        ''' <param name="cnnv7"></param>
        ''' <param name="cnnv10"></param>
        ''' <remarks></remarks>
        Public Sub Set_DefConInfo_To_Control(ByRef cnnv7 As Njc.Model.DefSQLConnection, ByRef cnnv10 As Njc.Model.DefSQLConnection)

            'V7
            Me.txtV7Server.Text = cnnv7.ServerName
            Me.txtV7Catalog.Text = cnnv7.InitialCatalog
            'Me.txtV7Networklib.Text = cnnv7.NetworkLibrary
            Me.txtV7User.Text = cnnv7.User
            Me.txtV7Pass.Text = cnnv7.Pass

            '10
            Me.txtV10Server.Text = cnnv10.ServerName
            Me.txtV10Catalog.Text = cnnv10.InitialCatalog
            'Me.txtV10Networklib.Text = cnnv10.NetworkLibrary
            Me.txtV10User.Text = cnnv10.User
            Me.txtV10Pass.Text = cnnv10.Pass

            '共通
            Me.txtTimeOut.Text = cnnv10.TimeOut

        End Sub

        ''' <summary>
        ''' 接続情報：コントロール→モデル
        ''' </summary>
        ''' <param name="model"></param>
        ''' <param name="flg"></param>
        ''' <remarks></remarks>
        Public Sub Set_Control_To_ConModel(ByRef model As Njc.Model.DefSQLConnection, Optional ByVal flg As Integer = 0)

            If flg = 1 Then
                '10
                model.ServerName = Me.txtV10Server.Text
                model.InitialCatalog = Me.txtV10Catalog.Text
                'model.NetworkLibrary = Me.txtV10Networklib.Text
                model.User = Me.txtV10User.Text
                model.Pass = Me.txtV10Pass.Text
            Else
                'V7
                model.ServerName = Me.txtV7Server.Text
                model.InitialCatalog = Me.txtV7Catalog.Text
                'model.NetworkLibrary = Me.txtV7Networklib.Text
                model.User = Me.txtV7User.Text
                model.Pass = Me.txtV7Pass.Text
            End If

            '共通
            model.TimeOut = Me.txtTimeOut.Text

        End Sub

        ''' <summary>
        ''' DB接続情報取得 (2回目以降起動時)  
        ''' </summary>
        ''' <remarks></remarks>
        Private Function Set_DBExistConInfo() As Boolean

            Dim rtn As Boolean = True

            '保管されている接続情報ファイルの読込
            Dim exepath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim exedir As String = IO.Path.GetDirectoryName(exepath)
            Dim coninfodirpath As String = EtcMethod.Set_Path(exedir, DIR_INI_NAME)
            Dim coninfofilepathV7 As String = EtcMethod.Set_Path(coninfodirpath, FILE_V7_CONNAME)
            Dim coninfofilepath10 As String = EtcMethod.Set_Path(coninfodirpath, FILE_10_CONNAME)

            'どちらかのファイルが存在しない(初回起動またはファイル名が編集されている)場合
            If EtcMethod.Chk_FileExist(coninfofilepathV7) = False OrElse EtcMethod.Chk_FileExist(coninfofilepath10) = False Then
                rtn = False
                Return rtn
            End If

            Dim xmlreader As System.Xml.XmlReader
            Dim item As String = ""
            Dim data As String = ""
            Dim svname As String = ""
            Dim dbname As String = ""
            Dim username As String = ""
            Dim password As String = ""
            Dim encryptedpassword As String = ""
            Dim passkey As String = "tRwmj5U4"
            Dim conmodel As Object = Nothing

            For cntii = 0 To 1

                Select Case cntii
                    Case 0
                        xmlreader = System.Xml.XmlReader.Create(coninfofilepathV7)
                        conmodel = fstmodelv7
                    Case 1
                        xmlreader = System.Xml.XmlReader.Create(coninfofilepath10)
                        conmodel = fstmodelv10
                End Select

                While xmlreader.Read

                    If xmlreader.NodeType = Xml.XmlNodeType.Element Then

                        'データ取得
                        item = xmlreader.LocalName
                        data = xmlreader.ReadString

                        'それぞれの要素で分岐
                        Select Case item
                            Case "ServerName"
                                svname = data
                            Case "InitialCatalog"
                                dbname = data
                            Case "UserID"
                                username = data
                            Case "Password"
                                'password = data
                            Case "EncryptedPassword"
                                encryptedpassword = data
                        End Select

                    End If

                End While

                'パスワード複合化
                password = Me.Decrypt(encryptedpassword, passkey)

                With conmodel
                    .ServerName = svname
                    .InitialCatalog = dbname
                    .User = username
                    .Pass = password
                End With

                Select Case cntii
                    Case 0
                        fstmodelv7 = conmodel
                    Case 1
                        fstmodelv10 = conmodel
                End Select

                xmlreader.Close()

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' DB接続情報作成 (2回目以降起動時用に退避させておく)
        ''' </summary>
        ''' <remarks>
        ''' ・接続情報パスワード保存処理
        ''' </remarks>
        Private Function MakeFile_DBConInfo(ByVal conmodel As Object, ByVal typeno As Integer) As Boolean

            Dim svname As String = conmodel.ServerName
            Dim catalog As String = conmodel.InitialCatalog
            Dim user As String = conmodel.User
            Dim pass As String = conmodel.Pass

            '接続情報保存ファイル格納先取得
            Dim exepath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim exedir As String = IO.Path.GetDirectoryName(exepath)
            Dim coninfodirpath As String = EtcMethod.Set_Path(exedir, DIR_INI_NAME)
            Dim coninfofilepath As String = ""

            Select Case typeno
                Case 0
                    coninfofilepath = EtcMethod.Set_Path(coninfodirpath, FILE_V7_CONNAME)
                Case 1
                    coninfofilepath = EtcMethod.Set_Path(coninfodirpath, FILE_10_CONNAME)
            End Select

            'Windows認証時にUserにデフォルト値を入れて保存する処理を追加
            '※空のままだとXML作成の際に改行が混入するためこれを防ぐ
            If user.Trim = "" Then
                user = "sa"
            End If

            Dim passkey As String = "tRwmj5U4"
            Dim encryptedpassword As String = Me.Encrypt(pass, passkey)

            Dim strXml As String = "<?xml version='1.0'?>" & _
                                   "<coninfo>" & _
                                   "<!--サーバー名、カタログ名、ユーザー名、パスワード-->" & _
                                   "<ServerName>" & svname & "</ServerName>" & _
                                   "<InitialCatalog>" & catalog & "</InitialCatalog>" & _
                                   "<UserID>" & user & "</UserID>" & _
                                   "<Password />" & _
                                   "<EncryptedPassword>" & encryptedpassword & "</EncryptedPassword>" & _
                                   "</coninfo>"

            Dim xmlDoc As New System.Xml.XmlDocument

            '文字列からDOMドキュメントを生成
            xmlDoc.LoadXml(strXml)

            '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
            'Try
            '    '作成したDOMドキュメントをファイルに保存
            '    xmlDoc.Save(coninfofilepath)
            'Catch ex As System.Xml.XmlException
            '    'XMLによる例外をキャッチ
            '    Console.WriteLine(ex.Message)
            'Catch ex As Exception
            '    'その他の例外をキャッチ
            '    Console.WriteLine(ex.Message)
            'End Try

            Try
                '作成したDOMドキュメントをファイルに保存
                xmlDoc.Save(coninfofilepath)

                'アクセス拒否
                'Catch ex As UnauthorizedAccessException
                '    MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                '                                "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
                '                                "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '    Return False
                'その他の例外(アクセス拒否を含む)
            Catch ex As Exception
                MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                                            "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False

            End Try

            Return True
            '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end
        End Function

        ''' <summary>
        ''' V7DB接続情報取得 (初期値取得) 
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_V7DBConInfo()

            'V7接続情報を読み取り専用で取得
            Dim regkey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\VB and VBA Program Settings\Fkanri50\Fk5Db", False)

            Dim svname As String = ""
            Dim dbname As String = ""
            Dim username As String = ""
            Dim chgpassword As String = ""
            Dim password As String = ""

            If regkey Is Nothing Then
                'レジストリ情報が存在しない場合は空文字で出力
            Else
                '値を取得 (存在しない場合は空文字)
                svname = DirectCast(regkey.GetValue("DATA SOURCE", ""), String)
                dbname = DirectCast(regkey.GetValue("INITIAL CATALOG", ""), String)
                username = DirectCast(regkey.GetValue("UID", ""), String)
                chgpassword = DirectCast(regkey.GetValue("PWD", ""), String)
                'パスワード変換
                password = Me.Get_V7DBPassword(chgpassword)
            End If

            'モデルへ格納
            With fstmodelv7
                .ServerName = svname
                .InitialCatalog = dbname
                .User = username
                .Pass = password
            End With

        End Sub

        ''' <summary>
        ''' パスワード取得 (V7プログラムから引用) 
        ''' </summary>
        ''' <param name="pass"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_V7DBPassword(ByVal pass As String) As String

            Dim rtn_str As String

            Dim PSW As String = "ogaryo"
            Dim tmp_pass As String = ""
            Dim n1 As Integer = 0
            Dim n2 As Integer = 0
            Dim n3 As Integer = 0

            '制御番号を形成
            For i = 1 To Len(PSW)
                n1 = n1 + Asc(Mid(PSW, i, 1))
                n1 = (n1 * 367 + 331) Mod &HFFF
                n2 = ((n2 + n1) * 743 + 599) Mod &HFFF
                n3 = ((n3 + n2) * 563 + 787) Mod &HFFF
            Next i

            '変換
            tmp_pass = pass
            Chg_Password(tmp_pass, n1, n2, n3)
            rtn_str = tmp_pass

            Return rtn_str

        End Function

        ''' <summary>
        ''' パスワード変換 (V7プログラムから引用)  
        ''' </summary>
        ''' <param name="str"></param>
        ''' <param name="int1"></param>
        ''' <param name="int2"></param>
        ''' <param name="int3"></param>
        ''' <remarks></remarks>
        Private Sub Chg_Password(ByRef str As String, ByVal int1 As Integer, ByVal int2 As Integer, ByVal int3 As Integer)

            Dim R As Integer
            Dim M As Integer
            Dim N As Integer
            Dim BigNum As Integer = 32768       '疑似乱数の周期(2の巾乗）
            Dim i As Integer
            Dim c As Integer
            Dim d As Integer

            R = int1                            'オプションの乱数出発点
            M = (int2 * 4 + 1) Mod BigNum       'オプションの乗数初期化値
            N = (int3 * 2 + 1) Mod BigNum       'オプションの付加項初期化値

            ' 文字列を処理します
            For i = 1 To Len(str)
                c = Asc(Mid(str, i, 1))
                ' プリント可能な文字の一部分だけを変更します
                Select Case c
                    Case 48 To 57
                        d = c - 48
                    Case 63 To 90
                        d = c - 53
                    Case 97 To 122
                        d = c - 59
                    Case Else
                        d = -1
                End Select
                ' １文字を暗号化する準備をします
                If d >= 0 Then
                    ' 次の疑似乱数を生成します
                    R = (R * M + N) Mod BigNum
                    ' 文字と疑似乱数のビットをXorします
                    d = (R And 63) Xor d
                    ' 文字をプリント可能な範囲に戻します
                    Select Case d
                        Case 0 To 9
                            c = d + 48
                        Case 10 To 37
                            c = d + 53
                        Case 38 To 63
                            c = d + 59
                    End Select
                    ' 元の文字を結果で置き換えます
                    Mid(str, i, 1) = Chr(c)
                End If
            Next i
        End Sub

        ''' <summary>
        ''' 10DB接続情報取得 
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_10DBConInfo()

            Dim xmlfilepath As String = ""

            'XMLファイルパスを取得
            Dim tmp_filepath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            tmp_filepath = EtcMethod.Set_Path(tmp_filepath, COMPANY_NAME)
            tmp_filepath = EtcMethod.Set_Path(tmp_filepath, PRODUCT_NAME)
            tmp_filepath = EtcMethod.Set_Path(tmp_filepath, String.Format("{0}.sql.Main.xml", PRODUCT_NAME))
            xmlfilepath = tmp_filepath

            'ファイル有無チェック
            If EtcMethod.Chk_FileExist(xmlfilepath) = False Then
                Exit Sub
            End If

            'ファイル読込
            Dim xmlreader As System.Xml.XmlReader = System.Xml.XmlReader.Create(xmlfilepath)
            Dim item As String = ""
            Dim data As String = ""
            Dim svname As String = ""
            Dim dbname As String = ""
            Dim username As String = ""
            Dim password As String = ""
            Dim encryptedpassword As String = ""
            Dim passkey As String = "tRwmj5U4"
            Dim tmp_pass As String = ""

            While xmlreader.Read
                If xmlreader.NodeType = Xml.XmlNodeType.Element Then

                    'データ取得
                    item = xmlreader.LocalName
                    data = xmlreader.ReadString

                    'それぞれの要素で分岐
                    Select Case item
                        Case "ServerName"
                            svname = data
                        Case "InitialCatalog"
                            dbname = data
                        Case "UserID"
                            username = data
                        Case "EncryptedPassword"
                            encryptedpassword = data
                    End Select
                End If
            End While

            'パスワード複合化
            password = Me.Decrypt(encryptedpassword, passkey)

            'モデルへ格納
            With fstmodelv10
                .ServerName = svname
                .InitialCatalog = dbname
                .User = username
                .Pass = password
            End With

            xmlreader.Close()

        End Sub

        ''' <summary>
        ''' DESで暗号化します。<br/> 
        ''' </summary>
        Public Function Encrypt(ByVal target As String, ByVal key As String) As String
            Dim targetBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(target)

            Dim des As New Security.Cryptography.DESCryptoServiceProvider
            Dim keyBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(key)
            des.Key = resizeBytesArray(keyBytes, des.Key.Length)
            des.IV = resizeBytesArray(keyBytes, des.IV.Length)

            Dim outStream As New IO.MemoryStream
            Dim desencrypt As Security.Cryptography.ICryptoTransform = des.CreateEncryptor
            Dim cryptStream As New Security.Cryptography.CryptoStream(outStream, desencrypt, Security.Cryptography.CryptoStreamMode.Write)

            cryptStream.Write(targetBytes, 0, targetBytes.Length)
            cryptStream.FlushFinalBlock()
            Dim outBytes As Byte() = outStream.ToArray

            cryptStream.Close()
            outStream.Close()

            Return Convert.ToBase64String(outBytes)
        End Function

        ''' <summary>
        ''' DESで復号化します。<br/> 
        ''' </summary>
        Public Function Decrypt(ByVal encryptedTarget As String, ByVal key As String) As String
            Dim encryptedBytes As Byte() = Convert.FromBase64String(encryptedTarget)

            Dim des As New Security.Cryptography.DESCryptoServiceProvider
            Dim keyBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(key)
            des.Key = resizeBytesArray(keyBytes, des.Key.Length)
            des.IV = resizeBytesArray(keyBytes, des.IV.Length)

            Dim inStream As New IO.MemoryStream(encryptedBytes)
            Dim desdecrypt As Security.Cryptography.ICryptoTransform = des.CreateDecryptor
            Dim cryptStream As New Security.Cryptography.CryptoStream(inStream, desdecrypt, Security.Cryptography.CryptoStreamMode.Read)

            Dim outStream As New IO.StreamReader(cryptStream, System.Text.Encoding.UTF8)
            Dim result As String = outStream.ReadToEnd

            outStream.Close()
            cryptStream.Close()
            inStream.Close()

            Return result
        End Function

        Private Function resizeBytesArray(ByVal bytes() As Byte, _
                                ByVal newSize As Integer) As Byte()
            Dim newBytes(newSize - 1) As Byte
            If bytes.Length <= newSize Then
                Dim i As Integer
                For i = 0 To bytes.Length - 1
                    newBytes(i) = bytes(i)
                Next i
            Else
                Dim pos As Integer = 0
                Dim i As Integer
                For i = 0 To bytes.Length - 1
                    newBytes(pos) = newBytes(pos) Xor bytes(i)
                    pos += 1
                    If pos >= newBytes.Length Then
                        pos = 0
                    End If
                Next i
            End If
            Return newBytes
        End Function

#End Region

#Region "コンバート実行処理(既存ユーザ用, 汎用共通)"

        ''' <summary>
        ''' コンバート処理 
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Cnv_Main(list_cv As List(Of String))

            Dim rowcnt As Integer = 0                                       '確認メモ：Exec_NonQueryの第3引数の為用意(実際は未使用)
            Dim obj_pgb As New ProgressBarManager
            Dim errstr As String = ""
            Dim normalflg As Boolean = True

            '--------------------------------------------------
            ' ログ状況初期設定・出力
            '--------------------------------------------------
            '初期設定
            Call Me.Set_LogInit(True)
            Call Me.Set_LogInit(False)
            'ログ出力
            Dim tmp_sql_sta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_CV_STA), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_sta, rowcnt)
            '状況出力
            '20161018 UI最終確認での修正 -add sta
            '進捗状況出力内容初期化
            Me.txtTotalSituation.Text = ""
            Me.txtPartialSituation.Text = ""
            '20161018 UI最終確認での修正 -add end
            Call Me.Set_Situation(SITUATION_CV_STA, 0)
            '20161017 進捗表示ラベル初期表示修正 -add sta
            Me.lblTotalSituation.Text = SITUATION_CV_BFRUN  '全体進捗ラベル初期化
            '20161104 2回目実行時の画面初期化処理修正 -add
            Me.lblCVItem.Text = ""                          '個別進捗ラベル初期化
            Call obj_com.ProgressOutPut(0, 1, True)         '全体進捗進捗率初期化
            Call obj_com.ProgressOutPut(0, 1)               '個別進捗進捗率初期化
            '20161017 進捗表示ラベル初期表示修正 -add end
            '--------------------------------------------------
            ' 設定値取得
            '--------------------------------------------------             
            Dim midfileflg As Boolean = Me.chkMiddleFile.Checked            '中間ファイル作成フラグ(開発用)
            Dim himoflg As Boolean = Me.chkRelation.Checked                 '紐付作業実行フラグ
            Dim tmp_cntstr As String = Me.txtLogOutputCnt.Text.Trim         'ログ出力件数設定値取得
            Dim tmp_cntint As Integer = 0                                   'ログ出力件数(作業用)
            If Not (Int32.TryParse(tmp_cntstr, tmp_cntint)) Then
                tmp_cntint = 1
            End If
            Log_OutputCnt = tmp_cntint


            '--------------------------------------------------
            ' ハッシュテーブル初期化 '20160711 オブジェクトインスタンス化の修正
            '--------------------------------------------------
            Call Me.Ini_HashTable()


            '================================================== メイン処理 ================================================== s
            '--------------------------------------------------
            ' 中間ファイル処理
            '--------------------------------------------------
            Call Chg_LeftProgress("dcfilewrite", 1)
            pnlMenuDatacv.Refresh()       'kekeke

            '----- 全体用・個別用プログレスバー初期化 -----
            Call obj_pgb.pgbInitPart(0)
            Call obj_pgb.pgbInitTotal(0)
            Me.tabPageJikko.Refresh()
            pgbtotalcnt = list_cv.Count
            Call obj_pgb.pgbInitTotal(pgbtotalcnt)

            If CNVNO = ConvertTypes._汎用 Then

                '汎用用中間ファイルから既存用中間ファイルへのコピー処理
                normalflg = Me.Set_BaseMidFileToExistingMidFile(list_cv, errstr)
                If normalflg = False Then
                    MsgResult = MessageBox.Show(errstr, "注意", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    GoTo skiplabel
                End If

            ElseIf CNVNO = ConvertTypes._既存ユーザ用 And midfileflg Then

                Me.lblTotalSituation.Text = SITUATION_MID_RUN

                '中間ファイル作成処理
                normalflg = Me.Set_MiddleFile(list_cv, errstr)              '確認メモ：Contorl."grp"-"chkbox"

                '処理続行確認メッセージ
                If normalflg = False Then
                    MsgResult = MessageBox.Show(errstr, "注意", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    GoTo skiplabel
                ElseIf Me.chkMidNotStop.Checked = False And CancelFlg = False Then
                    MsgResult = MessageBox.Show(MSG_STA_B, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.No Then
                        Dim tmp_sql_midend As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, LOG_SYORIKOMK_CV_END), False)
                        DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_midend, rowcnt)
                        Exit Sub
                    End If
                End If
            End If


            '--------------------------------------------------
            ' DB書込実行前準備 (作業用ファイルの仮TBL作成)
            '--------------------------------------------------                               
            If CancelFlg = False Then
                normalflg = Me.Before_DBWrite(list_cv)
                If normalflg = False Then
                    MsgResult = MessageBox.Show(MSG_ERR_MAKE_CVDB, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    Exit Sub
                End If
            End If


            '--------------------------------------------------
            ' 中間ファイルチェック
            '--------------------------------------------------
            If CancelFlg = False And CNVNO = ConvertTypes._汎用 Then
                normalflg = Me.Chk_MidFileChk_Main(list_cv)
                If normalflg = False Then
                    MsgResult = MessageBox.Show(MSG_ERR_CHK_MID, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    Exit Sub
                End If
            End If


            '--------------------------------------------------
            ' 紐付設定画面呼出
            '--------------------------------------------------
            Call Chg_LeftProgress("dcrelation", 1)
            pnlMenuDatacv.Refresh()       'kekeke

            If CancelFlg = False And himoflg Then
                '選択された移行項目に関連する紐付情報を起動させるために移行項目のチェック値を引数として渡す
                Dim cvitemtorel As String = Me.Set_CVItemChkToCmdLine(list_cv)
                '20160720 紐付設定ファイル出力機能の追加 -add
                relitemstr = cvitemtorel.Trim
                '移行項目と関連する紐付項目が存在する場合は紐付ツールを起動
                If cvitemtorel.Trim <> "" Then
                    '20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg sta
                    'Call Me.Rel_ExeCall(cvitemtorel)

                    '----- 全体用・個別用プログレスバー初期化 -----
                    Call obj_pgb.pgbInitPart(0)
                    Call obj_pgb.pgbInitTotal(0)
                    Call obj_com.ProgressOutPut(0, pgbtotalcnt)
                    Call obj_com.ProgressOutPut(0, pgbtotalcnt, True)
                    Me.tabPageJikko.Refresh()
                    pgbtotalcnt = 1
                    Call obj_pgb.pgbInitPart(pgbtotalcnt)
                    Call obj_pgb.pgbInitTotal(pgbtotalcnt)
                    Me.lblTotalSituation.Text = SITUATION_REL_RUN

                    '◎紐付設定画面呼出
                    CancelFlg = Me.Rel_ExeCall(cvitemtorel)
                    '20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg end
                End If
            End If


            '--------------------------------------------------
            ' DB書込処理
            '--------------------------------------------------
            Call Chg_LeftProgress("dcconv", 1)
            pnlMenuDatacv.Refresh()       'kekeke
            If CancelFlg = False Then

                '----- 全体用・個別用プログレスバー初期化 -----
                Call obj_pgb.pgbInitPart(0)
                Call obj_pgb.pgbInitTotal(0)
                Call obj_com.ProgressOutPut(0, pgbtotalcnt)
                Call obj_com.ProgressOutPut(0, pgbtotalcnt, True)
                Me.tabPageJikko.Refresh()
                pgbtotalcnt = list_cv.Count
                Call obj_pgb.pgbInitTotal(pgbtotalcnt)
                Me.lblTotalSituation.Text = SITUATION_CV_RUN

                '◎DB書き込み処理
                Call Me.Read_MidFile_WriteDB(list_cv)
            End If


            '20160929 データ調整用メソッドの作成 -add sta
            '--------------------------------------------------
            ' データ調整 
            '--------------------------------------------------
            If CancelFlg = False Then
                '◎データ調整処理
                Call Me.DataCond()
            End If
            '20160929 データ調整用メソッドの作成 -add end

            '================================================== メイン処理 ================================================== e


            '--------------------------------------------------
            ' 終了処理
            '--------------------------------------------------
skiplabel:

            'ログ出力
            Dim tmp_sql_end As String = ""
            If CancelFlg Then
                tmp_sql_end = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, LOG_SYORIKOMK_CV_STOP), False)
            Else
                tmp_sql_end = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, LOG_SYORIKOMK_CV_END), False)
            End If
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_end, rowcnt)

            '状況出力
            Dim tmp_situationstr As String = ""
            If CancelFlg Then
                tmp_situationstr = SITUATION_STOP
            Else
                tmp_situationstr = SITUATION_CV_END
            End If

            Call Me.Set_Situation(tmp_situationstr, 0)
            Me.lblTotalSituation.Text = tmp_situationstr

        End Sub

        ''' <summary>
        ''' DB書込処理 
        ''' </summary>
        ''' <param name="list_cv"></param>
        ''' <remarks></remarks>
        Private Sub Read_MidFile_WriteDB(ByVal list_cv As List(Of String))

            Dim pgbcnt As Integer = 0
            Dim tmp_totalsituation As String = ""
            Dim tmp_partsituation As String = ""
            Dim obj_pgb As New ProgressBarManager
            Dim rowcnt As Integer = 0


            '--------------------------------------------------
            ' ログ/状況出力
            '--------------------------------------------------
            'ログ出力
            Dim tmp_sql_sta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_DBCVSTA), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_sta, rowcnt)

            '状況出力
            tmp_partsituation = SITUATION_WRITESTA
            Call Me.Set_Situation(tmp_partsituation, 1)


            '--------------------------------------------------
            ' 設定値取得
            '--------------------------------------------------
            InitDBFlg = Me.optCVNew.Checked                 '新規コンバートフラグ
            OverWriteDBFlg = Me.chkOverWrite.Checked        '上書きコンバートフラグ


            '--------------------------------------------------
            ' コンバート実行前準備
            '--------------------------------------------------
            Call Me.Setting_BeforeCV()


            '--------------------------------------------------
            ' 紐付項目の挿入処理実行
            '--------------------------------------------------
            Dim relcvnormalflg As Boolean = Me.Set_RelItem()
            If relcvnormalflg = False Then
                MsgResult = MessageBox.Show(MSG_ERR_RELCV, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Exit Sub
            End If

            '20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -add sta
            '--------------------------------------------------
            ' ダミー用の口座情報作成
            '--------------------------------------------------
            If CNVNO = ConvertTypes._汎用 Then
                If Me.Set_DummyBankData() = False Then
                    MsgResult = MessageBox.Show(MSG_ERR_DUMMYBANKCV, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    Exit Sub
                End If
            End If
            '20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -add end

            '==================================================
            'メイン処理 (処理開始)
            '==================================================
            For Each cvitem In list_cv

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Exit For
                End If

                'グループボックス/チェックボックス名取得
                Dim tmp_arry() As String = cvitem.Split("-")
                Dim syorigrp As String = tmp_arry(0)
                Dim syoriitem As String = tmp_arry(1)

                '処理Repositoryの生成 (DB登録用)
                Dim obj_rep As Object = Get_ObjRep_DB(syoriitem)

                '----- デバッグ用処理 ----- sta
                'Dim list As New List(Of String) From {"送信設定athome情報"}
                'Dim list As New List(Of String) From {"送信設定自社web情報", "送信設定HOMES情報", "送信設定athome情報", "送信設定SUUMO情報"}
                'Dim list As New List(Of String) From {"送金ルール基本情報", "送金ルール送金先情報", "送金ルール入金項目情報", "送金ルール控除項目情報"}
                ''Dim list As New List(Of String) From {"広告補足自社web情報", "広告補足HOMES情報", "広告補足athome情報", "広告補足SUUMO情報"}
                'Dim list As New List(Of String) From {"送金ルール控除項目情報"}
                'If list.Contains(syoriitem) = False Then
                '    obj_rep = Nothing
                'End If
                '----- デバッグ用処理 ----- end

                '処理開始
                If obj_rep IsNot Nothing Then
                    '状況出力
                    tmp_totalsituation = syoriitem & INDENT_1 & SITUATION_WRITESTA
                    Call Me.Set_Situation(tmp_totalsituation, 0)
                    Me.lblCVItem.Text = syorigrp & " ： " & syoriitem

                    'テーブル初期化
                    If InitDBFlg Then
                        Dim sql_where As String = IIf(hash_deltblqrywhere(syoriitem) = Nothing, "", hash_deltblqrywhere(syoriitem))
                        '20160720 連動情報構築 -chg sta
                        'If syorigrp <> "初期設定" Then
                        '    Me.Initialize_Table(Hash_TblName_JpToAlpha(syoriitem), sql_where)
                        'End If
                        Select Case syorigrp
                            Case "初期設定"
                                '初期化しない
                            Case "物件データ連動情報"
                                Dim tmp_chgsyoriitem As String = syoriitem
                                If syoriitem = "広告補足自社web情報" Or syoriitem = "広告補足HOMES情報" Or syoriitem = "広告補足athome情報" Or syoriitem = "広告補足SUUMO情報" Then
                                    tmp_chgsyoriitem = "広告補足情報"
                                ElseIf syoriitem = "送信設定自社web情報" Or syoriitem = "送信設定HOMES情報" Or syoriitem = "送信設定athome情報" Or syoriitem = "送信設定SUUMO情報" Then
                                    tmp_chgsyoriitem = "送信設定情報"
                                End If
                                Me.Initialize_Table(Hash_TblName_JpToAlpha(tmp_chgsyoriitem), sql_where)
                                '20161028 物件/部屋鍵取得方法修正 -add sta
                            Case "各マスタ情報"
                                If syoriitem = "鍵タイトルマスタ" Then
                                    Select Case CNVNO
                                        Case ConvertTypes._汎用
                                            '初期化しない
                                        Case ConvertTypes._既存ユーザ用
                                            Me.Initialize_Table(Hash_TblName_JpToAlpha(syoriitem), sql_where)
                                    End Select
                                Else
                                    Me.Initialize_Table(Hash_TblName_JpToAlpha(syoriitem), sql_where)
                                End If
                                '20161028 物件/部屋鍵取得方法修正 -add end
                            Case Else
                                Me.Initialize_Table(Hash_TblName_JpToAlpha(syoriitem), sql_where)
                        End Select
                        '20160720 連動情報構築 -chg end
                    End If

                    '中間ファイル読込→DB書込          
                    Dim midrowcnt As Integer = 0                '中間ファイル件数                                                        
                    Dim cvrowcnt As Integer = 0                 'コンバート件数
                    Dim conditioncnt As Integer = 0             '調整件数 (文字列切り捨て等)
                    Dim tmp_str As String = ""
                    Dim normalflg As Boolean = True             '正常処理判定

                    '----- 要チェック ----- sta
                    '→必要？
                    'DB書込前の個別処理
                    Select Case syorigrp
                        Case "契約情報"
                            '革命10用の契約行No、更新No、契約改定Noを作成した仮テーブル作成
                    End Select
                    '----- 要チェック ----- end

                    '20161124 物件部屋所有者情報有無による所有者情報のデータ調整処理を追加 -add sta
                    If CNVNO = ConvertTypes._汎用 Then
                        If syoriitem = "送金ルール基本情報" Then
                            Dim tmp_cnt As Integer = 0

                            '---------------------------------------------------------------------
                            '汎用コンバートで送金ルール基本情報を移行する際は
                            '物件部屋所有者情報を成形(データ調整)してから行う              '20161124 レビュー結果：sorule_guidを使う為
                            '※送金ルール作成には所有者情報にあるsorule_guidを使用するため
                            '---------------------------------------------------------------------

                            '物件・部屋所有者情報が重複して存在する場合、区分所有を優先して移行するため物件所有者情報を削除する
                            Dim sql_bksyodelete As String = DataCondModule.Get_UseQry_BkSyoDelete()
                            DBExec.Exec_NonQuery(sqlcnnv10, sql_bksyodelete, tmp_cnt)

                            '物件・部屋基本情報が存在し、所有者情報のレコードが存在しないデータに対して仮レコードを作成する
                            Dim sql_bkhysyoinsert As List(Of String) = DataCondModule.Get_List_UseQry_BkSyoInsert()
                            For Each tmp_sql In sql_bkhysyoinsert
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
                            Next
                        End If
                    End If
                    '20161124 物件部屋所有者情報有無による所有者情報のデータ調整処理を追加 -add end

                    'メイン処理
                    normalflg = obj_rep.Set_Vari(Me.sqlcnnv10, syorigrp, syoriitem, midrowcnt, cvrowcnt, conditioncnt)

                    '状況出力用情報格納                                          
                    If normalflg Then
                        tmp_str = tmp_str & vbCrLf & INDENT_2 & SITUATION_CVITEMTOTALCNT & INDENT_1 & midrowcnt.ToString
                        tmp_str = tmp_str & vbCrLf & INDENT_2 & SITUATION_CVITEMCVCNT & INDENT_1 & cvrowcnt.ToString
                        tmp_str = tmp_str & vbCrLf & INDENT_2 & SITUATION_CVITEMNOTCVCNT & INDENT_1 & (midrowcnt - cvrowcnt).ToString
                        tmp_str = tmp_str & vbCrLf & INDENT_2 & SITUATION_CVITEMCONDCNT & INDENT_1 & conditioncnt.ToString
                    End If

                    'ログ出力
                    Dim totalcntstr As String = SITUATION_CVITEMTOTALCNT.Trim & " " & midrowcnt.ToString
                    Dim cvcntstr As String = SITUATION_CVITEMCVCNT.Trim & " " & cvrowcnt.ToString
                    Dim notcvcntstr As String = SITUATION_CVITEMNOTCVCNT.Trim & " " & (midrowcnt - cvrowcnt).ToString
                    Dim condcntstr As String = SITUATION_CVITEMCONDCNT.Trim & " " & conditioncnt.ToString
                    Dim endkbnstr As String = Me.Chg_FlgToStr(normalflg, 2)
                    Dim outstr As String = endkbnstr & " " & totalcntstr & "/" & cvcntstr & "/" & notcvcntstr & "/" & condcntstr
                    Dim tmp_sql_itemsta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, syoriitem, "-", outstr), False)
                    DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_itemsta, rowcnt)

                    '状況出力
                    tmp_partsituation = INDENT_1 & Chg_FlgToStr(normalflg, 1) & syoriitem & tmp_str
                    tmp_totalsituation = syoriitem & INDENT_1 & SITUATION_WRITEEND
                    Call Me.Set_Situation(tmp_totalsituation, 0)
                    Call Me.Set_Situation(tmp_partsituation, 1)
                End If

                '----- 全体用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsettingTotal(pgbcnt)      '全体進捗
                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt, True)
                '20161018 個別進捗プログレスバーの表示修正 -add sta
                '----- 個別用プログレスバー更新 -----
                Dim pgbmax As Integer = pgbpartial.Maximum
                If pgbmax = 0 Then
                    pgbmax = 1
                    Call obj_pgb.pgbInitPart(pgbmax)
                    Call obj_pgb.pgbsettingPart(pgbmax)
                    Call obj_com.ProgressOutPut(pgbmax, pgbmax)
                End If
                '20161018 個別進捗プログレスバーの表示修正 -add end
            Next


            '--------------------------------------------------
            ' ログ/状況出力
            '--------------------------------------------------
            'ログ出力
            Dim tmp_sql_end As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, LOG_SYORIKOMK_DBCVEND), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_end, rowcnt)

            '状況出力
            tmp_partsituation = SITUATION_WRITEEND
            Call Me.Set_Situation(tmp_partsituation, 1)

        End Sub

        ''' <summary>
        ''' DB書込処理前準備
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Setting_BeforeCV()

            Dim rowcnt As Integer = 0


            '--------------------------------------------------
            ' 革命10DBテーブル名を取得して削除対象
            ' (DELETE構文のWHERE句)をハッシュテーブルへ格納
            '--------------------------------------------------
            '契約者照合用カナ情報のように契約者以外に家主や業者など他の情報も含まれている場合は、初期化時に削除対象を指定する必要がある
            '契約者照合用カナを移行する際は、登録されているデータのうち、区分=契約者を削除する必要がある
            hash_deltblqrywhere = New Hashtable
            Call Me.Set_DelTablename_qrywhere(hash_deltblqrywhere)


            '--------------------------------------------------
            ' 金融機関・支店情報を取得して
            ' ハッシュテーブルへ格納(マスタ参照用)
            '--------------------------------------------------
            Dim tmp_sql_getkinyu As String = DBQuery.Qry_GetKinyuInfo()
            Call GetBankInfo.Get_HashKinyu(Me.sqlcnnv10, tmp_sql_getkinyu)


            '--------------------------------------------------                 
            ' 実行者設定                                                   
            ' 初期設定画面で設定した作業者名を履歴フィールド[history]へ登録
            ' 未設定時はデフォルト値を設定 ("CONVUSER"を履歴フィールドへ登録)
            '--------------------------------------------------
            Dim tmp_executor As String = Me.txtRecUser.Text
            If tmp_executor = "" Then
                tmp_executor = DEF_REC_USER
            End If
            RecUser = tmp_executor


            '--------------------------------------------------
            '履歴設定
            '--------------------------------------------------
            Dim xml As New Njc.Common.HistorySetting
            Call xml.Set_HistoryData(RecUser)
            Dim str_xml As String = xml.ToXmlString
            DefHistory = str_xml

        End Sub

        ''' <summary>
        ''' テーブル名(日本語/削除対象)紐付けハッシュテーブル作成処理
        ''' </summary>
        ''' <param name="hash"></param>
        ''' <remarks></remarks>
        Private Sub Set_DelTablename_qrywhere(ByRef hash As Hashtable)

            '※固定値
            hash.Add("契約者照合用カナ情報", " sqsaki_kbn = 100 ")
            hash.Add("自社担当者情報", " logonuser_no NOT IN (99001,99999) ")
            hash.Add("備考タイトルマスタ", " memo_kbn IN (1,2,4,6,12) ")
            hash.Add("備考入力補助リストマスタ", " memo_kbn IN (12) ")
            hash.Add("画像タイトルマスタ", " gazo_kbn IN (1,2,3,8) ")        '20160720 連動情報構築 -add
            hash.Add("送信設定自社web情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 10 ")                '20160720 連動情報構築 -add
            hash.Add("送信設定HOMES情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 20 ")                  '20160720 連動情報構築 -add
            hash.Add("送信設定athome情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 30 ")                 '20160720 連動情報構築 -add
            hash.Add("送信設定SUUMO情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 40 ")                  '20160720 連動情報構築 -add
            hash.Add("広告補足自社web情報", " site_no = 10 ")                '20160720 連動情報構築 -add
            hash.Add("広告補足HOMES情報", " site_no = 20 ")                  '20160720 連動情報構築 -add
            hash.Add("広告補足athome情報", " site_no = 30 ")                 '20160720 連動情報構築 -add
            hash.Add("広告補足SUUMO情報", " site_no = 40 ")                  '20160720 連動情報構築 -add
            hash.Add("送信設定基本情報", " setting_sortorder <> 0 ")         '20160720 連動情報構築 -add

        End Sub

        ''' <summary>
        ''' テーブル初期化処理
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <param name="qrywhere"></param>
        ''' <remarks></remarks>
        Private Sub Initialize_Table(ByVal tblname As String, ByVal qrywhere As String)

            Dim rowcnt As Integer = 0

            '初期化処理実行
            Dim tmp_sqldelete As String = DBQuery.Qry_DelInfo(tblname, qrywhere)
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldelete, rowcnt)

        End Sub

        ''' <summary>
        ''' DB書込Repositoryの生成
        ''' </summary>
        ''' <param name="chkboxtext"></param>
        ''' <remarks></remarks>
        Private Function Get_ObjRep_DB(ByVal chkboxtext As String)

            Dim obj_rep As Object = Nothing

            Select Case chkboxtext

                '各マスタ情報
                Case model_cvitem.MstBus : obj_rep = New Njc.Repository.M_buskotu_Repository.SubConv                                    'バス交通マスタ
                Case model_cvitem.MstBusKotu : obj_rep = New Njc.Repository.M_buskotu_stop_Repository.SubConv                           'バス停マスタ
                    '20161028 物件/部屋鍵取得方法修正 -chg sta
                    'Case model_cvitem.MstKagititle : obj_rep = New Njc.Repository.M_kagi_title_Repository.SubConv                           '鍵タイトルマスタ
                Case model_cvitem.MstKagititle
                    Select Case CNVNO
                        Case ConvertTypes._既存ユーザ用
                            obj_rep = New Njc.Repository.M_kagi_title_Repository.SubConv                                              '鍵タイトルマスタ(既存用)
                        Case ConvertTypes._汎用
                            obj_rep = New Njc.Repository.M_kagi_title_BaseMid_Repository.SubConv                                      '鍵タイトルマスタ(汎用用)
                    End Select
                    '20161028 物件/部屋鍵取得方法修正 -chg end
                Case model_cvitem.MstKasyoClaimrui : obj_rep = New Njc.Repository.M_claim_rui_Repository.SubConv                        '箇所クレーム分類マスタ
                Case model_cvitem.MstTokuyaku : obj_rep = New Njc.Repository.M_tokuyaku_Repository.SubConv                              '特約マスタ
                    'Case model_cvitem.MstKeiyakurui : obj_rep = New Njc.Repository.M_ky_rui_Repository.SubConv                              '契約分類マスタ      
                Case model_cvitem.MstHokenrui : obj_rep = New Njc.Repository.M_hoken_rui_Repository.SubConv                             '保険種類マスタ
                Case model_cvitem.MstSchool : obj_rep = New Njc.Repository.M_koku_add_Repository.SubConv                                '学校区マスタ
                Case model_cvitem.MstArea : obj_rep = New Njc.Repository.M_area_Repository.SubConv                                      'エリアマスタ
                Case model_cvitem.MstHendo : obj_rep = New Njc.Repository.M_hendorule_Repository.SubConv                                '変動費マスタ
                Case model_cvitem.MstHendoitiran : obj_rep = New Njc.Repository.M_hendorule_itiran_Repository.SubConv                   '変動費一覧マスタ
                Case model_cvitem.MstBikotitle : obj_rep = New Njc.Repository.M_memo_Repository.SubConv                                 '備考タイトルマスタ      
                Case model_cvitem.MstBikolst : obj_rep = New Njc.Repository.M_memo_lst_Repository.SubConv                               '備考入力補助リストマスタ
                Case model_cvitem.MstGazotitle : obj_rep = New Njc.Repository.M_gazo_title_Repository.SubConv                           '画像タイトルマスタ
                    '仮マスタ読込先を"仮テーブルから中間ファイル"へ変更するためコメントアウト
                    'Case model_cvitem.MstNkinkomok : obj_rep = New Njc.Repository.M_nkin_Repository.SubConv                                '入金項目マスタ(仮テーブル読込)
                    'Case model_cvitem.MstBkbrui : obj_rep = New Njc.Repository.M_bk_rui_Repository.SubConv                                 '物件分類マスタ(仮テーブル読込)
                    'Case model_cvitem.MstNkinkbn : obj_rep = New Njc.Repository.M_nkbn_Repository.SubConv                                  '入金区分マスタ(仮テーブル読込)
                Case model_cvitem.MstKozasyubetu    '口座種別マスタ
                Case model_cvitem.MstKozo           '物件構造マスタ
                Case model_cvitem.MstTorihikitaiyo  '取引態様マスタ
                Case model_cvitem.MstHyrui          '部屋分類マスタ
                Case model_cvitem.MstSetubi         '設備マスタ

                    '業者情報
                Case model_cvitem.GyCyukaiBase : obj_rep = New Njc.Repository.Gydata_fudo_Repository.SubConv                            '仲介・管理業者マスタ
                Case model_cvitem.GyCyukaiKoza : obj_rep = New Njc.Repository.Gydata_fudokoza_Repository.SubConv                        '仲介・管理業者マスタ口座
                Case model_cvitem.GyCyukaiMemo : obj_rep = New Njc.Repository.Gydata_fudomemo_Repository.SubConv                        '仲介・管理業者マスタメモ
                Case model_cvitem.GyHokenBase : obj_rep = New Njc.Repository.Gydata_hoken_Repository.SubConv                            '保険業者マスタ
                Case model_cvitem.GyHokenKoza : obj_rep = New Njc.Repository.Gydata_hokenkoza_Repository.SubConv                        '保険業者マスタ口座
                Case model_cvitem.GyHokenMemo : obj_rep = New Njc.Repository.Gydata_hokenmemo_Repository.SubConv                        '保険業者マスタメモ
                Case model_cvitem.GyYatinhosyoBase : obj_rep = New Njc.Repository.Gydata_hosyo_Repository.SubConv                       '家賃保証業者マスタ
                    'Case model_cvitem.GyYatinhosyoKoza : obj_rep = New Njc.Repository.gydata_hosyokoza_Repository.SubConv                  '家賃保証業者マスタ口座
                Case model_cvitem.GyYatinhosyoMemo : obj_rep = New Njc.Repository.gydata_hosyomemo_Repository.SubConv                   '家賃保証業者マスタメモ
                Case model_cvitem.GySyuzenBase : obj_rep = New Njc.Repository.Gydata_szen_Repository.SubConv                            '修繕業者マスタ
                Case model_cvitem.GySyuzenKoza : obj_rep = New Njc.Repository.Gydata_szenkoza_Repository.SubConv                        '修繕業者マスタ口座
                Case model_cvitem.GySyuzenMemo : obj_rep = New Njc.Repository.Gydata_szenmemo_Repository.SubConv                        '修繕業者マスタメモ
                Case model_cvitem.GyLifelineBase : obj_rep = New Njc.Repository.Gydata_lifeline_Repository.SubConv                      'ライフラインマスタ
                Case model_cvitem.GySekoBase : obj_rep = New Njc.Repository.Gydata_seko_Repository.SubConv                              '施工会社マスタ
                Case model_cvitem.GySisetuBase : obj_rep = New Njc.Repository.Gydata_sisetu_Repository.SubConv                          '施設保守会社マスタ

                    '自社情報
                Case model_cvitem.JisyaBase : obj_rep = New Njc.Repository.Jisyadata_Repository.SubConv                                 '自社基本情報
                    '20160913_2 自社口座移行処理の追加 -chg sta
                    'Case model_cvitem.JisyaKoza : obj_rep = New Njc.Repository.Jisyadata_koza_Repository.SubConv                            '自社口座情報
                Case model_cvitem.JisyaKoza
                    Select Case CNVNO
                        Case ConvertTypes._既存ユーザ用
                            obj_rep = New Njc.Repository.Jisyadata_koza_Repository.SubConv                                              '自社口座情報(既存用)
                        Case ConvertTypes._汎用
                            obj_rep = New Njc.Repository.Jisyadata_koza_BaseMid_Repository.SubConv                                      '自社口座情報(汎用用)
                    End Select
                    '20160913_2 自社口座移行処理の追加 -chg end
                Case model_cvitem.JisyaTanto : obj_rep = New Njc.Repository.Profile_logonuser_Repository.SubConv                        '自社担当者情報       
                Case model_cvitem.JisyaMemo : obj_rep = New Njc.Repository.Jisyadata_memo_Repository.SubConv                            '自社メモ情報

                    '口座関連情報
                Case model_cvitem.FBFuriirai : obj_rep = New Njc.Repository.M_fb_sgfirai_Repository.SubConv                             '振込依頼人情報
                Case model_cvitem.FBFuritesuryo : obj_rep = New Njc.Repository.M_fb_sgfiraitesu_Repository.SubConv                      '振込手数料情報
                Case model_cvitem.FBKozafurikae : obj_rep = New Njc.Repository.M_fb_fkaejyoho_Repository.SubConv                        '口座振替情報
                Case model_cvitem.FBNsSyutoku : obj_rep = New Njc.Repository.M_fb_nskinsetting_Repository.SubConv                       '入出金取得情報
                Case model_cvitem.MstYatinKoza : obj_rep = New Njc.Repository.M_yatinkoza_Repository.SubConv                            '家賃入金口座情報
                Case model_cvitem.MstANSERArea : obj_rep = New Njc.Repository.M_spcarea_Repository.SubConv                              'ANSERエリア情報                 
                Case model_cvitem.MstANSERAccpoint : obj_rep = New Njc.Repository.M_spcaccesspoint_Repository.SubConv                   'ANSERアクセスポイント情報       
                Case model_cvitem.FBANSERSetuzoku : obj_rep = New Njc.Repository.M_spcsetuzoku_Repository.SubConv                       'ANSER接続情報                   

                    '家主情報
                Case model_cvitem.OwBase : obj_rep = New Njc.Repository.Owdata_Repository.SubConv                                       '家主情報(基本)
                Case model_cvitem.OwKoza : obj_rep = New Njc.Repository.Owdata_koza_Repository.SubConv                                  '家主情報(口座)
                Case model_cvitem.OwEvent : obj_rep = New Njc.Repository.Owdata_event_Repository.SubConv                                '家主情報(イベント)
                Case model_cvitem.OwMemo : obj_rep = New Njc.Repository.Owdata_memo_Repository.SubConv                                  '家主情報(メモ)

                    '契約者情報
                Case model_cvitem.KysBase : obj_rep = New Njc.Repository.Kysdata_Repository.SubConv                                     '契約者情報(基本)
                Case model_cvitem.KysKoza : obj_rep = New Njc.Repository.Kysdata_koza_Repository.SubConv                                '契約者情報(口座)
                Case model_cvitem.KysMemo : obj_rep = New Njc.Repository.Kysdata_memo_Repository.SubConv                                '契約者情報(メモ)
                Case model_cvitem.KysSyogoKana : obj_rep = New Njc.Repository.M_fb_fkomsyogo_kys_Repository.SubConv                     '契約者情報(照合用カナ)
                Case model_cvitem.KysHosyonin : obj_rep = New Njc.Repository.Kysdata_hosyo_Repository.SubConv                           '契約者情報(保証人)

                    '物件情報
                Case model_cvitem.BkBase : obj_rep = New Njc.Repository.Bkdata_Repository.SubConv                                       '物件基本情報
                Case model_cvitem.BkSyosai : obj_rep = New Njc.Repository.Bkdata_detail_Repository.SubConv                              '物件詳細情報
                Case model_cvitem.Bksyo : obj_rep = New Njc.Repository.Bkdata_syo_Repository.SubConv                                    '物件所有者情報
                Case model_cvitem.BkGomi : obj_rep = New Njc.Repository.Bkdata_dust_Repository.SubConv                                  '物件ゴミ情報
                Case model_cvitem.BkKenri : obj_rep = New Njc.Repository.Bkdata_kenri_Repository.SubConv                                '物件権利情報
                Case model_cvitem.BkKotu : obj_rep = New Njc.Repository.Bkdata_kotu_Repository.SubConv                                  '物件交通情報
                Case model_cvitem.BkSetudo : obj_rep = New Njc.Repository.Bkdata_setudo_Repository.SubConv                              '物件接道情報
                Case model_cvitem.BkSyuhen : obj_rep = New Njc.Repository.Bkdata_syuhen_Repository.SubConv                              '物件周辺情報
                Case model_cvitem.BkSzeniji : obj_rep = New Njc.Repository.Bkdata_szeniji_Repository.SubConv                            '物件修繕維持管理連絡先情報
                Case model_cvitem.BkMemo : obj_rep = New Njc.Repository.Bkdata_memo_Repository.SubConv                                  '物件メモ情報
                Case model_cvitem.BkKagi : obj_rep = New Njc.Repository.Bkdata_kagi_Repository.SubConv                                  '物件鍵情報(汎用ツールのみ)
                Case model_cvitem.BkKinrincyusyajo : obj_rep = New Njc.Repository.Bkdata_parkingother_Repository.SubConv                '物件近隣駐車場情報(汎用ツールのみ)
                Case model_cvitem.BkSansyofile : obj_rep = New Njc.Repository.Bkdata_relfile_Repository.SubConv                         '物件参照ファイル情報(汎用ツールのみ)
                Case model_cvitem.BkHendo : obj_rep = New Njc.Repository.Bkdata_hendo_Repository.SubConv                                '物件変動費親メーター情報(汎用ツールのみ)

                    '部屋情報
                Case model_cvitem.HyBase : obj_rep = New Njc.Repository.Hydata_Repository.SubConv                                       '部屋基本情報
                Case model_cvitem.HySyosai : obj_rep = New Njc.Repository.Hydata_detail_Repository.SubConv                              '部屋詳細情報
                Case model_cvitem.Hysyo : obj_rep = New Njc.Repository.Hydata_syo_Repository.SubConv                                    '部屋所有者情報
                Case model_cvitem.HyParking : obj_rep = New Njc.Repository.Hydata_parking_Repository.SubConv                            '部屋駐車場情報
                Case model_cvitem.HyTokuyaku : obj_rep = New Njc.Repository.Hydata_tokuyaku_Repository.SubConv                          '部屋特約情報
                Case model_cvitem.HyKagi : obj_rep = New Njc.Repository.Hydata_kagi_Repository.SubConv                                  '部屋鍵情報
                Case model_cvitem.HyMenseki : obj_rep = New Njc.Repository.Hydata_othermenseki_Repository.SubConv                       '部屋面積情報
                Case model_cvitem.HyMadoriutiwake : obj_rep = New Njc.Repository.Hydata_madoriutiwake_Repository.SubConv                '部屋間取内訳情報
                Case model_cvitem.HySzeniji : obj_rep = New Njc.Repository.Hydata_szeniji_Repository.SubConv                            '部屋修繕維持管理連絡先情報
                    '20160913_2 部屋設備移行処理の追加 -chg sta
                    'Case model_cvitem.HySetubi : obj_rep = New Njc.Repository.Hydata_setubilst_Repository.SubConv                           '部屋設備情報
                Case model_cvitem.HySetubi
                    Select Case CNVNO
                        Case ConvertTypes._既存ユーザ用
                            obj_rep = New Njc.Repository.Hydata_setubilst_Repository.SubConv                                            '部屋設備情報(既存用)
                        Case ConvertTypes._汎用
                            obj_rep = New Njc.Repository.Hydata_setubilst_BaseMid_Repository.SubConv                                    '部屋設備情報(汎用用)
                    End Select
                    '20160913_2 部屋設備移行処理の追加 -chg end
                Case model_cvitem.HyNkinkomk : obj_rep = New Njc.Repository.Hydata_nkin_Repository.SubConv                              '部屋入金項目情報
                Case model_cvitem.HyHendo : obj_rep = New Njc.Repository.Hydata_hendo_Repository.SubConv                                '部屋変動費各戸メーター情報
                Case model_cvitem.HyMemo : obj_rep = New Njc.Repository.Hydata_memo_Repository.SubConv                                  '部屋メモ情報
                Case model_cvitem.HyCommonsalespoint : obj_rep = New Njc.Repository.Hydata_commonsalespointparts_Repository.SubConv     '部屋共通セールスポイント情報(汎用ツールのみ)
                Case model_cvitem.HyConfirm : obj_rep = New Njc.Repository.Hydata_confirm_Repository.SubConv                            '部屋契約解約確認事項情報(汎用ツールのみ)
                Case model_cvitem.HyKenri : obj_rep = New Njc.Repository.Hydata_kenri_Repository.SubConv                                '部屋権利情報(汎用ツールのみ)
                Case model_cvitem.HySansyofile : obj_rep = New Njc.Repository.Hydata_relfile_Repository.SubConv                         '部屋参照ファイル情報(汎用ツールのみ)
                Case model_cvitem.HyGenjotanka : obj_rep = New Njc.Repository.Hydata_szen_Repository.SubConv                            '部屋原状回復目安単価情報(汎用ツールのみ)

                    '送金ルール情報
                Case model_cvitem.SoruleBase : obj_rep = New Njc.Repository.Sorule_Repository.SubConv                                   '送金ルール基本情報
                Case model_cvitem.SoruleSosaki : obj_rep = New Njc.Repository.Sorule_sosaki_Repository.SubConv                          '送金ルール送金先情報
                Case model_cvitem.SoruleNkin : obj_rep = New Njc.Repository.Sorule_nk_cmrule_Repository.SubConv                         '送金ルール入金項目情報
                Case model_cvitem.SoruleKojo : obj_rep = New Njc.Repository.Sorule_kojo_cmrule_Repository.SubConv                       '送金ルール控除項目情報

                    '契約情報
                Case model_cvitem.KyBase : obj_rep = New Njc.Repository.Kydata_Repository.SubConv                                       '契約基本情報
                Case model_cvitem.KyRireki : obj_rep = New Njc.Repository.Kydata_kihon_Repository.SubConv                               '契約履歴情報
                Case model_cvitem.KyKys : obj_rep = New Njc.Repository.Kydata_kys_Repository.SubConv                                    '契約契約者情報
                Case model_cvitem.KyNyukyo : obj_rep = New Njc.Repository.Kydata_nyukyo_Repository.SubConv                              '契約入居者情報
                Case model_cvitem.KyHosyonin : obj_rep = New Njc.Repository.Kydata_hosyonin_Repository.SubConv                          '契約保証人情報
                Case model_cvitem.KyCar : obj_rep = New Njc.Repository.Kydata_car_Repository.SubConv                                    '契約車情報
                Case model_cvitem.KyHoken : obj_rep = New Njc.Repository.Kydata_hoken_Repository.SubConv                                '契約保険情報
                Case model_cvitem.KyTokuyaku : obj_rep = New Njc.Repository.Kydata_tokuyaku_Repository.SubConv                          '契約特約事項情報
                Case model_cvitem.KyMemo : obj_rep = New Njc.Repository.Kydata_memo_Repository.SubConv                                  '契約メモ情報
                Case model_cvitem.KyNkinkomk : obj_rep = New Njc.Repository.Kydata_nkin_Repository.SubConv                              '契約入金項目情報
                Case model_cvitem.KyNkinkomkNx : obj_rep = New Njc.Repository.Kydata_nkin_nx_Repository.SubConv                         '契約次回入金項目情報
                Case model_cvitem.KyHendo : obj_rep = New Njc.Repository.Kydata_hendo_Repository.SubConv                                '契約変動費各戸メーター情報
                Case model_cvitem.KyKojoRule : obj_rep = New Njc.Repository.Kydata_kojorule_Repository.SubConv                          '契約控除ルール情報
                Case model_cvitem.KySorule : obj_rep = New Njc.Repository.Kydata_sorule_Repository.SubConv                              '契約送金ルール情報
                Case model_cvitem.KyKai : obj_rep = New Njc.Repository.Kydata_kai_Repository.SubConv                                    '契約解約情報
                    '20160531 鍵情報移行処理の修正 -del sta
                    ''2016.04.06 契約鍵情報の移行処理追加 -add
                    'Case model_cvitem.KyKagi : obj_rep = New Njc.Repository.Kydata_kagi_Repository.SubConv                                  '契約鍵情報
                    '20160531 鍵情報移行処理の修正 -del end
                Case model_cvitem.KySzen : obj_rep = New Njc.Repository.Kydata_kaiszen_Repository.SubConv                               '契約修繕見積情報       
                Case model_cvitem.KySzenmeisai : obj_rep = New Njc.Repository.Kydata_kaiszenmeisai_Repository.SubConv                   '契約修繕見積詳細情報   

                    '請求情報
                Case model_cvitem.SqKajyo : obj_rep = New Njc.Repository.Azukanri_Repository.SubConv                                    '過剰金情報
                Case model_cvitem.SqUnyotaino : obj_rep = New Njc.Repository.Unyotainodata_Repository.SubConv                           '運用開始時未収滞納金情報
                Case model_cvitem.SqSq : obj_rep = New Njc.Repository.Sqdata_Repository.SubConv                                         '請求情報
                Case model_cvitem.SqHendokensin : obj_rep = New Njc.Repository.Hendodata_meisai_Repository.SubConv                      '変動費検針情報
                Case model_cvitem.SqKoteiKojo : obj_rep = New Njc.Repository.Koteirule_Repository.SubConv                               '家主固定控除情報
                Case model_cvitem.SqSqKojo : obj_rep = New Njc.Repository.Kjdata_Repository.SubConv                                     '家主請求控除情報

                    'クレーム情報
                Case model_cvitem.ClaimBase : obj_rep = New Njc.Repository.Claimdata_Repository.SubConv                                 'クレーム基本情報
                Case model_cvitem.ClaimTaiorireki : obj_rep = New Njc.Repository.Claim_taio_Repository.SubConv                          'クレーム対応履歴情報
                Case model_cvitem.ClaimRelfile : obj_rep = New Njc.Repository.Claimdata_relfile_Repository.SubConv                      'クレーム関連ファイル情報

                    '修繕関連
                Case model_cvitem.SzenBase : obj_rep = New Njc.Repository.Szendata_Repository.SubConv                                   '修繕基本情報
                Case model_cvitem.SzenSzen : obj_rep = New Njc.Repository.Szendata_szen_Repository.SubConv                              '修繕見積情報
                Case model_cvitem.SzenSzenmeisai : obj_rep = New Njc.Repository.Szendata_szenmeisai_Repository.SubConv                  '修繕見積詳細情報
                Case model_cvitem.SzenClaim : obj_rep = New Njc.Repository.Szendata_claim_Repository.SubConv                            '修繕クレーム関連付け情報
                Case model_cvitem.SzenRelfile : obj_rep = New Njc.Repository.Szendata_relfile_Repository.SubConv                            '修繕関連ファイル情報       '20160627 修繕関連ファイル移行修正
                Case model_cvitem.SzenMemo : obj_rep = New Njc.Repository.Szendata_memo_Repository.SubConv                              '修繕メモ情報

                    '初期設定情報
                Case model_cvitem.SyskanriBase : obj_rep = New Njc.Repository.Profile_fk_base_Repository.SubConv                        '初期設定情報
                Case model_cvitem.SyskanriZei : obj_rep = New Njc.Repository.Profile_fk_zei_Repository.SubConv                          '税編集情報
                Case model_cvitem.SyskanriHenkanmoji : obj_rep = New Njc.Repository.Profile_fk_henkanmoji_Repository.SubConv            '変換文字情報
                Case model_cvitem.SyskanriNkinkomkmerge : obj_rep = New Njc.Repository.Profile_fk_nkinkomkmerge_Repository.SubConv      '入金項目集約情報

                    '物件データ連動情報
                Case model_cvitem.RendoSosinBase : obj_rep = New Njc.Repository.M_sendsetting_Repository.SubConv                        '送信設定基本情報
                Case model_cvitem.RendoSosinJisyaweb : obj_rep = New Njc.Repository.M_site_sendsetting_jisyaweb_Repository.SubConv      '送信設定自社web情報
                Case model_cvitem.RendoSosinHomes : obj_rep = New Njc.Repository.M_site_sendsetting_homes_Repository.SubConv            '送信設定HOMES情報
                Case model_cvitem.RendoSosinAthome : obj_rep = New Njc.Repository.M_site_sendsetting_athome_Repository.SubConv          '送信設定athome情報
                Case model_cvitem.RendoSosinSuumo : obj_rep = New Njc.Repository.M_site_sendsetting_suumo_Repository.SubConv            '送信設定SUUMO情報

                    'ポータル連動情報
                Case model_cvitem.RendoKokokuJisyaweb : obj_rep = New Njc.Repository.Hydata_kokoku_jisyaweb_Repository.SubConv          '広告補足自社web情報
                Case model_cvitem.RendoKokokuHomes : obj_rep = New Njc.Repository.Hydata_kokoku_homes_Repository.SubConv                '広告補足HOMES情報
                Case model_cvitem.RendoKokokuAthome : obj_rep = New Njc.Repository.Hydata_kokoku_athome_Repository.SubConv              '広告補足athome情報
                Case model_cvitem.RendoKokokuSuumo : obj_rep = New Njc.Repository.Hydata_kokoku_suumo_Repository.SubConv                '広告補足SUUMO情報
                Case model_cvitem.RendoHyrui : obj_rep = New Njc.Repository.M_hy_ruisite_Repository.SubConv                             'ポータル連動部屋分類情報
                Case model_cvitem.RendoHysosin : obj_rep = New Njc.Repository.Hydata_sosin_Repository.SubConv                           '部屋毎送信情報
                Case model_cvitem.RendoBtoBgroup : obj_rep = New Njc.Repository.Hydata_btobgroup_Repository.SubConv                     'BtoBグループ設定情報
                Case model_cvitem.RendoMapdisp : obj_rep = New Njc.Repository.Hydata_mapdisp_Repository.SubConv                         '地図表示詳細設定情報

            End Select

            Return obj_rep

        End Function

#End Region

#Region "中間ファイルチェック処理"

        ''' <summary>
        ''' 中間ファイルチェック実行前準備
        ''' </summary>
        ''' <param name="list_cv"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' ①コンバート用のDB情報を格納する仮テーブル作成
        ''' ②DBの基本情報と①を紐付けた値をハッシュテーブルへ格納
        ''' ※一度処理を行えば良いのでフラグを持たせて処理の実行有無を分岐させる
        ''' </remarks>
        Private Function Before_DBWrite(ByVal list_cv As List(Of String)) As Boolean

            Dim rowcnt As Integer = 0
            Dim rtn As Boolean = True


            '----------------------------------------------------------------------
            ' 革命10DBの情報に対してコンバート用のDB情報を付加する
            '----------------------------------------------------------------------
            '①付加する情報の仮テーブルを作成(自作仕様)
            '初期化(DROP)
            Dim cvdbinfo_dropqry As String = DBQuery.Qry_DropInfo(CVDBINFO_DBNAME, True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, cvdbinfo_dropqry, rowcnt)

            '仮テーブル作成(CREATE)
            Dim cvdbinfo_createqry As String = CVDBInfoModule.Get_CVDBInfo_CreateQry()
            DBExec.Exec_NonQuery(Me.sqlcnnv10, cvdbinfo_createqry, rowcnt)

            '②仮テーブルへコンバート用のDB情報を挿入
            '2016.02.22 エラー時の対処4 -add
            Dim normalflg As Boolean = True
            For Each item In list_cv
                Dim cvitem As String = item
                Dim cvdbinfo_insertqry As String = CVDBInfoModule.Get_CVDBInfo_InsertQry(item)
                If cvdbinfo_insertqry <> "" Then
                    normalflg = DBExec.Exec_NonQuery(Me.sqlcnnv10, cvdbinfo_insertqry, rowcnt)
                End If
                If normalflg = False Then
                    rtn = False
                    Return rtn
                End If
            Next


            '----------------------------------------------------------------------
            ' 革命10DBフィールド情報(型/サイズ)を取得してハッシュテーブルへ格納
            '----------------------------------------------------------------------
            Dim tmp_sql_getfldinfo As String = CVDBInfoModule.Qry_GetTableInfo()
            normalflg = GetFieldInfo.Get_HashFieldInfo(Me.sqlcnnv10, tmp_sql_getfldinfo)
            If normalflg = False Then
                rtn = False
                Return rtn
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 中間ファイルチェック処理 (DBを元に型・サイズチェックを行う)
        ''' </summary>
        ''' <param name="list_chkitem"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_MidFileChk_Main(list_chkitem As List(Of String)) As Boolean

            Dim rtn As Boolean = True
            '20161007 実行ボタン押下→DB書込開始までの進捗表示処理追加 -add sta
            Dim obj_pgb As New ProgressBarManager


            '--------------------------------------------------
            'ログ初期設定
            '--------------------------------------------------
            Call obj_pgb.pgbInitPart(pgbtotalcnt)                               '----- 個別用プログレスバー更新 -----
            Me.lblTotalSituation.Text = SITUATION_MID_CHKBEFORE
            Call Me.Set_LogInit(True)                                           'ログファイル出力


            '--------------------------------------------------
            '中間ファイルチェック処理
            '--------------------------------------------------
            Me.lblTotalSituation.Text = SITUATION_MID_CHK
            Dim normalflg As Boolean = Me.Get_MidFile_Data(list_chkitem)


            '--------------------------------------------------
            '終了処理
            '--------------------------------------------------
            Call obj_pgb.pgbInitPart(pgbtotalcnt)                               '----- 個別用プログレスバー更新 -----
            Me.lblTotalSituation.Text = SITUATION_MID_CHKAFTER

            If EtcMethod.Chk_FileExist(MiddleLogFilePath) = False Then          'ログファイル出力
                'ログファイル(CSV)のデータ部出力
                Dim tmp_sql As String = LogSetting.Get_LogTblSelectQry(True)
                Dim logoutflg As Boolean = True
                Dim errstr As String = ""
                logoutflg = FileMethod.TblView_Output_CSV(sqlcnnv10, MiddleLogFilePath, "", "", errstr, tmp_sql, True)
                'ヘッダーを加えて加工
                If logoutflg Then
                    '20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg sta
                    'Call Me.Set_LogHeader(MiddleLogFilePath, LOG_HEADER_TOTAL)
                    Dim tmp_header As String = LOG_HEADER_TOTAL
                    If Dev_CVFlg = False Then
                        tmp_header = tmp_header.Replace(",対象TBL名", "")
                    End If
                    Call Me.Set_LogHeader(MiddleLogFilePath, tmp_header)
                    '20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg end
                Else
                    MsgResult = MessageBox.Show(errstr, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            End If

            '返却
            rtn = normalflg
            Return rtn

        End Function

        ''' <summary>
        ''' 中間ファイルチェック実行処理→ログ出力 '20161006 中間ファイルチェック処理の速度改善対応 -add
        ''' </summary>
        ''' <param name="midchklist"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_MidFile_Data(midchklist As List(Of String)) As Boolean
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
            'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用   
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
            Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
            Dim normalflg As Boolean = True                                 'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値

            '----- 全体用・中間用プログレスバー更新 -----
            Dim pgbcnt As Integer = 0
            Dim pgbtotalcnt_total As Integer = midchklist.Count
            Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
            Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)
            Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)


            'ログ出力
            Log_OutputCnt = Int32.Parse(Me.txtLogOutputCnt.Text)            'ログ出力件数
            Dim tmp_sql_sta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTA), True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_sta, tmp_logcnt)

            '中間ファイルチェック処理
            For Each filesheetname In midchklist

                Dim tmp_str() As String = filesheetname.Split("-")
                Dim filename As String = tmp_str(0)
                Dim sheetname As String = tmp_str(1)
                Dim headervalue As New Object
                Dim fldvaluegrp As New Object
                Dim list_chkduplicate As New List(Of String)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
                Dim tmptblname As String = PRE_TBL_NAME & sheetname

                '----- デバッグ用処理 ----- sta
                'If filesheetname <> "自社情報-自社口座情報" Then
                '    GoTo chkskiplbl
                'End If
                '----- デバッグ用処理 ----- end

                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                ''汎用用中間ファイルを直接チェックする処理
                'Select Case filesheetname
                '    Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"
                '        filename = "中間ファイル"
                '        rtn = Me.Get_BaseMidFile_Data(filename, sheetname)
                '        If rtn = False Then
                '            Return rtn
                '        Else
                '            GoTo chkskiplbl
                '        End If
                '    Case "物件情報-物件変動費親メーター情報", "物件情報-物件近隣駐車場情報", "物件情報-(未使用の契約者データは移行しない)", "物件情報-(未使用の家主データは移行しない)", _
                '         "部屋情報-部屋面積情報", "部屋情報-部屋権利情報", _
                '         "契約情報-契約送金ルール情報", "契約情報-契約控除ルール情報", "契約情報-契約解約情報", "契約情報-契約修繕見積情報", _
                '         "契約情報-契約修繕見積詳細情報", "契約情報-契約変動費親メーター情報", "契約情報-契約解約確認事項情報", "契約情報-契約同時契約情報", _
                '         "契約情報-契約原状回復目安単価情報", "契約情報-契約敷金保証金随時処理情報", "契約情報-契約関連ファイル情報", "契約情報-契約空室待ち情報"
                '        GoTo chkskiplbl
                'End Select

                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

                '----- ログ出力 -----                                       '20161009 ログ出力処理追加 -add
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDCHECK & "：" & LOG_SYORIKOMK_MIDFILE & "(" & filesheetname & ")")
                Dim tmptmpstr As String
                Dim tmptmpcnt As Long
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(905, LOG_SYORIKOMK_MIDFILE, filesheetname), False)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    Call excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                'データ取得
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, Me.sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end

                '----- 個別用プログレスバー初期化 -----
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                'Dim rowcnt As Integer = readtbl.Rows.Count                  '行数取得   
                Me.lblCVItem.Text = filename & " ： " & sheetname           'チェック項目表示
                Dim pgbtotalcnt_part As Integer = 0                         'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100                             '実件数で表示するか否かの基準値
                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt_part = rowcnt
                Else
                    pgbtotalcnt_part = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

                '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                'Call Me.Set_MidDataInfoToObj(filename, sheetname, False)
                Select Case filesheetname
                    Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"
                        Call Me.Set_MidDataInfoToObj(CV_FROM_MIDDLE, sheetname, True)
                    Case Else
                        Call Me.Set_MidDataInfoToObj(filename, sheetname, False)
                End Select
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                ''ヘッダーチェック
                'Dim list_errheader As New List(Of String)
                'normalflg = Me.Chk_MidFileHeader(sheetname, readtbl, list_errheader)

                ''ヘッダーエラー時の終了処理
                'If normalflg = False Then

                '    For Each errheader In list_errheader
                '        'ログ用に文字列を編集してクエリへ加工
                '        Dim logvalue As String = ""
                '        Dim logvalue_insertqry As String = ""
                '        Dim tmp_cnt As Integer = 0

                '        'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
                '        If errheader <> "" Then
                '            logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
                '        Else
                '            logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
                '        End If

                '        'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
                '        logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
                '        DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
                '    Next

                '    'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
                '    Call excelfile.ExcelFile_ReadClose(con_read)
                '    rtn = normalflg
                '    Return rtn

                'End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                'データチェック
                For cntii = 0 To rowcnt - 1

                    '中断処理
                    Application.DoEvents()
                    If MidChkCancelFlg Or CancelFlg Then
                        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                        'excelfile.ExcelFile_ReadClose(con_read)
                        Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
                        DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
                        Return rtn
                    End If

                    '作業用変数
                    Dim hash_errdata As New Hashtable
                    Dim midkeytblfldname(10) As String
                    Dim midkeyvalue(10) As String
                    Dim tmp_hash As New Hashtable

                    'データ取得
                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                        '項目名取得
                        Dim fldname As String = readtbl.Columns(cntjj).ColumnName.Trim

                        '登録値取得
                        Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash.Add(fldname, fldvalue)

                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    ''ログ出力用のキーフィールドと値を取得
                    'For Each miditem In tmp_hash
                    '    Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
                    '    If Hash_ExistMiddatakey.Contains(get_key) Then
                    '        Dim tmp_nostr As String = Hash_ExistMiddatakey.Item(get_key)
                    '        Dim tmp_keyno As Integer = 0
                    '        If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
                    '            midkeytblfldname(tmp_keyno) = miditem.Key
                    '            midkeyvalue(tmp_keyno) = miditem.Value
                    '        End If
                    '    End If
                    'Next

                    'Dim tmp_keyfldnamejp As String
                    'Dim tmp_keyvalue As String = ""
                    'Dim tmp_taisyodata As String = ""
                    'For cntkk = 1 To UBound(midkeytblfldname)
                    '    If midkeytblfldname(cntkk) <> "" Then
                    '        tmp_keyfldnamejp = midkeytblfldname(cntkk)
                    '        tmp_keyvalue = midkeyvalue(cntkk)
                    '        tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
                    '    End If
                    'Next
                    'tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

                    ''データチェック
                    'Dim tmp_flg As Boolean = True
                    ''20160927 ログの内容が不正になっているため修正 -add sta
                    'If filename = "送金ルール情報" Then
                    '    tmp_hash("物件No") = tmp_hash("物件No") & "-" & tmp_hash("部屋No")
                    'End If
                    ''20160927 ログの内容が不正になっているため修正 -add end
                    'tmp_flg = Not (DataChk.Chk_DataPerItem_Mid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

                    Dim tmp_taisyodata As String = ""

                    Select Case filesheetname

                        Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"

                            'ログ出力用のキーフィールドと値を取得
                            For Each miditem In tmp_hash
                                Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
                                If Hash_BaseMiddatakey.Contains(get_key) Then
                                    Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
                                    Dim tmp_keyno As Integer = 0
                                    If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
                                        midkeytblfldname(tmp_keyno) = miditem.Key
                                        midkeyvalue(tmp_keyno) = miditem.Value
                                    End If
                                End If
                            Next

                            Dim tmp_keyfldnamejp As String
                            Dim tmp_keyvalue As String = ""
                            For cntkk = 1 To UBound(midkeytblfldname)
                                If midkeytblfldname(cntkk) <> "" Then
                                    tmp_keyfldnamejp = midkeytblfldname(cntkk)
                                    tmp_keyvalue = midkeyvalue(cntkk)
                                    tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
                                End If
                            Next
                            tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

                            'データチェック
                            Dim tmp_flg As Boolean = True
                            tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

                        Case Else

                            'ログ出力用のキーフィールドと値を取得
                            For Each miditem In tmp_hash
                                Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
                                If Hash_ExistMiddatakey.Contains(get_key) Then
                                    Dim tmp_nostr As String = Hash_ExistMiddatakey.Item(get_key)
                                    Dim tmp_keyno As Integer = 0
                                    If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
                                        midkeytblfldname(tmp_keyno) = miditem.Key
                                        midkeyvalue(tmp_keyno) = miditem.Value
                                    End If
                                End If
                            Next

                            Dim tmp_keyfldnamejp As String
                            Dim tmp_keyvalue As String = ""
                            For cntkk = 1 To UBound(midkeytblfldname)
                                If midkeytblfldname(cntkk) <> "" Then
                                    tmp_keyfldnamejp = midkeytblfldname(cntkk)
                                    tmp_keyvalue = midkeyvalue(cntkk)
                                    tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
                                End If
                            Next
                            tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

                            '20161028 物件/部屋鍵取得方法修正 -add sta
                            '鍵タイトルマスタの「鍵区分」はプログラム内部で使用する項目であるためログ出力用に成形する
                            If sheetname = "鍵タイトルマスタ" Then
                                tmp_taisyodata = tmp_taisyodata.Replace("鍵区分 = 1", "物件鍵タイトルマスタ")
                                tmp_taisyodata = tmp_taisyodata.Replace("鍵区分 = 2", "部屋鍵タイトルマスタ")
                            End If
                            '20161028 物件/部屋鍵取得方法修正 -add end

                            'データチェック
                            Dim tmp_flg As Boolean = True   '実際に移行処理が実行された場合のレコードの移行可/不可  True…移行不可 False…移行可
                            If filename = "送金ルール情報" Then
                                tmp_hash("物件No") = tmp_hash("物件No") & "-" & tmp_hash("部屋No")
                            End If
                            tmp_flg = Not (DataChk.Chk_DataPerItem_Mid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

                            '20161108_2 送金ルール送金先情報の移行制御処理修正 -add sta
                            '送金先情報は管理形態を元に必須かどうかを判別する
                            '自社物件、送金保留     … 必須ではない
                            '自社物件、送金保留以外 … 必須
                            If sheetname = "送金ルール送金先情報" Then
                                If tmp_flg = False Then
                                    Call Me.Chk_SoruleSosaki(tmp_hash, hash_errdata)
                                End If
                            End If
                            '20161108_2 送金ルール送金先情報の移行制御処理修正 -add end

                            '20161125 レビュー指摘事項対応 -chg sta
                            ''20161124 物件部屋所有者情報が両方存在する場合のチェック機能の追加 -add sta
                            'If sheetname = "物件所有者情報" Then
                            '    If tmp_flg = False Then                                 '20161124 レビュー結果：所有者情報のチェックはあってもいいのでは？
                            '        Call Me.Chk_BkHySyo(tmp_hash, hash_errdata)
                            '    End If
                            'End If
                            ''20161124 物件部屋所有者情報が両方存在する場合のチェック機能の追加 -add end
                            If sheetname = "物件所有者情報" Then
                                Call Me.Chk_BkHySyo(tmp_hash, hash_errdata)
                            End If
                            '20161125 レビュー指摘事項対応 -chg end

                    End Select
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end

                    'ログ出力処理
                    Dim list_log As New List(Of String)
                    For Each errvalue In hash_errdata

                        '20160927 ログの内容が不正になっているため修正 -chg sta
                        'Dim tmp_tblfldname As String = errvalue.Key
                        'Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
                        'Dim fldname As String = tmp_fldname(1)
                        Dim tmp_tblfldname As String = errvalue.Key
                        Dim tmp_sfname() As String = Split(tmp_tblfldname, STR_SPLIT_2)
                        Dim fldname As String = ""

                        For cntjj = 0 To UBound(tmp_sfname)
                            Dim tmp_sfsplietname() As String = Split(tmp_sfname(cntjj), STR_SPLIT_1)
                            fldname = fldname & "、" & tmp_sfsplietname(1)
                        Next
                        If fldname <> "" Then
                            fldname = fldname.Remove(0, 1)
                        End If
                        'Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
                        'Dim fldname As String = tmp_fldname(1)
                        '20160927 ログの内容が不正になっているため修正 -chg end
                        Dim errkomk() As String = errvalue.Value.Split("-")
                        Dim logvalue As String = ""
                        Dim logvalue_insertqry As String = ""
                        logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
                        logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
                        list_log.Add(logvalue_insertqry)

                        'ログ出力
                        If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
                            Dim tmp_cnt As Integer = 0                      '※Exec_NonQueryの第3引数の為用意
                            '挿入
                            For Each loggrp In list_log
                                DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
                            Next
                            '初期化
                            list_log.Clear()
                        End If

                    Next

                    '----- 個別用プログレスバー更新/進捗率表示 -----
                    Dim tmp_pgbcnt As Integer = 0
                    If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        tmp_pgbcnt = cntii + 1
                    ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, tmp_pgbcnt)
                    ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        tmp_pgbcnt = pgbtotalcnt_part
                    End If
                    If tmp_pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                        Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                    End If
                    '20161017 進捗表示処理による速度低下の修正 -del
                    'Me.Refresh()

                Next
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                ''クローズ処理
                'Call excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
chkskiplbl:
                '----- 全体用・中間用プログレスバー更新 -----           
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsettingTotal(pgbcnt)
                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
                Call obj_pgb.pgbsettingChkTotal(pgbcnt)
                Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
                Me.Refresh()
                Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDCHECK & "済：" & LOG_SYORIKOMK_MIDFILE & "(" & filesheetname & ")")                    '20161009 ログ出力処理追加 -add
            Next

            'チェック項目表示文字列初期化
            Me.lblCVItem.Text = ""


            '--------------------------------------------------
            'ログ出力
            '--------------------------------------------------
            Dim tmp_sql_end As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKEND), True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_end, tmp_logcnt)

            Return rtn

        End Function

        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイルチェック処理 '20161006 中間ファイルチェック処理の速度改善対応 -add
        ' ''' </summary>
        ' ''' <param name="filename"></param>
        ' ''' <param name="sheetname"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Get_BaseMidFile_Data(ByVal filename As String, ByVal sheetname As String) As Boolean

        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
        '    Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
        '    Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
        '    Dim normalflg As Boolean = True                                 'INSERT正常終了フラグ
        '    Dim rtn As Boolean = True                                       '戻り値

        '    'Excelファイル初期設定                  
        '    Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
        '    Dim readtbl As New DataTable()
        '    Dim con_read As New OleDbConnection()


        '    'オープン処理
        '    rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

        '    'オープン処理失敗時は処理を抜ける
        '    If rtn = False Then
        '        Call excelfile.ExcelFile_ReadClose(con_read)
        '        Return rtn
        '    End If

        '    '行数取得
        '    Dim rowcnt As Integer = readtbl.Rows.Count

        '    '----- 個別用プログレスバー更新 -----
        '    Me.lblCVItem.Text = filename & " ： " & sheetname                'チェック項目表示
        '    Dim pgbtotalcnt_part As Integer = 0                              'プログレスバー総件数初期化
        '    Dim pgbbasecnt As Integer = 100                                  '実件数で表示するか否かの基準値
        '    If rowcnt <= pgbbasecnt Then
        '        pgbtotalcnt_part = rowcnt
        '    Else
        '        pgbtotalcnt_part = Math.Ceiling(rowcnt / pgbbasecnt)
        '    End If
        '    Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        '    '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        '    Call Me.Set_MidDataInfoToObj(filename, sheetname, True)

        '    '格納用変数初期化
        '    Dim list_errheader As New List(Of String)
        '    Dim list_chkduplicate As New List(Of String)
        '    Dim list_header As New SortedList(Of Integer, String)           '自社口座情報ヘッダー格納用

        '    'データ取得(汎用中間ファイルには1行目にヘッダーが存在しないためデータとして取得し、行Noから判断する)
        '    For cntii = 0 To rowcnt - 1

        '        '中断処理
        '        Application.DoEvents()
        '        If MidChkCancelFlg Or CancelFlg Then
        '            Call excelfile.ExcelFile_ReadClose(con_read)
        '            Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        '            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        '            Return rtn
        '        End If

        '        Select Case sheetname
        '            Case "自社口座情報"

        '                'データ格納用変数
        '                Dim tmp_hash As New Hashtable

        '                If cntii = 0 Then

        '                    'ヘッダー取得
        '                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        '                        Dim fldname As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        '                        '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg sta
        '                        'list_header.Add(cntjj, fldname)
        '                        If fldname <> "" Then
        '                            list_header.Add(cntjj, fldname)
        '                        End If
        '                        '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg end
        '                    Next

        '                    'ヘッダーチェック
        '                    normalflg = Me.Chk_BaseMidFileHeader_Jisya(sheetname, list_header, list_errheader)
        '                    If normalflg = False Then
        '                        Exit For
        '                    End If

        '                ElseIf cntii >= BASEMIDFILE_READWRITE_ROW - 2 Then

        '                    'データ取得→チェック
        '                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        '                        Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        '                        tmp_hash.Add(list_header(cntjj), fldvalue)
        '                    Next

        '                    '不要なフィールドを削除
        '                    tmp_hash.Remove("項目名")

        '                    Dim hash_errdata As New Hashtable
        '                    Dim midkeytblfldname(10) As String
        '                    Dim midkeyvalue(10) As String

        '                    'ログ出力用のキーフィールドと値を取得
        '                    For Each miditem In tmp_hash
        '                        Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        '                        If Hash_BaseMiddatakey.Contains(get_key) Then
        '                            Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        '                            Dim tmp_keyno As Integer = 0
        '                            If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        '                                midkeytblfldname(tmp_keyno) = miditem.Key
        '                                midkeyvalue(tmp_keyno) = miditem.Value
        '                            End If
        '                        End If
        '                    Next

        '                    Dim tmp_keyfldnamejp As String
        '                    Dim tmp_keyvalue As String = ""
        '                    Dim tmp_taisyodata As String = ""
        '                    For cntkk = 1 To UBound(midkeytblfldname)
        '                        If midkeytblfldname(cntkk) <> "" Then
        '                            tmp_keyfldnamejp = midkeytblfldname(cntkk)
        '                            tmp_keyvalue = midkeyvalue(cntkk)
        '                            tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        '                        End If
        '                    Next
        '                    tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        '                    'データチェック
        '                    Dim tmp_flg As Boolean = True
        '                    tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        '                    'ログ出力処理
        '                    Dim list_log As New List(Of String)
        '                    For Each errvalue In hash_errdata

        '                        Dim tmp_tblfldname As String = errvalue.Key
        '                        Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        '                        Dim fldname As String = tmp_fldname(1)
        '                        Dim errkomk() As String = errvalue.Value.Split("-")
        '                        Dim logvalue As String = ""
        '                        Dim logvalue_insertqry As String = ""
        '                        logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        '                        logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '                        list_log.Add(logvalue_insertqry)

        '                        'ログ出力
        '                        If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        '                            Dim tmp_cnt As Integer = 0              '※Exec_NonQueryの第3引数の為用意
        '                            '挿入
        '                            For Each loggrp In list_log
        '                                DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        '                            Next
        '                            '初期化
        '                            list_log.Clear()
        '                        End If

        '                    Next

        '                End If

        '            Case "部屋設備情報"

        '                'データ格納用変数
        '                Dim tmp_hash As New Hashtable

        '                If cntii = 0 Then

        '                    'グループヘッダー取得
        '                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        '                        Dim fldname As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        '                        '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg sta
        '                        'fldname = fldname.Replace(vbLf, STR_SPLIT_1)
        '                        'list_header.Add(cntjj, fldname)
        '                        If fldname <> "" Then
        '                            fldname = fldname.Replace(vbLf, STR_SPLIT_1)
        '                            list_header.Add(cntjj, fldname)
        '                        End If
        '                        '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg end
        '                    Next

        '                    'ヘッダーチェック
        '                    normalflg = Me.Chk_BaseMidFileHeader_Jisya(sheetname, list_header, list_errheader)
        '                    If normalflg = False Then
        '                        Exit For
        '                    End If

        '                ElseIf cntii >= BASEMIDFILE_READWRITE_ROW - 2 Then

        '                    'データ取得→チェック
        '                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        '                        Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        '                        tmp_hash.Add(list_header(cntjj), fldvalue)
        '                    Next

        '                    '不要なフィールドを削除
        '                    tmp_hash.Remove("項目名")

        '                    Dim hash_errdata As New Hashtable
        '                    Dim midkeytblfldname(10) As String
        '                    Dim midkeyvalue(10) As String

        '                    'ログ出力用のキーフィールドと値を取得
        '                    For Each miditem In tmp_hash
        '                        Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        '                        If Hash_BaseMiddatakey.Contains(get_key) Then
        '                            Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        '                            Dim tmp_keyno As Integer = 0
        '                            If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        '                                midkeytblfldname(tmp_keyno) = miditem.Key
        '                                midkeyvalue(tmp_keyno) = miditem.Value
        '                            End If
        '                        End If
        '                    Next

        '                    Dim tmp_keyfldnamejp As String
        '                    Dim tmp_keyvalue As String = ""
        '                    Dim tmp_taisyodata As String = ""
        '                    For cntkk = 1 To UBound(midkeytblfldname)
        '                        If midkeytblfldname(cntkk) <> "" Then
        '                            tmp_keyfldnamejp = midkeytblfldname(cntkk)
        '                            tmp_keyvalue = midkeyvalue(cntkk)
        '                            tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        '                        End If
        '                    Next
        '                    tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        '                    'データチェック
        '                    Dim tmp_flg As Boolean = True
        '                    tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        '                    'ログ出力処理
        '                    Dim list_log As New List(Of String)
        '                    For Each errvalue In hash_errdata

        '                        Dim tmp_tblfldname As String = errvalue.Key
        '                        Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        '                        Dim fldname As String = ""
        '                        If UBound(tmp_fldname) = 1 Then
        '                            fldname = tmp_fldname(1)
        '                        ElseIf UBound(tmp_fldname) = 2 Then
        '                            fldname = tmp_fldname(1) & "-" & tmp_fldname(2)
        '                        End If

        '                        Dim errkomk() As String = errvalue.Value.Split("-")
        '                        Dim logvalue As String = ""
        '                        Dim logvalue_insertqry As String = ""
        '                        logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        '                        logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '                        list_log.Add(logvalue_insertqry)

        '                        'ログ出力
        '                        If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        '                            Dim tmp_cnt As Integer = 0              '※Exec_NonQueryの第3引数の為用意
        '                            '挿入
        '                            For Each loggrp In list_log
        '                                DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        '                            Next
        '                            '初期化
        '                            list_log.Clear()
        '                        End If

        '                    Next

        '                End If

        '        End Select


        '        '----- 個人用プログレスバー更新/進捗率表示 -----
        '        Dim tmp_pgbcnt As Integer = 0
        '        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
        '            tmp_pgbcnt = cntii + 1
        '        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
        '            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, tmp_pgbcnt)
        '        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
        '            tmp_pgbcnt = pgbtotalcnt_part
        '        End If
        '        If tmp_pgbcnt <> 0 Then
        '            Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
        '            Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
        '        End If
        '        '20161017 進捗表示処理による速度低下の修正 -del
        '        'Me.Refresh()

        '    Next

        '    'ヘッダーエラー時の終了処理
        '    If normalflg = False Then
        '        For Each errheader In list_errheader
        '            'ログ用に文字列を編集してクエリへ加工
        '            Dim logvalue As String = ""
        '            Dim logvalue_insertqry As String = ""
        '            Dim tmp_cnt As Integer = 0

        '            'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        '            If errheader <> "" Then
        '                logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        '            Else
        '                logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        '            End If

        '            'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        '            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '            DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        '        Next

        '        'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        '        Call excelfile.ExcelFile_ReadClose(con_read)
        '        rtn = normalflg
        '        Return rtn

        '    End If

        '    'クローズ処理
        '    Call excelfile.ExcelFile_ReadClose(con_read)

        '    '返却
        '    Return rtn

        'End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        '20161006 中間ファイルチェック処理の速度改善対応 -del sta
        '        ''' <summary>
        '        ''' 中間ファイルチェック実行処理→ログ出力
        '        ''' </summary>
        '        ''' <param name="midchklist"></param>
        '        ''' <returns></returns>
        '        ''' <remarks></remarks>
        '        Private Function Get_MidFile_Data(midchklist As List(Of String)) As Boolean

        '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '            Dim startrow As Integer                                         '書込開始行
        '            Dim columncnt As Integer                                        '列数
        '            Dim maxrowcnt As Integer                                        '既存データの行数
        '            Dim rowcnt As Integer                                           '書込行数
        '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '            Dim tmp_logcnt As Integer                                       '作業用

        '            Dim normalflg As Boolean                                        '正常終了フラグ
        '            Dim rtn As Boolean = True                                       '戻り値

        '            'ログ出力
        '            Log_OutputCnt = Int32.Parse(Me.txtLogOutputCnt.Text)            'ログ出力件数
        '            Dim tmp_sql_sta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTA), True)
        '            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_sta, tmp_logcnt)

        '            '中間ファイルチェック処理
        '            For Each filesheetname In midchklist

        '                Dim tmp_str() As String = filesheetname.Split("-")
        '                Dim filename As String = tmp_str(0)
        '                Dim sheetname As String = tmp_str(1)
        '                Dim headervalue As New Object
        '                Dim fldvaluegrp As New Object
        '                Dim list_chkduplicate As New List(Of String)    '20160812 汎用コンバート対応 -add

        '                '汎用用中間ファイルを直接チェックする処理
        '                '20160812 汎用コンバート対応 -add sta
        '                Select Case filesheetname
        '                    Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"
        '                        filename = "中間ファイル"
        '                        rtn = Me.Get_BaseMidFile_Data(filename, sheetname)
        '                        If rtn = False Then
        '                            Return rtn
        '                        Else
        '                            GoTo chkskiplbl
        '                        End If
        '                    Case "物件情報-物件変動費親メーター情報", "物件情報-物件近隣駐車場情報", "物件情報-(未使用の契約者データは移行しない)", "物件情報-(未使用の家主データは移行しない)", _
        '                         "部屋情報-部屋面積情報", "部屋情報-部屋権利情報", _
        '                         "契約情報-契約送金ルール情報", "契約情報-契約控除ルール情報", "契約情報-契約解約情報", "契約情報-契約修繕見積情報", _
        '                         "契約情報-契約修繕見積詳細情報", "契約情報-契約変動費親メーター情報", "契約情報-契約解約確認事項情報", "契約情報-契約同時契約情報", _
        '                         "契約情報-契約原状回復目安単価情報", "契約情報-契約敷金保証金随時処理情報", "契約情報-契約関連ファイル情報", "契約情報-契約空室待ち情報"
        '                        GoTo chkskiplbl         '不要なチェックボックスは削除すること
        '                End Select
        '                '20160812 汎用コンバート対応 -add end    

        '                'Excelファイル初期設定
        '                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

        '                'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '                If Not rtn Then
        '                    Return rtn
        '                End If
        '                '20160905 中間ファイルコピー処理改善 -add sta
        '                '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        '                Call Me.Set_MidDataInfoToObj(filename, sheetname, False)
        '                '20160905 中間ファイルコピー処理改善 -add end
        '                For cntii = 0 To rowcnt

        '                    '中断処理
        '                    Application.DoEvents()
        '                    If MidChkCancelFlg Or CancelFlg Then '20160929 データチェックのキャンセル処理を追加 Or CancelFlg の条件も追加 -chg
        '                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '                        Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        '                        DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        '                        Return rtn
        '                    End If

        '                    '格納用変数初期化
        '                    Dim list_errheader As New List(Of String)
        '                    Dim hash_errdata As New Hashtable

        '                    'データ取得
        '                    Dim rowvalue As Object
        '                    rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

        '                    If cntii = 0 Then
        '                        'ヘッダー格納用配列へコピー
        '                        headervalue = rowvalue.Clone

        '                        'ヘッダーチェック
        '                        normalflg = Me.Chk_MidFileHeader(sheetname, headervalue, list_errheader)

        '                        'エラー時の終了処理
        '                        If normalflg = False Then
        '                            For Each errheader In list_errheader
        '                                'ログ用に文字列を編集してクエリへ加工
        '                                Dim logvalue As String = ""
        '                                Dim logvalue_insertqry As String = ""
        '                                Dim tmp_cnt As Integer = 0

        '                                'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        '                                If errheader <> "" Then
        '                                    logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        '                                Else
        '                                    logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        '                                End If

        '                                'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        '                                logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '                                DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        '                            Next

        '                            'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        '                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '                            rtn = normalflg
        '                            Return rtn
        '                        End If
        '                    Else
        '                        '20160812 汎用コンバート対応 -chg sta
        '                        '                        Dim midkeytblfldname(10) As String
        '                        '                        Dim midkeyvalue(10) As String

        '                        '                        'ヘッダー格納用配列へコピー
        '                        '                        fldvaluegrp = rowvalue.Clone

        '                        '                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        '                        '                        Dim tmp_hash As New Hashtable
        '                        '                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(sheetname, headervalue, fldvaluegrp)

        '                        '                        'データチェック
        '                        '                        For Each fldvalue In tmp_hash

        '                        '                            Dim hashkey As String = fldvalue.Key
        '                        '                            Dim value As String = fldvalue.Value
        '                        '                            Dim errstr As String = ""

        '                        '                            'キーフィールドの判別
        '                        '                            If Hash_MidKey_FieldJp.Contains(hashkey) Then
        '                        '                                Dim tmp_nostr As String = Hash_MidKey_FieldJp.Item(hashkey)
        '                        '                                Dim tmp_keyno As Integer = 0
        '                        '                                If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        '                        '                                    midkeytblfldname(tmp_keyno) = hashkey
        '                        '                                    midkeyvalue(tmp_keyno) = value
        '                        '                                Else
        '                        '                                    errstr = LOG_NAIYO_ERR_DATACHK & "-" & LOG_HUBI_ERR_DATACHK_C & "-" & LOG_TAISYO_DATACHK_A
        '                        '                                    GoTo chkskiplabel
        '                        '                                End If
        '                        '                            End If

        '                        '                            'データチェック処理
        '                        '                            Call DataChk.Call_ChkMethodMidTotal(hashkey, value, errstr)
        '                        'chkskiplabel:
        '                        '                            If errstr <> "" Then
        '                        '                                Dim fldname() As String = hashkey.Split("-")
        '                        '                                hash_errdata.Add(fldname(1), errstr)
        '                        '                            End If

        '                        '                        Next

        '                        '                        Dim tmp_keyfldname() As String
        '                        '                        Dim tmp_keyvalue As String = ""
        '                        '                        Dim tmp_taisyodata As String = ""
        '                        '                        Dim list_log As New List(Of String)

        '                        '                        '退避しておいたキーフィールドと値を取得
        '                        '                        For cntkk = 1 To UBound(midkeytblfldname)
        '                        '                            If midkeytblfldname(cntkk) <> "" Then
        '                        '                                tmp_keyfldname = midkeytblfldname(cntkk).Split("-")
        '                        '                                tmp_keyvalue = tmp_hash.Item(midkeytblfldname(cntkk))
        '                        '                                tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldname(1) & " = " & tmp_keyvalue
        '                        '                            End If
        '                        '                        Next
        '                        '                        tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        '                        '                        'ログ出力処理
        '                        '                        For Each errvalue In hash_errdata

        '                        '                            Dim fldname As String = errvalue.Key
        '                        '                            Dim errkomk() As String = errvalue.Value.Split("-")
        '                        '                            Dim logvalue As String = ""
        '                        '                            Dim logvalue_insertqry As String = ""
        '                        '                            logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        '                        '                            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '                        '                            list_log.Add(logvalue_insertqry)

        '                        '                            'ログ出力
        '                        '                            If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        '                        '                                Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        '                        '                                '挿入
        '                        '                                For Each loggrp In list_log
        '                        '                                    DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        '                        '                                Next
        '                        '                                '初期化
        '                        '                                list_log.Clear()
        '                        '                            End If

        '                        '                        Next
        '                        Dim midkeytblfldname(10) As String
        '                        Dim midkeyvalue(10) As String

        '                        'ヘッダー格納用配列へコピー
        '                        fldvaluegrp = rowvalue.Clone

        '                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        '                        Dim tmp_hash As New Hashtable
        '                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)

        '                        'ログ出力用のキーフィールドと値を取得
        '                        For Each miditem In tmp_hash
        '                            Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        '                            If Hash_ExistMiddatakey.Contains(get_key) Then
        '                                Dim tmp_nostr As String = Hash_ExistMiddatakey.Item(get_key)
        '                                Dim tmp_keyno As Integer = 0
        '                                If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        '                                    midkeytblfldname(tmp_keyno) = miditem.Key
        '                                    midkeyvalue(tmp_keyno) = miditem.Value
        '                                End If
        '                            End If
        '                        Next

        '                        Dim tmp_keyfldnamejp As String
        '                        Dim tmp_keyvalue As String = ""
        '                        Dim tmp_taisyodata As String = ""
        '                        For cntkk = 1 To UBound(midkeytblfldname)
        '                            If midkeytblfldname(cntkk) <> "" Then
        '                                tmp_keyfldnamejp = midkeytblfldname(cntkk)
        '                                tmp_keyvalue = midkeyvalue(cntkk)
        '                                tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        '                            End If
        '                        Next
        '                        tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        '                        'データチェック
        '                        Dim tmp_flg As Boolean = True
        '                        '20160927 ログの内容が不正になっているため修正 -add sta
        '                        If filename = "送金ルール情報" Then
        '                            tmp_hash("物件No") = tmp_hash("物件No") & "-" & tmp_hash("部屋No")
        '                        End If
        '                        '20160927 ログの内容が不正になっているため修正 -add end
        '                        tmp_flg = Not (DataChk.Chk_DataPerItem_Mid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        '                        'ログ出力処理
        '                        Dim list_log As New List(Of String)
        '                        For Each errvalue In hash_errdata

        '                            '20160927 ログの内容が不正になっているため修正 -chg sta
        '                            'Dim tmp_tblfldname As String = errvalue.Key
        '                            'Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        '                            'Dim fldname As String = tmp_fldname(1)
        '                            Dim tmp_tblfldname As String = errvalue.Key
        '                            Dim tmp_sfname() As String = Split(tmp_tblfldname, STR_SPLIT_2)
        '                            Dim fldname As String = ""

        '                            For cntjj = 0 To UBound(tmp_sfname)
        '                                Dim tmp_sfsplietname() As String = Split(tmp_sfname(cntjj), STR_SPLIT_1)
        '                                fldname = fldname & "、" & tmp_sfsplietname(1)
        '                            Next
        '                            If fldname <> "" Then
        '                                fldname = fldname.Remove(0, 1)
        '                            End If
        '                            'Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        '                            'Dim fldname As String = tmp_fldname(1)
        '                            '20160927 ログの内容が不正になっているため修正 -chg end
        '                            Dim errkomk() As String = errvalue.Value.Split("-")
        '                            Dim logvalue As String = ""
        '                            Dim logvalue_insertqry As String = ""
        '                            logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        '                            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '                            list_log.Add(logvalue_insertqry)

        '                            'ログ出力
        '                            If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        '                                Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        '                                '挿入
        '                                For Each loggrp In list_log
        '                                    DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        '                                Next
        '                                '初期化
        '                                list_log.Clear()
        '                            End If

        '                        Next
        '                        '20160812 汎用コンバート対応 -chg end
        '                    End If

        '                Next

        '                'Excelファイル終了設定
        '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        'chkskiplbl:     '20160812 汎用コンバート対応 -add

        '            Next

        '            '--------------------------------------------------
        '            'ログ出力
        '            '--------------------------------------------------
        '            Dim tmp_sql_end As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKEND), True)
        '            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_end, tmp_logcnt)

        '            '20160907 EXCELが開いたままになっている不具合 -coment：汎用用の中間ファイルが活きているので閉じておく
        '            '(別エクセルファイルを開いているとき、裏で活きたままになる)
        '            Return rtn

        '        End Function
        '20161006 中間ファイルチェック処理の速度改善対応 -del end

        '20161006 中間ファイルチェック処理の速度改善対応 -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイルチェック処理 20160812 汎用コンバート対応
        ' ''' </summary>
        ' ''' <param name="filename"></param>
        ' ''' <param name="sheetname"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Get_BaseMidFile_Data(ByVal filename As String, ByVal sheetname As String) As Boolean

        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim tmp_logcnt As Integer                                       '作業用

        '    Dim normalflg As Boolean                                        '正常終了フラグ
        '    Dim rtn As Boolean = True                                       '戻り値
        '    Dim headervalue As New Object
        '    Dim fldvaluegrp As New Object
        '    Dim list_chkduplicate As New List(Of String)
        '    '20160913_2 不要処理の削除 -del sta
        '    ''ヘッダー行不定のためここで一時的に設定しておく
        '    'If sheetname = "部屋設備情報" Then                    '20160913 これなんでしたっけ？
        '    '    startrow = 3
        '    'End If
        '    '20160913_2 不要処理の削除 -del end
        '    '20160907 EXCELが開いたままになっている不具合 -coment：オープン処理内でエラーの場合の処理(ログ出力)がない？
        '    'Excelファイル初期設定
        '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

        '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '    If Not rtn Then
        '        Return rtn
        '    End If
        '    '20160905 中間ファイルコピー処理改善 -add sta
        '    '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        '    Call Me.Set_MidDataInfoToObj(filename, sheetname, True)
        '    '20160905 中間ファイルコピー処理改善 -add end
        '    '20160913_2 汎用用中間ファイルチェック処理の全体的な修正 -chg sta
        '    '↓↓↓旧srcコメントアウト↓↓↓
        '    'For cntii = 0 To rowcnt

        '    '    '中断処理
        '    '    Application.DoEvents()
        '    '    If MidChkCancelFlg Then

        '    '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '    '        Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        '    '        DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        '    '        Return rtn
        '    '    End If

        '    '    '格納用変数初期化
        '    '    Dim list_errheader As New List(Of String)
        '    '    Dim hash_errdata As New Hashtable

        '    '    'データ取得
        '    '    Dim rowvalue As Object
        '    '    '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg sta
        '    '    'rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value
        '    '    If cntii = 0 Then                                                       '20160913 ★cnt=0の時の処理について別途教えて下さい。★
        '    '        Dim headerrow As Integer = 2
        '    '        rowvalue = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value
        '    '    Else
        '    '        rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 2), wsheet.Cells(cntii + startrow - 1, columncnt)).Value
        '    '    End If
        '    '    '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg end
        '    '    If cntii = 0 Then                                               '20160913 ↑と↓の処理は纏められない？
        '    '        'ヘッダー格納用配列へコピー
        '    '        headervalue = rowvalue.Clone

        '    '        'ヘッダーチェック
        '    '        normalflg = Me.Chk_BaseMidFileHeader(sheetname, headervalue, list_errheader)

        '    '        'エラー時の終了処理
        '    '        If normalflg = False Then
        '    '            For Each errheader In list_errheader
        '    '                'ログ用に文字列を編集してクエリへ加工
        '    '                Dim logvalue As String = ""
        '    '                Dim logvalue_insertqry As String = ""
        '    '                Dim tmp_cnt As Integer = 0

        '    '                'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        '    '                If errheader <> "" Then
        '    '                    logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        '    '                Else
        '    '                    logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        '    '                End If

        '    '                'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        '    '                logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '    '                DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        '    '            Next

        '    '            'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        '    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '    '            rtn = normalflg
        '    '            Return rtn
        '    '        End If
        '    '    Else

        '    '        Dim midkeytblfldname(10) As String
        '    '        Dim midkeyvalue(10) As String

        '    '        'ヘッダー格納用配列へコピー
        '    '        fldvaluegrp = rowvalue.Clone

        '    '        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        '    '        Dim tmp_hash As New Hashtable
        '    '        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)

        '    '        'ログ出力用のキーフィールドと値を取得
        '    '        For Each miditem In tmp_hash
        '    '            Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        '    '            If Hash_BaseMiddatakey.Contains(get_key) Then
        '    '                Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        '    '                Dim tmp_keyno As Integer = 0
        '    '                If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        '    '                    midkeytblfldname(tmp_keyno) = miditem.Key
        '    '                    midkeyvalue(tmp_keyno) = miditem.Value
        '    '                End If
        '    '            End If
        '    '        Next

        '    '        Dim tmp_keyfldnamejp As String
        '    '        Dim tmp_keyvalue As String = ""
        '    '        Dim tmp_taisyodata As String = ""
        '    '        For cntkk = 1 To UBound(midkeytblfldname)
        '    '            If midkeytblfldname(cntkk) <> "" Then
        '    '                tmp_keyfldnamejp = midkeytblfldname(cntkk)
        '    '                tmp_keyvalue = midkeyvalue(cntkk)
        '    '                tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        '    '            End If
        '    '        Next
        '    '        tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        '    '        'データチェック
        '    '        Dim tmp_flg As Boolean = True
        '    '        tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        '    '        'ログ出力処理
        '    '        Dim list_log As New List(Of String)
        '    '        For Each errvalue In hash_errdata

        '    '            Dim tmp_tblfldname As String = errvalue.Key
        '    '            Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        '    '            Dim fldname As String = tmp_fldname(1)
        '    '            Dim errkomk() As String = errvalue.Value.Split("-")
        '    '            Dim logvalue As String = ""
        '    '            Dim logvalue_insertqry As String = ""
        '    '            logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        '    '            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '    '            list_log.Add(logvalue_insertqry)

        '    '            'ログ出力
        '    '            If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        '    '                Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        '    '                '挿入
        '    '                For Each loggrp In list_log
        '    '                    DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        '    '                Next
        '    '                '初期化
        '    '                list_log.Clear()
        '    '            End If

        '    '        Next
        '    '    End If

        '    'Next
        '    '↑↑↑旧srcコメントアウト↑↑↑

        '    '格納用変数初期化
        '    Dim list_errheader As New List(Of String)
        '    Dim hash_errdata As New Hashtable

        '    'ヘッダー取得
        '    Select Case sheetname
        '        Case "自社口座情報"

        '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
        '            headervalue = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value

        '            'ヘッダーチェック
        '            normalflg = Me.Chk_BaseMidFileHeader(sheetname, headervalue, list_errheader)

        '        Case "部屋設備情報"

        '            Dim headerrow_grp As Integer = BASEMIDFILE_HEADER_ROW_SETUBIGRP
        '            Dim headerrow_setubi As Integer = BASEMIDFILE_HEADER_ROW
        '            Dim headervalue_grp As Object = wsheet.Range(wsheet.Cells(headerrow_grp, 2), wsheet.Cells(headerrow_grp, columncnt)).Value
        '            Dim headervalue_setubi As Object = wsheet.Range(wsheet.Cells(headerrow_setubi, 2), wsheet.Cells(headerrow_setubi, columncnt)).Value
        '            headervalue = EtcMethod.Set_Header(headervalue_grp, headervalue_setubi)

        '            'ヘッダーチェック
        '            normalflg = Me.Chk_BaseMidFileHeader_Setubi(sheetname, headervalue, list_errheader)

        '    End Select

        '    'エラー時の終了処理
        '    If normalflg = False Then
        '        For Each errheader In list_errheader
        '            'ログ用に文字列を編集してクエリへ加工
        '            Dim logvalue As String = ""
        '            Dim logvalue_insertqry As String = ""
        '            Dim tmp_cnt As Integer = 0

        '            'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        '            If errheader <> "" Then
        '                logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        '            Else
        '                logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        '            End If

        '            'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        '            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '            DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        '        Next

        '        'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '        rtn = normalflg
        '        Return rtn
        '    End If

        '    'データチェック
        '    For cntii = 1 To rowcnt

        '        '中断処理
        '        Application.DoEvents()
        '        If MidChkCancelFlg Or CancelFlg Then '20160929 データチェックのキャンセル処理を追加 Or CancelFlg の条件も追加 -chg
        '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        '            Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        '            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        '            Return rtn
        '        End If

        '        'データ取得
        '        Dim rowvalue As Object
        '        rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 2), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

        '        Dim midkeytblfldname(10) As String
        '        Dim midkeyvalue(10) As String

        '        'ヘッダー格納用配列へコピー
        '        fldvaluegrp = rowvalue.Clone

        '        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        '        Dim tmp_hash As New Hashtable
        '        '20160913_2 部屋設備移行処理の追加 -chg sta
        '        'tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)
        '        Select Case sheetname
        '            Case "自社口座情報"
        '                tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)
        '            Case "部屋設備情報"
        '                tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue_Setubi(sheetname, headervalue, fldvaluegrp)
        '        End Select
        '        '20160913_2 部屋設備移行処理の追加 -chg end

        '        'ログ出力用のキーフィールドと値を取得
        '        For Each miditem In tmp_hash
        '            Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        '            If Hash_BaseMiddatakey.Contains(get_key) Then
        '                Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        '                Dim tmp_keyno As Integer = 0
        '                If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        '                    midkeytblfldname(tmp_keyno) = miditem.Key
        '                    midkeyvalue(tmp_keyno) = miditem.Value
        '                End If
        '            End If
        '        Next

        '        Dim tmp_keyfldnamejp As String
        '        Dim tmp_keyvalue As String = ""
        '        Dim tmp_taisyodata As String = ""
        '        For cntkk = 1 To UBound(midkeytblfldname)
        '            If midkeytblfldname(cntkk) <> "" Then
        '                tmp_keyfldnamejp = midkeytblfldname(cntkk)
        '                tmp_keyvalue = midkeyvalue(cntkk)
        '                tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        '            End If
        '        Next
        '        tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        '        'データチェック
        '        Dim tmp_flg As Boolean = True
        '        tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        '        'ログ出力処理
        '        Dim list_log As New List(Of String)
        '        For Each errvalue In hash_errdata

        '            Dim tmp_tblfldname As String = errvalue.Key
        '            Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        '            Dim fldname As String = tmp_fldname(1)
        '            Dim errkomk() As String = errvalue.Value.Split("-")
        '            Dim logvalue As String = ""
        '            Dim logvalue_insertqry As String = ""
        '            logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        '            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        '            list_log.Add(logvalue_insertqry)

        '            'ログ出力
        '            If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        '                Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        '                '挿入
        '                For Each loggrp In list_log
        '                    DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        '                Next
        '                '初期化
        '                list_log.Clear()
        '            End If

        '        Next

        '    Next
        '    '20160913_2 汎用用中間ファイルチェック処理の全体的な修正 -chg end

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    Return rtn

        'End Function
        '20161006 中間ファイルチェック処理の速度改善対応 -del end

        ' ''' <summary>
        ' ''' 個別でチェックを行う処理
        ' ''' </summary>
        ' ''' <param name="filename"></param>
        ' ''' <param name="sheetname"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Chk_MidFile_Chg(ByVal filename As String, ByVal sheetname As String)

        '    Dim rtn As Boolean = False
        '    Dim filesheetname As String = filename & "-" & sheetname

        '    Select Case filesheetname

        '        Case "家主情報-家主イベント情報"
        '            Call Njc.Repository.Chk_middata_Repository.Chk_owdata_event_mid(sqlcnnv10, filename, sheetname) : Return rtn
        '        Case "家主情報-家主メモ情報"
        '            Call Njc.Repository.Chk_middata_Repository.Chk_owdata_memo_mid(sqlcnnv10, filename, sheetname) : Return rtn




        '        Case Else
        '            rtn = True
        '    End Select

        '    Return rtn

        'End Function

        '20161006 中間ファイルチェック処理の速度改善対応 -del sta
        ' ''' <summary>
        ' ''' 既存用中間ファイルヘッダーチェック
        ' ''' </summary>
        ' ''' <param name="tblnamejp"></param>
        ' ''' <param name="headervalue"></param>
        ' ''' <param name="list_errheader"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Chk_MidFileHeader(ByVal tblnamejp As String, ByVal headervalue As Object, ByRef list_errheader As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True

        '    For cntii = 1 To headervalue.Length

        '        Dim tmp_header As String = tblnamejp & STR_SPLIT_1 & headervalue(1, cntii)

        '        '20160812 汎用コンバート対応 -chg sta
        '        '汎用-既存照合用仮テーブルから取得したものと比較するように修正
        '        'If Not List_FldNameJp.Contains(tmp_header) Then
        '        '    list_errheader.Add(headervalue(1, cntii))
        '        'End If
        '        If List_Existmidheader.Contains(tmp_header) = False Then
        '            list_errheader.Add(headervalue(1, cntii))
        '        End If
        '        '20160812 汎用コンバート対応 -chg end

        '    Next

        '    If list_errheader.Count <> 0 Then
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20161006 中間ファイルチェック処理の速度改善対応 -del end

        ''' <summary>
        ''' 既存用中間ファイルヘッダーチェック '20161006 中間ファイルチェック処理の速度改善対応 -add
        ''' </summary>
        ''' <param name="tblnamejp"></param>
        ''' <param name="dt"></param>
        ''' <param name="list_errheader"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_MidFileHeader(ByVal tblnamejp As String, ByVal dt As DataTable, ByRef list_errheader As List(Of String)) As Boolean

            Dim rtn As Boolean = True

            For cntjj As Integer = 0 To dt.Columns.Count - 1
                Dim tmp_header As String = tblnamejp & STR_SPLIT_1 & dt.Columns(cntjj).ColumnName.Trim
                If List_Existmidheader.Contains(tmp_header) = False Then
                    list_errheader.Add(dt.Columns(cntjj).ColumnName.Trim)
                End If
            Next

            If list_errheader.Count <> 0 Then
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 汎用用中間ファイル自社口座情報ヘッダーチェック '20161006 中間ファイルチェック処理の速度改善対応 -add
        ''' </summary>
        ''' <param name="tblnamejp"></param>
        ''' <param name="solist_header"></param>
        ''' <param name="list_errheader"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_BaseMidFileHeader_Jisya(ByVal tblnamejp As String, ByVal solist_header As SortedList(Of Integer, String), ByRef list_errheader As List(Of String)) As Boolean

            Dim rtn As Boolean = True

            For cntii = 0 To solist_header.Count - 1
                Dim tmp_header As String = tblnamejp & STR_SPLIT_1 & solist_header(cntii)
                If solist_header(cntii) <> "項目名" Then
                    If List_Basemidheader.Contains(tmp_header) = False Then
                        list_errheader.Add(solist_header(cntii))
                    End If
                End If
            Next

            If list_errheader.Count <> 0 Then
                rtn = False
            End If

            Return rtn

        End Function

        '20161006 中間ファイルチェック処理の速度改善対応 -del sta
        ' ''' <summary>
        ' ''' 汎用用中間ファイルヘッダーチェック
        ' ''' </summary>
        ' ''' <param name="tblnamejp"></param>
        ' ''' <param name="headervalue"></param>
        ' ''' <param name="list_errheader"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Chk_BaseMidFileHeader(ByVal tblnamejp As String, ByVal headervalue As Object, ByRef list_errheader As List(Of String)) As Boolean

        '    Dim rtn As Boolean = True

        '    For cntii = 1 To headervalue.Length

        '        Dim tmp_header As String = tblnamejp & STR_SPLIT_1 & headervalue(1, cntii)

        '        '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        '        'If List_Basemidheader.Contains(tmp_header) = False Then
        '        '    list_errheader.Add(headervalue(1, cntii))
        '        'End If
        '        If headervalue(1, cntii) <> "項目名" Then
        '            If List_Basemidheader.Contains(tmp_header) = False Then
        '                list_errheader.Add(headervalue(1, cntii))
        '            End If
        '        End If
        '        '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        '    Next

        '    If list_errheader.Count <> 0 Then
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20161006 中間ファイルチェック処理の速度改善対応 -del end

        ''' <summary>
        ''' 汎用用中間ファイルヘッダーチェック(設備用) 20160913_2 部屋設備移行処理の追加
        ''' </summary>
        ''' <param name="tblnamejp"></param>
        ''' <param name="headervalue"></param>
        ''' <param name="list_errheader"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_BaseMidFileHeader_Setubi(ByVal tblnamejp As String, ByVal headervalue As Object, ByRef list_errheader As List(Of String)) As Boolean

            Dim rtn As Boolean = True

            For cntii = 1 To headervalue.Length - 1

                Dim tmp_header As String = tblnamejp & STR_SPLIT_1 & headervalue(0, cntii)

                If headervalue(0, cntii) <> "項目名" Then
                    If List_Basemidheader.Contains(tmp_header) = False Then
                        list_errheader.Add(headervalue(0, cntii))
                    End If
                End If

            Next

            If list_errheader.Count <> 0 Then
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 送金ルール送金先情報の送金先No、送金先口座Noの有無を確認する
        ''' 管理形態が自社物件、送金保留の場合は必須ではない(無くても良い)
        ''' それ以外は必須
        ''' </summary>
        ''' <param name="hash_miditem">中間ファイルのヘッダーと値を紐付けたオブジェクト</param>
        ''' <param name="hash_log">エラー時のログ格納用オブジェクト</param>
        ''' <returns>True…移行可 False…移行不可 ※戻り値は移行可/不可だがここではチェックのみなので実際の移行可/不可はコンバート処理時に判別する</returns>
        ''' <remarks>20161108_2 送金ルール送金先情報の移行制御処理修正 新規追加</remarks>
        Private Function Chk_SoruleSosaki(ByRef hash_miditem As Hashtable, ByRef hash_log As Hashtable) As Boolean

            '送金先情報から必要なデータを取得
            Dim tmp_bkno() As String = hash_miditem("物件No").ToString.Split("-")     'チェック用に(「物件No-部屋No」の形で保持しているため分割)
            Dim sosakibkno As String = tmp_bkno(0)
            Dim sosakihyno As String = hash_miditem("部屋No")
            Dim sosakiknno As String = hash_miditem("送金ルール管理No")
            Dim sosakisono As String = hash_miditem("送金先家主NO")
            Dim sosakisokozano As String = hash_miditem("送金先口座NO")

            '送金ルール情報から管理形態を取得
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT 管理形態 FROM CVTBL_送金ルール基本情報 "
            tmp_sql = tmp_sql & " WHERE [物件No] = '" & sosakibkno & "'"
            tmp_sql = tmp_sql & " AND   [部屋No] = '" & sosakihyno & "'"
            tmp_sql = tmp_sql & " AND   [送金ルール管理No] = '" & sosakiknno & "'"
            Dim soknkei As String = DBExec.Exec_Scalar(tmp_sql, Me.sqlcnnv10)

            '管理形態による送金先No、送金先口座Noの必須チェック
            Dim log_key As String = ""
            Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED

            If soknkei = "4" Or soknkei = "5" Then
                '自社物件、送金保留の物件は必須ではないためチェック処理を行わない
            Else
                '自社物件、送金保留以外は必須チェックを行う
                If sosakisono = "" Then
                    'ログ出力
                    hash_log.Clear()
                    log_key = "送金ルール送金先情報" & STR_SPLIT_1 & "送金先家主NO"
                    hash_log.Add(log_key, log_value)
                    Return False
                ElseIf sosakisokozano = "" Then
                    'ログ出力
                    hash_log.Clear()
                    log_key = "送金ルール送金先情報" & STR_SPLIT_1 & "送金先口座NO"
                    hash_log.Add(log_key, log_value)
                    Return False
                End If
            End If

            Return True

        End Function


        ''' <summary>
        ''' 物件・部屋所有者情報重複チェック
        ''' </summary>
        ''' <param name="hash_miditem">物件所有者情報1レコードを格納したオブジェクト</param>
        ''' <param name="hash_log">ログ格納用オブジェクト</param>
        ''' <remarks>
        ''' 20161124 物件部屋所有者情報が両方存在する場合のチェック機能の追加 新規追加
        ''' 物件所有者情報チェック時に部屋所有者情報と照合して重複した場合にログを出力する
        ''' </remarks>
        Private Sub Chk_BkHySyo(ByRef hash_miditem As Hashtable, ByRef hash_log As Hashtable)

            '物件所有者情報から必要な情報を取得
            Dim tmp_bkno As String = hash_miditem("物件No")
            Dim tmp_knno As String = hash_miditem("管理No")

            '部屋所有者情報に存在するか確認
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT COUNT(*) FROM CVTBL_部屋所有者情報 "
            tmp_sql = tmp_sql & " WHERE [物件No] = '" & tmp_bkno & "'"
            tmp_sql = tmp_sql & " AND   [管理No] = '" & tmp_knno & "'"
            Dim hysyocnt As String = DBExec.Exec_Scalar(tmp_sql, Me.sqlcnnv10)

            'データが存在する場合は区分所有として移行される内容をログに出力する
            '※既に物件Noに関するエラーがある場合はキーのエラーなのでそちらの内容を優先する
            If hysyocnt <> "0" Then

                '20161124 レビュー結果：文言修正「log_err=物件所有者情報と部屋所有者情報へ重複してデータが登録されています。」
                '20161124 レビュー結果：文言修正「log_taisyo=区分所有として移行されます。(部屋所有者情報優先)」
                '20161124 レビュー結果：文言修正「一棟所有の場合は物件所有者情報のみ、区分所有の場合は部屋所有者情報のみに登録を行って下さい。」

                Dim log_key As String = "物件所有者情報" & STR_SPLIT_1 & "物件No"
                If hash_log.Contains(log_key) = False Then
                    '20161125 レビュー指摘事項対応 -chg sta
                    'Dim log_err As String = "物件所有者情報と部屋所有者情報が重複しています。"
                    'Dim log_taisyo As String = "部屋所有者情報が優先され区分所有として移行されます。" & _
                    '                           "一棟所有物件として移行する場合は部屋所有者情報を削除して下さい。"
                    Dim log_err As String = "物件所有者情報と部屋所有者情報へ重複してデータが登録されています。"
                    Dim log_taisyo As String = "区分所有として移行されます。(部屋所有者情報優先)" & _
                                               "一棟所有の場合は物件所有者情報のみ、区分所有の場合は部屋所有者情報のみに登録を行って下さい。"
                    '20161125 レビュー指摘事項対応 -chg end
                    Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & log_err & "-" & log_taisyo
                    hash_log.Add(log_key, log_value)
                End If

            End If

        End Sub

#End Region

#Region "画面制御：共通処理"

        ''' <summary>
        ''' 全移行対象ブロック名取得
        ''' </summary>
        ''' <remarks>
        ''' ・全移行ブロック(物件情報、修繕情報など)を取得
        ''' →必要な移行ブロックは全てコンバーターの対象項目タブにチェックボックスで用意する
        ''' ・グループボックス、チェックボックスの名称を取得
        ''' ※上記より、対処項目タブ内のチェックボックス名を変更する場合、関連する内部処理も修正する
        ''' </remarks>
        Private Sub Set_CVItemName()

            'マスタ系
            model_cvitem.MstBkbrui = Me.chkMstBkrui.Text
            model_cvitem.MstBusKotu = Me.chkMstBusKotu.Text
            model_cvitem.MstBus = Me.chkMstBus.Text
            model_cvitem.MstHyrui = Me.chkMstHyrui.Text
            model_cvitem.MstKagititle = Me.chkMstKagititle.Text
            model_cvitem.MstKasyoClaimrui = Me.chkMstKasyoClaimrui.Text
            model_cvitem.MstTokuyaku = Me.chkMstTokuyaku.Text
            '20160516 契約分類削除に伴うコメントアウト -del
            'model_cvitem.MstKeiyakurui = Me.chkMstKeiyakurui.Text
            model_cvitem.MstHokenrui = Me.chkMstHokenrui.Text
            model_cvitem.MstNkinkbn = Me.chkMstNkinkbn.Text
            model_cvitem.MstSchool = Me.chkMstSchool.Text
            model_cvitem.MstArea = Me.chkMstArea.Text
            model_cvitem.MstKozasyubetu = Me.chkMstKozasyubetu.Text
            model_cvitem.MstKozo = Me.chkMstKozo.Text
            model_cvitem.MstTorihikitaiyo = Me.chkMstTorihikitaiyo.Text
            model_cvitem.MstNkinkomok = Me.chkMstNkinkomok.Text
            model_cvitem.MstSetubi = Me.chkMstSetubi.Text
            model_cvitem.MstHendo = Me.chkMstHendo.Text
            model_cvitem.MstHendoitiran = Me.chkMstHendoitiran.Text
            model_cvitem.MstBikotitle = Me.chkMstBikotitle.Text
            model_cvitem.MstBikolst = Me.chkMstBikolst.Text
            model_cvitem.MstGazotitle = Me.chkMstGazotitle.Text                     '20160801 タイトルマスタ統合処理 -add

            '業者
            model_cvitem.GyCyukaiBase = Me.chkGyCyukaiBase.Text
            model_cvitem.GyCyukaiKoza = Me.chkGyCyukaiKoza.Text
            model_cvitem.GyCyukaiMemo = Me.chkGyCyukaiMemo.Text                     '汎用分(V7には無い)
            model_cvitem.GyHokenBase = Me.chkGyHokenBase.Text
            model_cvitem.GyHokenKoza = Me.chkGyHokenKoza.Text
            model_cvitem.GyHokenMemo = Me.chkGyHokenMemo.Text                       '汎用分(V7には無い)
            model_cvitem.GyYatinhosyoBase = Me.chkGyYatinhosyoBase.Text
            model_cvitem.GyYatinhosyoKoza = Me.chkGyYatinhosyoKoza.Text             '革命10に過去にあったが現在(2015/10/23)は無くなっている
            model_cvitem.GyYatinhosyoMemo = Me.chkGyYatinhosyoMemo.Text             '汎用分(V7には無い)
            model_cvitem.GySyuzenBase = Me.chkGySyuzenBase.Text
            model_cvitem.GySyuzenKoza = Me.chkGySyuzenKoza.Text
            model_cvitem.GySyuzenMemo = Me.chkGySyuzenMemo.Text                     '汎用分(V7には無い)
            model_cvitem.GyLifelineBase = Me.chkGyLifelineBase.Text
            model_cvitem.GySekoBase = Me.chkGySekoBase.Text
            model_cvitem.GySisetuBase = Me.chkGySisetuBase.Text

            '自社情報
            model_cvitem.JisyaBase = Me.chkJisyaBase.Text
            model_cvitem.JisyaKoza = Me.chkJisyaKoza.Text
            model_cvitem.JisyaTanto = Me.chkJisyaTanto.Text
            model_cvitem.JisyaMemo = Me.chkJisyaMemo.Text
            model_cvitem.FBFuriirai = Me.chkFBFuriirai.Text
            model_cvitem.FBFuritesuryo = Me.chkFBFuritesuryo.Text
            model_cvitem.FBKozafurikae = Me.chkFBKozafurikae.Text
            model_cvitem.FBNsSyutoku = Me.chkFBNsSyutoku.Text
            model_cvitem.MstYatinKoza = Me.chkMstYatinKoza.Text
            model_cvitem.MstANSERArea = Me.chkMstANSERArea.Text
            model_cvitem.MstANSERAccpoint = Me.chkMstANSERAccpoint.Text
            model_cvitem.FBANSERSetuzoku = Me.chkFBANSERSetuzoku.Text

            '家主情報
            model_cvitem.OwBase = Me.chkOwBase.Text
            model_cvitem.OwKoza = Me.chkOwKoza.Text
            model_cvitem.OwEvent = Me.chkOwEvent.Text
            model_cvitem.OwMemo = Me.chkOwMemo.Text

            '契約者情報
            model_cvitem.KysBase = Me.chkKysBase.Text
            model_cvitem.KysKoza = Me.chkKysKoza.Text
            model_cvitem.KysMemo = Me.chkKysMemo.Text
            model_cvitem.KysSyogoKana = Me.chkKysSyogoKana.Text
            model_cvitem.KysHosyonin = Me.chkKysHosyonin.Text

            '物件情報
            model_cvitem.BkBase = Me.chkBkBase.Text
            model_cvitem.BkSyosai = Me.chkBkSyosai.Text
            model_cvitem.Bksyo = Me.chkBkSyo.Text
            model_cvitem.BkGomi = Me.chkBkGomi.Text
            model_cvitem.BkKenri = Me.chkBkKenri.Text
            model_cvitem.BkKotu = Me.chkBkKotu.Text
            model_cvitem.BkSetudo = Me.chkBkSetudo.Text
            model_cvitem.BkSyuhen = Me.chkBkSyuhen.Text
            model_cvitem.BkSzeniji = Me.chkBkSzeniji.Text
            model_cvitem.BkMemo = Me.chkBkMemo.Text
            model_cvitem.BkKagi = Me.chkBkKagi.Text                                 '汎用分(V7には無い)   
            model_cvitem.BkHendo = Me.chkBkHendo.Text                               '汎用分(V7には無い)
            model_cvitem.BkKinrincyusyajo = Me.chkBkKinrincyusyajo.Text             '汎用分(V7には無い)
            model_cvitem.BkSansyofile = Me.chkBkSansyofile.Text                     '汎用分(V7には無い)

            '部屋情報
            model_cvitem.HyBase = Me.chkHyBase.Text
            model_cvitem.HySyosai = Me.chkHySyosai.Text
            model_cvitem.Hysyo = Me.chkHySyo.Text
            model_cvitem.HyParking = Me.chkHyParking.Text
            model_cvitem.HyTokuyaku = Me.chkHyTokuyaku.Text
            model_cvitem.HyKagi = Me.chkHyKagi.Text
            model_cvitem.HyMadoriutiwake = Me.chkHyMadoriutiwake.Text
            model_cvitem.HyMenseki = Me.chkHyMenseki.Text
            model_cvitem.HySetubi = Me.chkHySetubi.Text
            model_cvitem.HyNkinkomk = Me.chkHyNkinkomk.Text
            model_cvitem.HyHendo = Me.chkHyHendo.Text
            model_cvitem.HyMemo = Me.chkHyMemo.Text
            model_cvitem.HySzeniji = Me.chkHySzeniji.Text
            model_cvitem.HyCommonsalespoint = Me.chkHyCommonsalespoint.Text         '汎用分(V7には無い)
            model_cvitem.HyConfirm = Me.chkHyConfirm.Text                           '汎用分(V7には無い)
            model_cvitem.HyKenri = Me.chkHyKenri.Text                               '汎用分(V7には無い)
            model_cvitem.HySansyofile = Me.chkHySansyofile.Text                     '汎用分(V7には無い)
            model_cvitem.HyGenjotanka = Me.chkHyGenjotanka.Text                     '汎用分(V7には無い)

            '送金ルール
            model_cvitem.SoruleBase = Me.chkSoruleBase.Text
            model_cvitem.SoruleSosaki = Me.chkSoruleSosaki.Text
            model_cvitem.SoruleNkin = Me.chkSoruleNkin.Text
            model_cvitem.SoruleKojo = Me.chkSoruleKojo.Text

            '契約情報
            model_cvitem.KyBase = Me.chkKyBase.Text
            model_cvitem.KyRireki = Me.chkKyRireki.Text
            model_cvitem.KyCar = Me.chkKyCar.Text
            model_cvitem.KyKys = Me.chkKyKys.Text
            model_cvitem.KyHosyonin = Me.chkKyHosyonin.Text
            model_cvitem.KyNyukyo = Me.chkKyNyukyo.Text
            model_cvitem.KyTokuyaku = Me.chkKyTokuyaku.Text
            model_cvitem.KyHoken = Me.chkKyHoken.Text
            model_cvitem.KyMemo = Me.chkKyMemo.Text
            model_cvitem.KyNkinkomk = Me.chkKyNkinkomk.Text
            model_cvitem.KyNkinkomkNx = Me.chkKyNkinkomkNx.Text
            model_cvitem.KyHendo = Me.chkKyHendo.Text
            model_cvitem.KyKojoRule = Me.chkKyKojoRule.Text
            model_cvitem.KySorule = Me.chkKySorule.Text
            model_cvitem.KyKai = Me.chkKyKai.Text
            '20160531 鍵情報移行処理の修正 -del sta
            ''2016.04.06 契約鍵情報の移行処理追加 -add
            'model_cvitem.KyKagi = Me.chkKyKagi.Text
            '20160531 鍵情報移行処理の修正 -del end
            model_cvitem.KySzen = Me.chkKySzen.Text
            model_cvitem.KySzenmeisai = Me.chkKySzenmeisai.Text

            '請求情報
            model_cvitem.SqKajyo = Me.chkSqKajyo.Text
            model_cvitem.SqUnyotaino = Me.chkSqUnyotaino.Text
            model_cvitem.SqSq = Me.chkSqSq.Text
            model_cvitem.SqHendokensin = Me.chkSqHendokensin.Text
            model_cvitem.SqKoteiKojo = Me.chkSqKoteiKojo.Text
            model_cvitem.SqSqKojo = Me.chkSqSqKojo.Text

            'クレーム情報
            model_cvitem.ClaimBase = Me.chkClaimBase.Text
            model_cvitem.ClaimTaiorireki = Me.chkClaimTaiorireki.Text
            model_cvitem.ClaimRelfile = Me.chkClaimRelfile.Text
            model_cvitem.SzenBase = Me.chkSzenBase.Text
            model_cvitem.SzenSzen = Me.chkSzenSzen.Text
            model_cvitem.SzenSzenmeisai = Me.chkSzenSzenmeisai.Text
            model_cvitem.SzenClaim = Me.chkSzenClaim.Text
            model_cvitem.SzenRelfile = Me.chkSzenRelfile.Text                   	'20160627 修繕関連ファイル移行修正 -add
            model_cvitem.SzenMemo = Me.chkSzenMemo.Text

            '初期設定情報
            model_cvitem.SyskanriBase = Me.chkSyskanriBase.Text
            model_cvitem.SyskanriZei = Me.chkSyskanriZei.Text
            model_cvitem.SyskanriHenkanmoji = Me.chkSyskanriHenkanmoji.Text
            model_cvitem.SyskanriNkinkomkmerge = Me.chkSyskanriNkinkomkmerge.Text

            '物件データ連動情報 													'20160720 連動情報構築 -add
            model_cvitem.RendoSosinBase = Me.chkRendoSosinBase.Text
            model_cvitem.RendoSosinJisyaweb = Me.chkRendoSosinJisyaweb.Text
            model_cvitem.RendoSosinHomes = Me.chkRendoSosinHomes.Text
            model_cvitem.RendoSosinAthome = Me.chkRendoSosinAthome.Text
            model_cvitem.RendoSosinSuumo = Me.chkRendoSosinSuumo.Text
            model_cvitem.RendoMapdisp = Me.chkRendoMapdisp.Text
            model_cvitem.RendoBtoBgroup = Me.chkRendoBtoBgroup.Text
            model_cvitem.RendoHysosin = Me.chkRendoHysosin.Text
            model_cvitem.RendoHyrui = Me.chkRendoHyrui.Text
            model_cvitem.RendoKokokuSuumo = Me.chkRendoKokokuSuumo.Text
            model_cvitem.RendoKokokuAthome = Me.chkRendoKokokuAthome.Text
            model_cvitem.RendoKokokuHomes = Me.chkRendoKokokuHomes.Text
            model_cvitem.RendoKokokuJisyaweb = Me.chkRendoKokokuJisyaweb.Text

        End Sub

        ''' <summary>
        ''' チェックONのチェックボックス数を取得
        ''' </summary>
        ''' <param name="container"></param>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Private Sub Set_ChkBoxCount(ByVal container As Control, ByRef int As Integer)

            For Each item As Control In container.Controls
                If item.GetType().Equals(GetType(CheckBox)) Then
                    If DirectCast(item, CheckBox).Checked Then
                        int = int + 1
                    End If
                End If
            Next

        End Sub

        ''' <summary>
        ''' チェック項目取得→リスト(オブジェクト)へ格納 '20160719 中間ファイルチェック処理時のエラー対応 引数に中間ファイルチェックフラグを追加(Optional)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' ・仕様上、全タブを表示する
        ''' </remarks>
        Private Function Get_ListCVChkitem(Optional ByVal midchkflg As Boolean = False) As List(Of String)

            Dim rtn_list As New List(Of String)
            Dim tmp_grpname As String = ""
            Dim tmp_chkname As String = ""
            '20160720 連動情報構築 -add リストオブジェクトへDC_SELTAB_B11を追加
            '20160707 作業用DB作成時エラー修正 -add
            Dim list_tabpagename As New List(Of String) From _
                {DC_SELTAB_B01, DC_SELTAB_B02, DC_SELTAB_B03, DC_SELTAB_B04, DC_SELTAB_B05, DC_SELTAB_B06, DC_SELTAB_B07, DC_SELTAB_B08, DC_SELTAB_B09, DC_SELTAB_B10, DC_SELTAB_B11}

            '20160707 作業用DB作成時エラー修正 -chg sta
            'もともと作成していたテーブル毎のチェックボックスを見に行く必要があるためtabPageBase110～200の表示処理を追加
            'tabPageKizon110～130のチェックボックスは見ていないのでコメントアウト(残しておくと実行ボタン押下時に一瞬画面が変化する)
            'tabPageManagerCvitem.ChangeTabPageVisible(0, True)
            'tabPageManagerCvitem.ChangeTabPageVisible(1, True)
            'tabPageManagerCvitem.ChangeTabPageVisible(2, True)
            tabPageManagerCvitem.ChangeTabPageVisible(3, True)
            tabPageManagerCvitem.ChangeTabPageVisible(4, True)
            tabPageManagerCvitem.ChangeTabPageVisible(5, True)
            tabPageManagerCvitem.ChangeTabPageVisible(6, True)
            tabPageManagerCvitem.ChangeTabPageVisible(7, True)
            tabPageManagerCvitem.ChangeTabPageVisible(8, True)
            tabPageManagerCvitem.ChangeTabPageVisible(9, True)
            tabPageManagerCvitem.ChangeTabPageVisible(10, True)
            tabPageManagerCvitem.ChangeTabPageVisible(11, True)
            tabPageManagerCvitem.ChangeTabPageVisible(12, True)
            tabPageManagerCvitem.ChangeTabPageVisible(13, True) '20160720 連動情報構築 -add
            '20160707 作業用DB作成時エラー修正 -chg end

            'タブコントロール内
            For Each item As Control In Me.tabCtrlCVItem.Controls
                If item.GetType().Equals(GetType(TabPage)) Then
                    Dim container_tab As TabPage = DirectCast(item, TabPage)

                    If list_tabpagename.Contains(container_tab.Name.Trim) Then
                        'タブページ内
                        For Each item_sub1 As Control In container_tab.Controls
                            If item_sub1.GetType().Equals(GetType(GroupBox)) Then

                                Dim container_grp As GroupBox = DirectCast(item_sub1, GroupBox)

                                '移行グループ名取得
                                tmp_grpname = container_grp.Text

                                'グループボックス内
                                For Each item_sub2 As Control In container_grp.Controls
                                    If item_sub2.GetType().Equals(GetType(CheckBox)) Then
                                        Dim control_chk As CheckBox = DirectCast(item_sub2, CheckBox)
                                        Select Case CNVNO
                                            Case ConvertTypes._既存ユーザ用
                                                If list_baseCVchkitem.Contains(control_chk) = False Then
                                                    If control_chk.Checked Then
                                                        tmp_chkname = control_chk.Text
                                                        'リスト作成
                                                        rtn_list.Add(Trim(tmp_grpname) & "-" & Trim(tmp_chkname))
                                                    End If
                                                End If

                                                Console.WriteLine(Trim(tmp_grpname) & "-" & Trim(tmp_chkname))

                                            Case ConvertTypes._汎用
                                                If list_existCVchkitem.Contains(control_chk) = False Then
                                                    If control_chk.Checked Then
                                                        tmp_chkname = control_chk.Text
                                                        'リスト作成
                                                        rtn_list.Add(Trim(tmp_grpname) & "-" & Trim(tmp_chkname))
                                                        Console.WriteLine(control_chk.Text)
                                                    End If
                                                End If
                                        End Select
                                    End If
                                Next

                            End If
                        Next
                    End If

                End If
            Next

            '※表示状態を戻す
            If midchkflg Then
                Dim seltab As String = Me.tabCtrlCVItem.SelectedTab.Name
                Dim seltabindex As Integer = 0
                Dim seltabindexhid1 As Integer = 0
                Dim seltabindexhid2 As Integer = 0
                Select Case seltab
                    Case "tabPageKizon110"
                        seltabindex = 0
                        seltabindexhid1 = 1
                        seltabindexhid2 = 2
                    Case "tabPageKizon120"
                        seltabindex = 1
                        seltabindexhid1 = 0
                        seltabindexhid2 = 2
                    Case "tabPageKizon130"
                        seltabindex = 2
                        seltabindexhid1 = 0
                        seltabindexhid2 = 1
                End Select
                tabPageManagerCvitem.ChangeTabPageVisible(seltabindex, True)
                tabPageManagerCvitem.ChangeTabPageVisible(seltabindexhid1, False)
                tabPageManagerCvitem.ChangeTabPageVisible(seltabindexhid2, False)
            Else
                tabPageManagerCvitem.ChangeTabPageVisible(0, False)
                tabPageManagerCvitem.ChangeTabPageVisible(1, False)
                tabPageManagerCvitem.ChangeTabPageVisible(2, True)
            End If
            tabPageManagerCvitem.ChangeTabPageVisible(3, False)
            tabPageManagerCvitem.ChangeTabPageVisible(4, False)
            tabPageManagerCvitem.ChangeTabPageVisible(5, False)
            tabPageManagerCvitem.ChangeTabPageVisible(6, False)
            tabPageManagerCvitem.ChangeTabPageVisible(7, False)
            tabPageManagerCvitem.ChangeTabPageVisible(8, False)
            tabPageManagerCvitem.ChangeTabPageVisible(9, False)
            tabPageManagerCvitem.ChangeTabPageVisible(10, False)
            tabPageManagerCvitem.ChangeTabPageVisible(11, False)
            tabPageManagerCvitem.ChangeTabPageVisible(12, False)



            Console.WriteLine("dddd")
            For Each ee In rtn_list
                Console.WriteLine(ee)
            Next
            Console.WriteLine("huju")

            Return rtn_list

        End Function

        ''' <summary>
        ''' ボタン状態変更処理
        ''' </summary>
        ''' <param name="status">設定条件文字列</param>
        ''' <remarks></remarks>
        Private Sub Chg_BtnKariStatus(ByVal status As String)

            btnBack.Visible = True
            btnNext.Visible = True
            btnEnd.Visible = True
            btnBack.Enabled = True
            btnNext.Enabled = True
            btnEnd.Enabled = True

            Select Case status

                '--- メイン画面 ---
                Case "mstart"
                    btnBack.Text = ""
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   終  了"
                    btnBack.Visible = False
                Case "msession"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   終  了"
                    If Me.btnConnectTest.Enabled Then
                        btnNext.Enabled = False
                    End If
                Case "msyoki"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   終  了"
                Case "mmenu"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   終  了"
                    btnNext.Enabled = False

                    '--- 作業選択 ---
                Case "jizen1"
                    btnBack.Text = ""
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False
                    btnNext.Visible = True
                    btnEnd.Visible = False
                Case "jizen2"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = True
                    btnNext.Visible = True
                    btnEnd.Visible = False
                Case "jizen3"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = True
                    btnNext.Visible = True
                    btnEnd.Visible = False
                Case "jizen4"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "作業選択へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = True
                    btnNext.Visible = True
                    btnEnd.Visible = False
                    
                Case "jigo1"
                    btnBack.Text = ""
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False
                    btnNext.Visible = True
                    btnEnd.Visible = False
                Case "jigo2"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "作業選択へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = True
                    btnNext.Visible = True
                    btnEnd.Visible = False

                Case "hojyo1"
                    btnBack.Text = ""
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False
                    btnNext.Visible = True
                    btnEnd.Visible = False
                Case "hojyo2"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "作業選択へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = True
                    btnNext.Visible = True
                    btnEnd.Visible = False

                    '--- 作業選択(汎用) ---
                Case "jizenh"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "作業選択へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False
                    btnNext.Visible = True
                    btnEnd.Visible = False
                Case "jigoh"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "作業選択へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False
                    btnNext.Visible = True
                    btnEnd.Visible = False

                    '--- コンバート画面 ---
                Case "dchajimeni"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False
                    btnNext.Enabled = Not btnDoui.Enabled
                Case "dcselect1"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                Case "dcselect2"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                Case "dcselect3"
                    '※ボタン押下イベント(btnNext_Click)にて、ボタン名が「実行」の場合にメイン処理を実行(ボタン名変更不可)
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   実  行"
                    btnEnd.Text = "   中  止"
                Case "dcjikko1"
                    '※ボタン押下イベント(btnEnd_Click)にて、ボタン名が「キャンセル」の場合に終了処理を実行(ボタン名変更不可)
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   実行中"
                    btnEnd.Text = "キャンセル"
                    btnBack.Enabled = False
                    btnNext.Enabled = False
                Case "dcjikko2"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False
                    btnEnd.Visible = False
                Case "dcend"
                    btnBack.Text = ""
                    btnNext.Text = " ログを開く"
                    btnEnd.Text = "作業選択へ"
                    btnBack.Visible = False

                    '--- 中間ファイルチェック ---
                Case "midstop"
                    btnMidFileCheck.Text = " キャンセル"
                    btnBack.Enabled = False
                    btnNext.Enabled = False
                    btnEnd.Enabled = False
                    btnAllChk.Enabled = False
                Case "midcheck"
                    btnMidFileCheck.Text = " チェック"
                    btnBack.Enabled = True
                    btnNext.Enabled = True
                    btnEnd.Enabled = True
                    btnAllChk.Enabled = True

                    '--- 99.その他 ---
                Case "etc"
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   終  了"
                    btnBack.Visible = False
                    btnNext.Visible = False
                    btnEnd.Visible = False

                Case Else

            End Select

        End Sub

        ''' <summary>
        ''' セットフォーカス
        ''' </summary>
        ''' <param name="ctrlmoji">セット用文字列</param>
        ''' <remarks>
        ''' ・イベントやコントロール遷移時にフォーカスをあてる
        ''' </remarks>
        Private Sub Set_Forcus(ByVal ctrlmoji As String)

            Select Case ctrlmoji

                '--- メイン画面 ---
                Case "prev"                             '[戻る]ボタンにセット
                    btnBack.Focus()
                Case "next"                             '[次へ]ボタンにセット
                    btnNext.Focus()
                Case "end"                              '[終了]ボタンにセット]
                    btnEnd.Focus()

                    '--- 接続設定 ---
                Case "testset"                          '接続設定.[接続テスト]ボタンへセット
                    btnConnectTest.Focus()

                    '--- 作業選択 ---
                Case "menu_jizen"
                    btnMenuJizen.Focus()                '作業選択.[事前作業]ボタンへセット
                Case "menu_datacv"                      '作業選択.[データコンバート]ボタンへセット
                    btnMenuDatacv.Focus()
                Case "menu_gazocv"                      '作業選択.[画像コンバート]ボタンへセット
                    btnMenuGazocv.Focus()
                Case "menu_jigo"                        '作業選択.[事後作業]ボタンへセット
                    btnMenuJigo.Focus()
                Case "menu_hojyo"                       '作業選択.[補助機能]ボタンへセット
                    btnMenuHojyo.Focus()

                    '--- 事前作業1 ---
                Case "jizen_kai"
                    chkJizenKai.Focus()
                Case "jizen_azu"
                    chkJizenAzu.Focus()
                Case "jizen_hurikae"
                    chkJizenHurikae.Focus()
                Case "jizen_so"
                    chkJizenSo.Focus()
                    '--- 事前作業2 ---
                Case "jizen_minus"
                    chkJizenMinus.Focus()
                Case "jizen_bkhourei"
                    chkJizenBkhourei.Focus()
                Case "jizen_hasseiow"
                    chkJizenHasseiOw.Focus()
                Case "jizen_hasseihen"
                    chkJizenHasseiHen.Focus()
                    '--- 事前作業3 ---
                Case "jizen_yanuso"
                    chkJizenYanuso.Focus()
                Case "jizen_szenkyshutan"
                    chkJizenSzenKyshutan.Focus()
                Case "jizen_kagi"
                    chkJizenKagi.Focus()
                    '-- 事前作業4 --
                Case "jizen_nonjisyakoza"
                    chkJizenNonJisyaKoza.Focus()
                Case "jizen_simesokin"     '20160829 事前作業_送金予定日リスト出力機能を追加 -add
                    chkJizenSimeSokin.Focus()
                Case "jizen_kozameigikana"
                    chkJizenKozameigikana.Focus()

                    '--- 事前作業(汎用)
                Case "jizen_keisiki"
                    chkHJizenGazoKeisiki.Focus()
                Case "jizen_tyukan"
                    chkHJizenTyukan.Focus()

                    '--- コンバーター ---
                Case "dcdoui"                           'はじめに.[同意]ボタンにセット
                    btnDoui.Focus()

                Case Else

            End Select

        End Sub

        ''' <summary>
        ''' 移行項目から紐付項目をセット → コマンドライン用に成形
        ''' </summary>
        ''' <param name="list_cvitem"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_CVItemChkToCmdLine(ByVal list_cvitem As List(Of String)) As String

            Dim tmp_bkrui As String = ""
            Dim tmp_hyrui As String = ""
            Dim tmp_nkinkbn As String = ""
            Dim tmp_toritaiyo As String = ""
            Dim tmp_nkinkomk As String = ""
            Dim tmp_kozo As String = ""
            Dim tmp_kozasyu As String = ""
            Dim tmp_setubi As String = ""
            Dim tmp_kagi As String = ""
            Dim tmp_jisyakoza As String = ""
            Dim tmp_bktosiyoto As String = ""
            Dim tmp_kyrui As String = ""
            Dim tmp_sohoken As String = ""
            Dim tmp_gazo As String = ""
            Dim tmp_fbfmt As String = ""
            Dim rtn_str As String = ""


            For Each item In list_cvitem

                Dim tmp_item() As String = item.Split("-")
                Dim tmp_grpname As String = tmp_item(0)
                Dim tmp_komkname As String = tmp_item(1)

                '20161012 紐付ツール呼出起動時の項目過不足修正 -chg sta
                'グループ名から判別可能なものを取得
                'Select Case tmp_grpname
                '    Case "物件情報"
                '        tmp_bkrui = "bkruichk"
                '        tmp_kozo = "kozochk"
                '        tmp_kagi = "kagichk"
                '        tmp_bktosiyoto = "tosiyotochk"
                '        tmp_gazo = "gazochk"
                '    Case "部屋情報"
                '        tmp_hyrui = "hyruichk"
                '        tmp_kagi = "kagichk"
                '        tmp_kyrui = "kyruichk"
                '        tmp_nkinkbn = "nkinkbnchk"
                '        tmp_toritaiyo = "toritaiyochk"
                '        tmp_nkinkomk = "nkinkomkchk"
                '        tmp_setubi = "setubichk"
                '        tmp_sohoken = "sohokenchk"
                '        tmp_gazo = "gazochk"
                '    Case "契約情報"
                '        tmp_kagi = "kagichk"
                '        tmp_kyrui = "kyruichk"
                '        tmp_nkinkbn = "nkinkbnchk"
                '        tmp_nkinkomk = "nkinkomkchk"
                '        tmp_sohoken = "sohokenchk"
                '    Case "送金ルール情報"
                '        tmp_nkinkbn = "nkinkbnchk"
                '        tmp_nkinkomk = "nkinkomkchk"
                '        tmp_sohoken = "sohokenchk"
                '    Case Else
                '        'グループ名から判別できないものを項目名から取得
                '        Select Case True
                '            '※メインの方へも反映させる
                '            Case tmp_komkname = "自社口座情報" Or tmp_komkname = "振込依頼人情報" Or tmp_komkname = "口座振替情報"
                '                tmp_jisyakoza = "jisyakozachk"
                '                tmp_nkinkbn = "nkinkbnchk"
                '                tmp_kozasyu = "kozasyuchk"
                '                tmp_fbfmt = "fbfmtchk"
                '            Case tmp_komkname <> tmp_komkname.Replace("口座情報", "")
                '                tmp_nkinkbn = "nkinkbnchk"
                '                tmp_kozasyu = "kozasyuchk"
                '            Case tmp_komkname <> tmp_komkname.Replace("入金項目情報", "")
                '                tmp_nkinkomk = "nkinkomkchk"
                '            Case tmp_komkname <> tmp_komkname.Replace("変動費", "")
                '                tmp_nkinkomk = "nkinkomkchk"
                '            Case tmp_komkname = "鍵タイトルマスタ"
                '                tmp_kagi = "kagichk"
                '            Case Else

                '        End Select
                'End Select
                Select Case True
                    '※メインの方へも反映させる
                    Case tmp_komkname = "自社口座情報" Or tmp_komkname = "振込依頼人情報" Or tmp_komkname = "口座振替情報"
                        tmp_jisyakoza = "jisyakozachk"
                        ''20161025 紐付不要項目の除去 -del
                        'tmp_nkinkbn = "nkinkbnchk"
                        tmp_kozasyu = "kozasyuchk"
                        tmp_fbfmt = "fbfmtchk"
                    Case tmp_komkname <> tmp_komkname.Replace("口座情報", "")
                        ''20161025 紐付不要項目の除去 -del
                        'tmp_nkinkbn = "nkinkbnchk"
                        tmp_kozasyu = "kozasyuchk"
                    Case tmp_komkname <> tmp_komkname.Replace("入金項目情報", "")
                        tmp_nkinkomk = "nkinkomkchk"
                    Case tmp_komkname <> tmp_komkname.Replace("変動費", "")
                        tmp_nkinkomk = "nkinkomkchk"
                    Case tmp_komkname = "鍵タイトルマスタ"
                        tmp_kagi = "kagichk"
                    Case tmp_komkname = "部屋設備情報"
                        tmp_setubi = "setubichk"
                    Case Else
                        Select Case tmp_grpname
                            Case "物件情報"
                                tmp_bkrui = "bkruichk"
                                tmp_kozo = "kozochk"
                                tmp_kagi = "kagichk"
                                tmp_bktosiyoto = "tosiyotochk"
                                tmp_gazo = "gazochk"
                            Case "部屋情報"
                                tmp_hyrui = "hyruichk"
                                tmp_kagi = "kagichk"
                                tmp_kyrui = "kyruichk"
                                tmp_nkinkbn = "nkinkbnchk"
                                tmp_toritaiyo = "toritaiyochk"
                                tmp_nkinkomk = "nkinkomkchk"
                                tmp_sohoken = "sohokenchk"
                                tmp_gazo = "gazochk"
                            Case "契約情報"
                                tmp_kagi = "kagichk"
                                tmp_kyrui = "kyruichk"
                                tmp_nkinkbn = "nkinkbnchk"
                                tmp_nkinkomk = "nkinkomkchk"
                                tmp_sohoken = "sohokenchk"
                            Case "送金ルール情報"
                                tmp_nkinkbn = "nkinkbnchk"
                                tmp_nkinkomk = "nkinkomkchk"
                                tmp_sohoken = "sohokenchk"
                            Case "請求情報"
                                tmp_nkinkbn = "nkinkbnchk"
                                tmp_nkinkomk = "nkinkomkchk"
                        End Select
                End Select
                '20161012 紐付ツール呼出起動時の項目過不足修正 -chg end
            Next
            '20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_本体 -add sta
            '汎用では紐付を行わない項目を除去
            If CNVNO = ConvertTypes._汎用 Then
                tmp_setubi = ""
                tmp_jisyakoza = ""
                tmp_gazo = ""
                tmp_fbfmt = ""
                tmp_bktosiyoto = ""  '20161012 汎用時に紐付対象外となる項目非表示修正 -add
                tmp_kagi = ""        '20161025 紐付不要項目の除去 -add
            End If
            '20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_本体 -add end
            'コマンドライン用に半角スペースで結合
            rtn_str = tmp_bkrui & IIf(tmp_bkrui = "", "", "-") & _
                      tmp_hyrui & IIf(tmp_hyrui = "", "", "-") & _
                      tmp_nkinkbn & IIf(tmp_nkinkbn = "", "", "-") & _
                      tmp_toritaiyo & IIf(tmp_toritaiyo = "", "", "-") & _
                      tmp_nkinkomk & IIf(tmp_nkinkomk = "", "", "-") & _
                      tmp_kozo & IIf(tmp_kozo = "", "", "-") & _
                      tmp_kozasyu & IIf(tmp_kozasyu = "", "", "-") & _
                      tmp_setubi & IIf(tmp_setubi = "", "", "-") & _
                      tmp_kagi & IIf(tmp_kagi = "", "", "-") & _
                      tmp_jisyakoza & IIf(tmp_jisyakoza = "", "", "-") & _
                      tmp_bktosiyoto & IIf(tmp_bktosiyoto = "", "", "-") & _
                      tmp_kyrui & IIf(tmp_kyrui = "", "", "-") & _
                      tmp_gazo & IIf(tmp_gazo = "", "", "-") & _
                      tmp_fbfmt & IIf(tmp_fbfmt = "", "", "-") & _
                      tmp_sohoken

            '20160829 紐付ツール起動位置修正_本体→紐付 -add sta
            '※強制でコンバートPG本体と同位置にセット
            If rtn_str <> "" Then
                Dim tmp_strcnt As Integer = rtn_str.Length
                If rtn_str.Substring(tmp_strcnt - 1) = "-" Then
                    rtn_str = rtn_str.Remove(tmp_strcnt - 1)
                End If
            End If
            '20160829 紐付ツール起動位置修正_本体→紐付 -add end

            Return rtn_str

        End Function

        ''' <summary>
        ''' 項目選択1～3の全チェックボックス名を取得→リスト(オブジェクト)へ格納 '20160707 各抽出件数の出力処理追加
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Private Function Get_ListCVChkitemAll(Optional ByVal chkonoffflg As Boolean = False) As List(Of String)

            Dim rtn_list As New List(Of String)
            Dim tmp_grpname As String = ""
            Dim tmp_chkname As String = ""

            tabPageManagerCvitem.ChangeTabPageVisible(0, True)
            tabPageManagerCvitem.ChangeTabPageVisible(1, True)
            tabPageManagerCvitem.ChangeTabPageVisible(2, True)

            'タブコントロール内
            For Each item As Control In Me.tabCtrlCVItem.Controls
                If item.GetType().Equals(GetType(TabPage)) Then
                    Dim container_tab As TabPage = DirectCast(item, TabPage)

                    'タブページ内
                    For Each item_sub1 As Control In container_tab.Controls
                        If item_sub1.GetType().Equals(GetType(GroupBox)) Then

                            Dim container_grp As GroupBox = DirectCast(item_sub1, GroupBox)
                            '移行グループ名取得
                            tmp_grpname = container_grp.Text

                            'グループボックス内
                            For Each item_sub2 As Control In container_grp.Controls
                                If item_sub2.GetType().Equals(GetType(Panel)) Then

                                    Dim control_panel As Panel = DirectCast(item_sub2, Panel)

                                    For Each item_sub3 As Control In control_panel.Controls

                                        If item_sub3.GetType().Equals(GetType(CheckBox)) Then

                                            Dim control_chk As CheckBox = DirectCast(item_sub3, CheckBox)
                                            '20161014 着色処理の修正 -chg sta
                                            'If CNVNO = ConvertTypes._既存ユーザ用 And list_baseCVchkitem.Contains(control_chk) = False Then

                                            '    If chkonoffflg = False Then

                                            '        '移行項目取得
                                            '        tmp_chkname = control_chk.Text

                                            '        'リスト作成
                                            '        rtn_list.Add(tmp_chkname.Trim)

                                            '    ElseIf control_chk.Checked Then

                                            '        '移行項目取得
                                            '        tmp_chkname = control_chk.Text

                                            '        'リスト作成
                                            '        rtn_list.Add(tmp_chkname.Trim)

                                            '    End If

                                            'End If
                                            'コンバート実績あり項目の着色設定
                                            Select Case CNVNO
                                                Case ConvertTypes._既存ユーザ用
                                                    If list_baseCVchkitem.Contains(control_chk) = False Then
                                                        If chkonoffflg = False Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'リスト作成
                                                            rtn_list.Add(tmp_chkname.Trim)

                                                        ElseIf control_chk.Checked Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'リスト作成
                                                            rtn_list.Add(tmp_chkname.Trim)

                                                        End If
                                                    End If
                                                Case ConvertTypes._汎用
                                                    If list_existCVchkitem.Contains(control_chk) = False Then
                                                        If chkonoffflg = False Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'リスト作成
                                                            rtn_list.Add(tmp_chkname.Trim)

                                                        ElseIf control_chk.Checked Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'リスト作成
                                                            rtn_list.Add(tmp_chkname.Trim)

                                                        End If
                                                    End If
                                            End Select
                                            '20161014 着色処理の修正 -chg end
                                        End If

                                    Next

                                End If
                            Next

                        End If
                    Next

                End If
            Next

            '※表示状態を戻す
            tabPageManagerCvitem.ChangeTabPageVisible(0, True)
            tabPageManagerCvitem.ChangeTabPageVisible(1, False)
            tabPageManagerCvitem.ChangeTabPageVisible(2, False)

            Return rtn_list

        End Function

        ''' <summary>
        ''' 項目選択1～3の全チェックボックスをキャプションとセットで取得→ハッシュテーブルへ(オブジェクト)へ格納 '20160707 コンバート実績保持の処理追加
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Private Function Get_Hash_CVChkitemAll(Optional ByVal chkonoffflg As Boolean = False) As Hashtable

            Dim rtn_hash As New Hashtable
            Dim tmp_grpname As String = ""
            Dim tmp_chkname As String = ""

            tabPageManagerCvitem.ChangeTabPageVisible(0, True)
            tabPageManagerCvitem.ChangeTabPageVisible(1, True)
            tabPageManagerCvitem.ChangeTabPageVisible(2, True)

            'タブコントロール内
            For Each item As Control In Me.tabCtrlCVItem.Controls
                If item.GetType().Equals(GetType(TabPage)) Then
                    Dim container_tab As TabPage = DirectCast(item, TabPage)

                    'タブページ内
                    For Each item_sub1 As Control In container_tab.Controls
                        If item_sub1.GetType().Equals(GetType(GroupBox)) Then

                            Dim container_grp As GroupBox = DirectCast(item_sub1, GroupBox)
                            '移行グループ名取得
                            tmp_grpname = container_grp.Text

                            'グループボックス内
                            For Each item_sub2 As Control In container_grp.Controls
                                If item_sub2.GetType().Equals(GetType(Panel)) Then

                                    Dim control_panel As Panel = DirectCast(item_sub2, Panel)

                                    For Each item_sub3 As Control In control_panel.Controls
                                        If item_sub3.GetType().Equals(GetType(CheckBox)) Then

                                            Dim control_chk As CheckBox = DirectCast(item_sub3, CheckBox)
                                            '20161014 着色処理の修正 -chg sta
                                            'If CNVNO = ConvertTypes._既存ユーザ用 And list_baseCVchkitem.Contains(control_chk) = False Then

                                            '    If chkonoffflg = False Then

                                            '        '移行項目取得
                                            '        tmp_chkname = control_chk.Text

                                            '        'ハッシュテーブル作成
                                            '        If rtn_hash.Contains(tmp_chkname) = False Then
                                            '            rtn_hash.Add(tmp_chkname, control_chk)
                                            '        End If

                                            '    ElseIf control_chk.Checked Then

                                            '        '移行項目取得
                                            '        tmp_chkname = control_chk.Text

                                            '        'ハッシュテーブル作成
                                            '        If rtn_hash.Contains(tmp_chkname) = False Then
                                            '            rtn_hash.Add(tmp_chkname, control_chk)
                                            '        End If

                                            '    End If

                                            'End If
                                            'コンバート実績あり項目の着色設定
                                            Select Case CNVNO
                                                Case ConvertTypes._既存ユーザ用
                                                    If list_baseCVchkitem.Contains(control_chk) = False Then
                                                        If chkonoffflg = False Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'ハッシュテーブル作成
                                                            If rtn_hash.Contains(tmp_chkname) = False Then
                                                                rtn_hash.Add(tmp_chkname, control_chk)
                                                            End If

                                                        ElseIf control_chk.Checked Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'ハッシュテーブル作成
                                                            If rtn_hash.Contains(tmp_chkname) = False Then
                                                                rtn_hash.Add(tmp_chkname, control_chk)
                                                            End If

                                                        End If
                                                    End If
                                                Case ConvertTypes._汎用
                                                    If list_existCVchkitem.Contains(control_chk) = False Then
                                                        If chkonoffflg = False Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'ハッシュテーブル作成
                                                            If rtn_hash.Contains(tmp_chkname) = False Then
                                                                rtn_hash.Add(tmp_chkname, control_chk)
                                                            End If

                                                        ElseIf control_chk.Checked Then

                                                            '移行項目取得
                                                            tmp_chkname = control_chk.Text

                                                            'ハッシュテーブル作成
                                                            If rtn_hash.Contains(tmp_chkname) = False Then
                                                                rtn_hash.Add(tmp_chkname, control_chk)
                                                            End If

                                                        End If
                                                    End If
                                            End Select
                                            '20161014 着色処理の修正 -chg end
                                        End If
                                    Next

                                End If
                            Next

                        End If
                    Next

                End If
            Next

            '※表示状態を戻す
            tabPageManagerCvitem.ChangeTabPageVisible(0, True)
            tabPageManagerCvitem.ChangeTabPageVisible(1, False)
            tabPageManagerCvitem.ChangeTabPageVisible(2, False)

            Return rtn_hash

        End Function

        ''' <summary>
        ''' 任意タブ内コントロールを別タブへ移動
        ''' </summary>
        ''' <param name="befpage">移動元タブページ</param>
        ''' <param name="afpage">移動先タブページ</param>
        ''' <param name="ctrl">移動するコントロール</param>
        ''' <remarks></remarks>
        Sub Chg_CtrlInTabLocation(ByVal befpage As TabPage, ByVal afpage As TabPage, ByVal ctrl As Control)
            befpage.Controls.Remove(ctrl)                   '指定コントロールを元ページから除去
            afpage.Controls.Add(ctrl)                       '指定コントロールを移動先へ登録
        End Sub

#End Region

#Region "画面制御：処理状況画面処理"

        ''' <summary>
        ''' 処理状況をテキストボックスへ出力(画面表示)<br/>
        ''' </summary>
        Private Sub Set_Situation(ByVal value As String, ByVal type As Integer)

            Dim tmp_str As String

            Select Case type
                Case 0  '処理状況
                    tmp_str = Me.txtTotalSituation.Text
                    value = Now & "  " & value & vbCrLf
                    Me.txtTotalSituation.Text = tmp_str & value
                Case 1  '移行完了項目
                    tmp_str = Me.txtPartialSituation.Text
                    value = value & vbCrLf
                    Me.txtPartialSituation.Text = tmp_str & value
            End Select

            'スクロールを下に移動
            Me.txtTotalSituation.SelectionStart = txtTotalSituation.Text.Length
            Me.txtTotalSituation.ScrollToCaret()

            Me.txtPartialSituation.SelectionStart = txtPartialSituation.Text.Length
            Me.txtPartialSituation.ScrollToCaret()

        End Sub

        ''' <summary>
        ''' 論理値を文字列へ変換
        ''' </summary>
        ''' <param name="flg">True.正常, False.異常</param>         
        ''' <param name="typeno">1.個別進捗出力用, 2.全体進捗出力用</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chg_FlgToStr(ByVal flg As Boolean, ByVal typeno As Integer) As String

            Dim rtn_str As String = ""

            Select Case typeno

                Case 1
                    If flg Then
                        rtn_str = SITUATION_CVITEMSUCCESS
                    Else
                        rtn_str = SITUATION_CVITEMFAILURE
                    End If
                Case 2
                    If flg Then
                        rtn_str = LOG_NAIYO_NORMALEND
                    Else
                        rtn_str = LOG_NAIYO_NOTNORMALEND
                    End If
            End Select

            Return rtn_str

        End Function

        ''' <summary>
        ''' 移行項目の数を元にメッセージを表示　　　'20160912 既存用→汎用用中間ファイルへのコピー処理 開発用引数(existmidtobasemidflg)を追加
        ''' </summary>
        ''' <param name="list"></param>
        ''' <param name="midchkbtnflg"></param>
        ''' <returns>True.実行可, False.実行不可</returns>
        ''' <remarks></remarks>
        Private Function Chk_CVStartFlg(ByRef list As List(Of String), Optional ByVal midchkbtnflg As Boolean = False) As Boolean

            Dim tmp_msg As String = ""
            Dim tmp_list As New List(Of String)

            'チェックON項目を格納したリストを取得 (グループ-チェックボックス名)
            tmp_list = Me.Get_ListCVChkitem(midchkbtnflg)

            '移行項目チェック
            If tmp_list.Count = 0 Then
                tmp_msg = tmp_msg & vbCrLf & MSG_ERR_CHK_CVITEM
            End If

            '紐付ファイル格納先フォルダ有無確認
            Dim tmp_reldir As String = Me.txtRelationDirPath.Text
            If EtcMethod.Chk_DirExist(tmp_reldir) Then
                RelationDirPath = tmp_reldir
            Else
                tmp_msg = tmp_msg & vbCrLf & MSG_ERR_DIR_REL
            End If

            'ログファイル格納先フォルダ有無確認
            Dim tmp_logdir As String = Me.txtLogDirPath.Text
            If EtcMethod.Chk_DirExist(tmp_logdir) Then
                Dim logfilename As String = LOG_TMP_TABLENAME & "_" & Replace((Replace(Replace((String.Format(Now)), ":", ""), "/", "")), " ", "") & ".csv"
                LogFilePath = EtcMethod.Set_Path(tmp_logdir, logfilename)
            Else
                tmp_msg = tmp_msg & vbCrLf & MSG_ERR_DIR_LOG
            End If

            '中間ファイル格納先フォルダ有無確認
            Dim tmp_existmiddir As String = EtcMethod.Set_Path(dcv_exedir, DIR_MID_NAME)    '既存用中間ファイル格納先
            Dim tmp_basemidfile As String = Me.txtMidDirPath.Text                           '汎用用中間ファイル格納先

            '既存用中間ファイル格納先有無確認
            If EtcMethod.Chk_DirExist(tmp_existmiddir) Then
                MiddleDirPath = tmp_existmiddir
            Else
                tmp_msg = tmp_msg & vbCrLf & MSG_ERR_DIR_MID
            End If

            '汎用用中間ファイル有無確認
            If CNVNO = ConvertTypes._汎用 Then
                If EtcMethod.Chk_FileExist(tmp_basemidfile) Then
                    BaseMidDirPath = Path.GetDirectoryName(tmp_basemidfile)
                    CV_FROM_MIDDLE = Path.GetFileNameWithoutExtension(tmp_basemidfile)
                Else
                    tmp_msg = tmp_msg & vbCrLf & MSG_ERR_DIR_MID
                End If
            End If
          
            '中間ファイルログ格納先フォルダ有無確認
            Dim tmp_midlogdir As String = Me.txtMidDirLogPath.Text
            If EtcMethod.Chk_DirExist(tmp_midlogdir) Then
                Dim midlogfilename As String = LOG_TMP_MIDCHKTABLENAME & "_" & Replace((Replace(Replace((String.Format(Now)), ":", ""), "/", "")), " ", "") & ".csv"
                MiddleLogFilePath = EtcMethod.Set_Path(tmp_midlogdir, midlogfilename)
            Else
                tmp_msg = tmp_msg & vbCrLf & MSG_ERR_DIR_MIDLOG
            End If


            '--------------------------------------------------
            ' ファイルオープンチェック処理
            '--------------------------------------------------
            If tmp_msg = "" Then
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                'Dim list_chkzumi As New List(Of String) 
                Dim openfilename As String = ""
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                '既存用中間ファイルのオープン確認
                'For Each filesheetname In tmp_list
                '    'ファイル名、ファイルパスを取得
                '    Dim tmp_str() As String = filesheetname.Split("-")
                '    Dim filename As String = tmp_str(0) & ".xlsx"
                '    Dim filepath As String = EtcMethod.Set_Path(MiddleDirPath, filename)

                '    'ファイルオープンチェック
                '    If list_chkzumi.Contains(filepath) = False Then
                '        Dim fileopenflg As Boolean = EtcMethod.Chk_FileOpen(filepath)
                '        If fileopenflg = False Then
                '            openfilename = openfilename & vbCrLf & INDENT_0 & "・" & filename
                '        End If
                '        list_chkzumi.Add(filepath)
                '    End If
                'Next
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                '汎用用中間ファイルのオープン確認
                If CNVNO = ConvertTypes._汎用 Then
                    Dim basemidfilepath As String = EtcMethod.Set_Path(BaseMidDirPath, CV_FROM_MIDDLE & ".xlsx")
                    Dim fileopenflg As Boolean = EtcMethod.Chk_FileOpen(basemidfilepath)
                    If fileopenflg = False Then
                        openfilename = openfilename & vbCrLf & INDENT_0 & "・" & CV_FROM_MIDDLE & ".xlsx"
                    End If
                End If

                '開かれているファイルが存在する場合
                If openfilename <> "" Then
                    tmp_msg = tmp_msg & vbCrLf & MSG_ERR_FILEOPEN
                    tmp_msg = tmp_msg & vbCrLf & "格納先：" & MiddleDirPath
                    tmp_msg = tmp_msg & openfilename
                End If
            End If

            '20161018 フォルダ自動生成箇所の修正 -del sta
            ''--------------------------------------------------
            '' フォルダ自動生成処理
            ''--------------------------------------------------
            'Dim tmp_middirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_MID_NAME)             'コンバーター実行ファイル格納フォルダ
            'Dim tmp_reldirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_RELEXEDIR_NAME)       '紐付ツール実行ファイル格納フォルダ
            'Dim tmp_gazodirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_GAZOEXEDIR_NAME)     '画像CVツール実行ファイル格納フォルダ
            'Dim list_makedir As New List(Of String)

            ''コンバーターフォルダ同階層
            'list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_MAINLOG_NAME))
            'list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_INI_NAME))
            'list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_LIST_NAME))
            ''中間ファイルフォルダ
            'If EtcMethod.Chk_DirExist(tmp_middirpath) Then
            '    list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, DIR_TEMP_CSV))
            '    list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, DIR_MIDLOG_NAME))
            'End If
            ''紐付ツールフォルダ
            'If EtcMethod.Chk_DirExist(tmp_reldirpath) Then
            '    list_makedir.Add(EtcMethod.Set_Path(tmp_reldirpath, DIR_RELLOG_NAME))
            'End If
            ''画像CVツールフォルダ
            'If EtcMethod.Chk_DirExist(tmp_gazodirpath) Then
            '    list_makedir.Add(EtcMethod.Set_Path(tmp_gazodirpath, DIR_INI_NAME))
            'End If

            ''自動生成処理 
            'For Each chkdir In list_makedir
            '    If System.IO.Directory.Exists(chkdir) = False Then
            '        System.IO.Directory.CreateDirectory(chkdir)
            '    End If
            'Next
            '20161018 フォルダ自動生成箇所の修正 -del end

            '--------------------------------------------------
            ' 運用開始年月の追加処理
            '--------------------------------------------------
            Dim tmp_unyoymd As String = Me.txtUnyoYYYYMM.Text
            tmp_unyoymd = tmp_unyoymd.Replace("/", "").Trim
            If tmp_unyoymd = "" Then
                tmp_msg = tmp_msg & vbCrLf & MSG_ERR_UNYOYMD
            Else
                UnyoYMD = tmp_unyoymd & "01"
            End If


            '--------------------------------------------------
            ' 中間ファイルチェック or コンバート実行 ダイアログ出力
            '--------------------------------------------------
            'メッセージ表示
            If tmp_msg <> "" Then
                tmp_msg = tmp_msg.Remove(0, 1)
                MsgResult = MessageBox.Show(tmp_msg, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            End If

            '実行確認
            Dim tmp_msgsta As String = ""

            If midchkbtnflg Then
                tmp_msgsta = MSG_STA_C
            Else
                tmp_msgsta = MSG_STA_A
                '20161014 改善対応：コンバート処理前に中間ファイルチェックを促す -add sta
                If CNVNO = ConvertTypes._汎用 Then
                    tmp_msgsta = tmp_msgsta & vbCrLf & MSG_STA_D
                End If
                '20161014 改善対応：コンバート処理前に中間ファイルチェックを促す -add end
            End If
            MsgResult = MessageBox.Show(tmp_msgsta, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)

            If MsgResult = DialogResult.No Then
                Set_Forcus("next")
                Return False
            End If

            '実行時は移行項目を格納
            list = tmp_list

            '20161028 物件/部屋鍵取得方法修正 -add sta
            '汎用CVの場合、鍵タイトルをプログラム内で自動実行するためにオブジェクトへ鍵タイトルを追加する
            If CNVNO = ConvertTypes._汎用 Then
                list.Add("各マスタ情報-鍵タイトルマスタ")
            End If
            '20161028 物件/部屋鍵取得方法修正 -add end

            Return True

        End Function

#End Region

#Region "画面制御：チェックボックス処理・ボタン状態処理"

        ''' <summary>
        ''' 条件(開発用、コンバートタイプ)によるチェックボックスの表示設定
        ''' </summary>
        ''' <remarks>
        ''' ・条件はコマンドラインより取得
        ''' </remarks>
        Private Sub Set_ChkBoxVisible()

            If Not Dev_CVFlg Then

                Select Case CNVNO
                    Case ConvertTypes._汎用

                    Case ConvertTypes._既存ユーザ用

                        '各マスタ情報

                        '業者情報
                        Me.chkGyCyukaiMemo.Visible = False
                        Me.chkGyHokenKoza.Visible = False
                        Me.chkGyHokenMemo.Visible = False
                        Me.chkGySyuzenMemo.Visible = False
                        Me.chkGyYatinhosyoMemo.Visible = False
                        '各基本情報

                        '物件情報
                        Me.chkBkKagi.Visible = False
                        Me.chkBkHendo.Visible = False
                        Me.chkBkKinrincyusyajo.Visible = False
                        Me.chkBkSansyofile.Visible = False

                        '部屋情報
                        Me.chkHyCommonsalespoint.Visible = False
                        Me.chkHyConfirm.Visible = False
                        Me.chkHyKenri.Visible = False
                        Me.chkHySansyofile.Visible = False
                        Me.chkHyGenjotanka.Visible = False

                    Case ConvertTypes._他社システム用

                    Case ConvertTypes._不明

                End Select

            End If

        End Sub

        ''' <summary>
        ''' 関連コントロール内の全チェックボックスON/OFF
        ''' </summary>
        ''' <param name="seltabname">コントロール名</param>
        ''' <param name="setnotflg">True.[setbln]の逆値をセット, False.[setbln]をそのままセット [op]</param>
        ''' <param name="setbln">True or False [op]</param>
        ''' <remarks>
        ''' 例：setnotflg=True, setbln=False の場合、全てTrueでセットする　
        ''' </remarks>
        Public Sub Set_ChkboxOnOff(ByVal seltabname As String, _
                                   Optional ByVal setnotflg As Boolean = False, _
                                   Optional ByVal setbln As Boolean = False)

            Dim chk_checked As Boolean
            Dim chkflg As Integer = 0
            chkflg = IIf(setnotflg, 0, 1)

            Select Case seltabname
                Case DC_SELTAB_K01                       '項目選択1
                    chk_checked = IIf(setnotflg, Not Me.chkKiMstBus.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpKizon1, chk_checked, chkflg)
                Case DC_SELTAB_K02                       '項目選択2
                    Select Case CNVNO
                        Case ConvertTypes._汎用
                            chk_checked = IIf(setnotflg, Not Me.chkKiJisyaBase.Checked, setbln)
                        Case Else
                            chk_checked = IIf(setnotflg, Not Me.chkKiSyskanriBase.Checked, setbln)
                    End Select
                    Call Me.Set_ChkBoxValue(Me.grpKizon5, chk_checked, chkflg)
                    Call Me.Set_ChkBoxValue(Me.grpKizon3, chk_checked, chkflg)
                Case DC_SELTAB_K03                       '項目選択3
                    chk_checked = IIf(setnotflg, Not Me.chkKiBkBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpKizon2, chk_checked, chkflg)

                Case DC_SELTAB_B01                      '各マスタ情報
                    chk_checked = IIf(setnotflg, Not Me.chkMstBus.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpMst, chk_checked, chkflg)
                Case DC_SELTAB_B02                      '業者情報
                    chk_checked = IIf(setnotflg, Not Me.chkGyCyukaiBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpGy, chk_checked, chkflg)
                Case DC_SELTAB_B03                      '各基本情報
                    chk_checked = IIf(setnotflg, Not Me.chkJisyaBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpJisya, chk_checked, chkflg)
                    Call Me.Set_ChkBoxValue(Me.grpOw, chk_checked, chkflg)
                    Call Me.Set_ChkBoxValue(Me.grpKys, chk_checked, chkflg)
                Case DC_SELTAB_B04                      '物件情報
                    chk_checked = IIf(setnotflg, Not Me.chkBkBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpBk, chk_checked, chkflg)
                Case DC_SELTAB_B05                      '部屋情報
                    chk_checked = IIf(setnotflg, Not Me.chkHyBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpHy, chk_checked, chkflg)
                Case DC_SELTAB_B06                      '送金ルール
                    chk_checked = IIf(setnotflg, Not Me.chkSoruleBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpSorule, chk_checked, chkflg)
                Case DC_SELTAB_B07                      '契約情報
                    chk_checked = IIf(setnotflg, Not Me.chkKyBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpKy, chk_checked, chkflg)
                Case DC_SELTAB_B08                      '未収金・過剰金情報
                    chk_checked = IIf(setnotflg, Not Me.chkSqKajyo.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpSq, chk_checked, chkflg)
                Case DC_SELTAB_B09                      'クレーム修繕情報
                    chk_checked = IIf(setnotflg, Not Me.chkClaimBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpClaim, chk_checked, chkflg)
                Case DC_SELTAB_B10                      '初期設定情報
                    chk_checked = IIf(setnotflg, Not Me.chkSyskanriBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpSyskanri, chk_checked, chkflg)
                Case DC_SELTAB_B10                      '物件データ連動情報  '20160720 連動情報構築
                    chk_checked = IIf(setnotflg, Not Me.chkRendoSosinBase.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.grpRendo, chk_checked, chkflg)
                Case DC_SELTAB_H01                      '没選択1(汎用)
                    'chk_checked = IIf(setnotflg, Not Me.chkKiMstBus.Checked, setbln)
                    'Call Me.Set_ChkBoxValue(Me.grpKizon1, chk_checked, chkflg)
                Case DC_SELTAB_H01                      '没選択2(汎用)
                    'chk_checked = IIf(setnotflg, Not Me.chkKiMstBus.Checked, setbln)
                    'Call Me.Set_ChkBoxValue(Me.grpKizon1, chk_checked, chkflg)
                Case DC_SELTAB_H03                      '没選択3(汎用)
                    'chk_checked = IIf(setnotflg, Not Me.chkKiMstBus.Checked, setbln)
                    'Call Me.Set_ChkBoxValue(Me.grpKizon1, chk_checked, chkflg)
                Case M_SYOKITAB_OPS                     'メイン.オプション選択パネル(賃貸革命V7オプション全チェックON)
                    'chk_checked = True
                    chk_checked = IIf(setnotflg, Not Me.chkOptionSelectKy.Checked, setbln)
                    Call Me.Set_ChkBoxValue(Me.pnlOptionSelect, chk_checked, chkflg)
                Case Else

            End Select

        End Sub

        ''' <summary>
        ''' 鍵選択制御(ラジオボタン選択値をチェックボックスへ反映)
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Set_OptSelect()

            optHyKagi.Checked = optKiHyKagi.Checked
            optKyKagi.Checked = optKiKyKagi.Checked
            '20160531 鍵情報移行処理の修正 -del sta
            'chkHyKagi.Enabled = optKiHyKagi.Checked
            'chkKyKagi.Enabled = Not optKiHyKagi.Checked
            '20160531 鍵情報移行処理の修正 -del end

        End Sub

        ''' <summary>
        ''' コンテナへの格納順序指定
        ''' </summary>
        ''' <param name="container"></param>
        ''' <remarks></remarks>
        Private Sub Set_ControlIndex(ByVal container As Control)

            Dim controls As New List(Of Control)

            Select Case container.Name

                '契約者情報より先に自社情報がコンバートされるように順序を指定する
                Case DC_SELTAB_B03
                    '20161026 自社情報と家主情報のコンバート順序入れ替え -chg sta
                    'controls.Add(Me.grpJisya)
                    'controls.Add(Me.grpOw)
                    Select Case CNVNO
                        Case ConvertTypes._汎用
                            controls.Add(Me.grpOw)
                            controls.Add(Me.grpJisya)
                        Case ConvertTypes._既存ユーザ用
                            controls.Add(Me.grpJisya)
                            controls.Add(Me.grpOw)
                    End Select
                    '20161026 自社情報と家主情報のコンバート順序入れ替え -chg end
                    controls.Add(Me.grpKys)
                Case DC_SELTAB_B09
                    controls.Add(Me.grpClaim)
                    controls.Add(Me.grpSzen)

                Case "grpMst"
                    controls.Add(Me.chkMstBus)
                    controls.Add(Me.chkMstBusKotu)
                    controls.Add(Me.chkMstSchool)
                    controls.Add(Me.chkMstArea)
                    controls.Add(Me.chkMstKagititle)
                    controls.Add(Me.chkMstHokenrui)
                    'controls.Add(Me.chkMstKeiyakurui)
                    controls.Add(Me.chkMstTokuyaku)
                    controls.Add(Me.chkMstKasyoClaimrui)
                    controls.Add(Me.chkMstHendo)
                    controls.Add(Me.chkMstHendoitiran)
                    controls.Add(Me.chkMstBikotitle)
                    controls.Add(Me.chkMstBikolst)
                    controls.Add(Me.chkMstGazotitle)                '20160801 タイトルマスタ統合処理 -add
                Case "grpGy"
                    controls.Add(Me.chkGyCyukaiBase)
                    controls.Add(Me.chkGyCyukaiKoza)
                    controls.Add(Me.chkGyCyukaiMemo)
                    controls.Add(Me.chkGyHokenBase)
                    controls.Add(Me.chkGyHokenKoza)
                    controls.Add(Me.chkGyHokenMemo)
                    controls.Add(Me.chkGyYatinhosyoBase)
                    'controls.Add(Me.chkGyYatinhosyoKoza)
                    controls.Add(Me.chkGyYatinhosyoMemo)
                    controls.Add(Me.chkGySyuzenBase)
                    controls.Add(Me.chkGySyuzenKoza)
                    controls.Add(Me.chkGySyuzenMemo)
                    controls.Add(Me.chkGyLifelineBase)
                    controls.Add(Me.chkGySekoBase)
                    controls.Add(Me.chkGySisetuBase)
                Case "grpJisya"
                    controls.Add(Me.chkJisyaBase)
                    controls.Add(Me.chkJisyaKoza)
                    controls.Add(Me.chkJisyaTanto)
                    controls.Add(Me.chkJisyaMemo)
                    controls.Add(Me.chkFBFuriirai)
                    controls.Add(Me.chkFBFuritesuryo)
                    controls.Add(Me.chkFBKozafurikae)
                    controls.Add(Me.chkFBNsSyutoku)
                    controls.Add(Me.chkMstYatinKoza)
                    controls.Add(Me.chkMstANSERArea)
                    controls.Add(Me.chkMstANSERAccpoint)
                    controls.Add(Me.chkFBANSERSetuzoku)
                Case "grpOw"
                    controls.Add(Me.chkOwBase)
                    controls.Add(Me.chkOwKoza)
                    controls.Add(Me.chkOwEvent)
                    controls.Add(Me.chkOwMemo)
                Case "grpKys"
                    controls.Add(Me.chkKysBase)
                    controls.Add(Me.chkKysKoza)
                    controls.Add(Me.chkKysMemo)
                    controls.Add(Me.chkKysSyogoKana)
                    controls.Add(Me.chkKysHosyonin)
                Case "grpBk"
                    controls.Add(Me.chkBkBase)
                    controls.Add(Me.chkBkSyosai)
                    controls.Add(Me.chkBkSyo)
                    controls.Add(Me.chkBkGomi)
                    controls.Add(Me.chkBkKenri)
                    controls.Add(Me.chkBkKotu)
                    controls.Add(Me.chkBkSetudo)
                    controls.Add(Me.chkBkSyuhen)
                    controls.Add(Me.chkBkSzeniji)
                    controls.Add(Me.chkBkMemo)
                    controls.Add(Me.chkBkKagi)                      '汎用分(V7には無い)
                    controls.Add(Me.chkBkHendo)                     '汎用分(V7には無い)
                    controls.Add(Me.chkBkKinrincyusyajo)            '汎用分(V7には無い)
                    controls.Add(Me.chkBkSansyofile)                '汎用分(V7には無い)
                Case "grpHy"
                    controls.Add(Me.chkHyBase)
                    controls.Add(Me.chkHySyosai)
                    controls.Add(Me.chkHySyo)
                    controls.Add(Me.chkHyParking)
                    controls.Add(Me.chkHyTokuyaku)
                    controls.Add(Me.chkHyKagi)
                    controls.Add(Me.chkHyMadoriutiwake)
                    controls.Add(Me.chkHyMenseki)
                    controls.Add(Me.chkHySetubi)
                    controls.Add(Me.chkHyNkinkomk)
                    controls.Add(Me.chkHyHendo)
                    controls.Add(Me.chkHyMemo)
                    controls.Add(Me.chkHyCommonsalespoint)
                    controls.Add(Me.chkHyConfirm)
                    controls.Add(Me.chkHyKenri)
                    controls.Add(Me.chkHySansyofile)
                    controls.Add(Me.chkHyGenjotanka)
                    controls.Add(Me.chkHySzeniji)
                Case "grpSorule"
                    controls.Add(Me.chkSoruleBase)
                    controls.Add(Me.chkSoruleSosaki)
                    controls.Add(Me.chkSoruleNkin)
                    controls.Add(Me.chkSoruleKojo)
                Case "grpKy"
                    controls.Add(Me.chkKyBase)
                    controls.Add(Me.chkKyRireki)
                    controls.Add(Me.chkKyCar)
                    controls.Add(Me.chkKyKys)
                    controls.Add(Me.chkKyHosyonin)
                    controls.Add(Me.chkKyNyukyo)
                    controls.Add(Me.chkKyTokuyaku)
                    controls.Add(Me.chkKyHoken)
                    controls.Add(Me.chkKyMemo)
                    controls.Add(Me.chkKyNkinkomk)
                    controls.Add(Me.chkKyNkinkomkNx)
                    controls.Add(Me.chkKyHendo)
                    controls.Add(Me.chkKyKojoRule)
                    controls.Add(Me.chkKySorule)
                    controls.Add(Me.chkKyKai)
                    controls.Add(Me.chkKySzen)
                    controls.Add(Me.chkKySzenmeisai)
                Case "grpSq"
                    controls.Add(Me.chkSqKajyo)
                    controls.Add(Me.chkSqUnyotaino)
                    controls.Add(Me.chkSqSq)
                    controls.Add(Me.chkSqHendokensin)
                    controls.Add(Me.chkSqKoteiKojo)
                    controls.Add(Me.chkSqSqKojo)
                Case "grpClaim"
                    controls.Add(Me.chkClaimBase)
                    controls.Add(Me.chkClaimTaiorireki)
                    controls.Add(Me.chkClaimRelfile)
                Case "grpSzen"
                    controls.Add(Me.chkSzenBase)
                    controls.Add(Me.chkSzenSzen)
                    controls.Add(Me.chkSzenSzenmeisai)
                    controls.Add(Me.chkSzenClaim)
                    controls.Add(Me.chkSzenRelfile)     			'20160627 修繕関連ファイル移行修正 -add
                    controls.Add(Me.chkSzenMemo)
                Case "grpSyskanri"
                    controls.Add(Me.chkSyskanriBase)
                    controls.Add(Me.chkSyskanriZei)
                    controls.Add(Me.chkSyskanriHenkanmoji)
                    controls.Add(Me.chkSyskanriNkinkomkmerge)
                Case "grpRendo" 									'20160720 連動情報構築 -add
                    controls.Add(Me.chkRendoSosinBase)
                    controls.Add(Me.chkRendoSosinJisyaweb)
                    controls.Add(Me.chkRendoSosinHomes)
                    controls.Add(Me.chkRendoSosinAthome)
                    controls.Add(Me.chkRendoSosinSuumo)
                    controls.Add(Me.chkRendoMapdisp)
                    controls.Add(Me.chkRendoBtoBgroup)
                    controls.Add(Me.chkRendoHysosin)
                    controls.Add(Me.chkRendoHyrui)
                    controls.Add(Me.chkRendoKokokuSuumo)
                    controls.Add(Me.chkRendoKokokuAthome)
                    controls.Add(Me.chkRendoKokokuHomes)
                    controls.Add(Me.chkRendoKokokuJisyaweb)
            End Select

            For i As Integer = 0 To controls.Count - 1
                container.Controls.SetChildIndex(controls(i), i)
            Next

        End Sub

        ''' <summary>
        ''' チェックボックスのON/OFF一括設定(コンテナ毎)
        ''' </summary>
        ''' <param name="container">コンテナ名</param>
        ''' <param name="value">設定値</param>
        ''' <param name="chkflg">強制チェックONOFFフラグ…1.強制, 1以外.通常</param>
        ''' <remarks></remarks>
        Private Sub Set_ChkBoxValue(ByVal container As Control, ByVal value As Boolean, Optional chkflg As Integer = 0)

            Dim list_tab As New List(Of String) From {"grpKizon1", "grpKizon2", "grpKizon3", "grpKizon5"}
           
            If list_tab.Contains(container.Name) Then
                For Each item As Control In container.Controls
                    If item.GetType().Equals(GetType(Panel)) Then
                        Dim tmp_panel As Panel = DirectCast(item, Panel)
                        For Each item_inpanel As Control In tmp_panel.Controls
                            If item_inpanel.GetType().Equals(GetType(CheckBox)) Then
                                Dim tmp_chkbox As CheckBox = DirectCast(item_inpanel, CheckBox)
                                If chkflg = 1 OrElse tmp_chkbox.Visible Then
                                    If tmp_chkbox.Enabled Then
                                        tmp_chkbox.Checked = value
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            Else
                For Each item As Control In container.Controls
                    If item.GetType().Equals(GetType(CheckBox)) Then
                        Dim tmp_chkbox As CheckBox = DirectCast(item, CheckBox)
                        If chkflg = 1 OrElse tmp_chkbox.Visible Then
                            If tmp_chkbox.Enabled Then
                                tmp_chkbox.Checked = value
                            End If
                        End If
                    End If
                Next
            End If
         
        End Sub

        ''' <summary>
        ''' チェックボックスの全チェック(事前作業画面)
        ''' </summary>
        ''' <param name="ctrl">チェック対象コントロール</param>
        ''' <returns>True.全チェック, False.チェックOFFあり</returns>
        ''' <remarks>
        ''' ・"Get_ListCVChkitem"メソッドを参考に作成
        ''' ・Me.tabCtrlJizen.Controls用
        ''' ・仕様上、全タブを表示する
        ''' </remarks>
        Private Function Chk_JizenChkitem(ByVal ctrl As Object) As Boolean

            Dim rtn As Boolean = True
            Dim tmp_grpname As String = ""
            Dim tmp_chkname As String = ""

 
            Select Case CNVNO
                Case ConvertTypes._汎用
                    tabPageManagerJizen.ChangeTabPageVisible(0, False)
                    tabPageManagerJizen.ChangeTabPageVisible(1, False)
                    tabPageManagerJizen.ChangeTabPageVisible(2, False)
                    tabPageManagerJizen.ChangeTabPageVisible(3, False)
                    tabPageManagerJizen.ChangeTabPageVisible(4, False)
                    tabPageManagerJizen.ChangeTabPageVisible(5, True)
                Case Else
                    tabPageManagerJizen.ChangeTabPageVisible(0, True)
                    tabPageManagerJizen.ChangeTabPageVisible(1, True)
                    tabPageManagerJizen.ChangeTabPageVisible(2, True)
                    tabPageManagerJizen.ChangeTabPageVisible(3, True)
                    tabPageManagerJizen.ChangeTabPageVisible(4, True)
                    tabPageManagerJizen.ChangeTabPageVisible(5, False)
            End Select


            'タブコントロール内
            For Each item As Control In ctrl
                If item.GetType().Equals(GetType(TabPage)) Then
                    Dim container_tab As TabPage = DirectCast(item, TabPage)

                    'タブページ内
                    For Each item_sub1 As Control In container_tab.Controls
                        If item_sub1.GetType().Equals(GetType(GroupBox)) Then
                            Dim container_grp As GroupBox = DirectCast(item_sub1, GroupBox)

                            '移行グループ名取得
                            tmp_grpname = container_grp.Text

                            'グループボックス内
                            For Each item_sub2 As Control In container_grp.Controls
                                If item_sub2.GetType().Equals(GetType(Panel)) Then
                                    Dim container_panel As Panel = DirectCast(item_sub2, Panel)
                                    For Each item_sub3 As Control In container_panel.Controls
                                        If item_sub3.GetType().Equals(GetType(CheckBox)) Then
                                            Dim control_chk As CheckBox = DirectCast(item_sub3, CheckBox)
                                            If control_chk.Enabled And control_chk.Checked = False Then
                                                Return False
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 子項目チェックボックスの活性制御 
        ''' </summary>
        ''' <param name="flg"></param>
        ''' <remarks></remarks>
        Private Sub Set_ChkBoxCondition(ByVal flg As Boolean)

            If optCVNew.Checked Then

                '各マスタ情報
                Me.chkMstBusKotu.Enabled = flg
                Me.chkMstHendoitiran.Enabled = flg
                Me.chkMstBikolst.Enabled = flg
                '業者情報
                Me.chkGyCyukaiKoza.Enabled = flg
                Me.chkGyCyukaiMemo.Enabled = flg
                Me.chkGyHokenKoza.Enabled = flg
                Me.chkGyHokenMemo.Enabled = flg
                Me.chkGyYatinhosyoKoza.Enabled = flg
                Me.chkGyYatinhosyoMemo.Enabled = flg
                Me.chkGySyuzenKoza.Enabled = flg
                Me.chkGySyuzenMemo.Enabled = flg
                '自社情報
                Me.chkJisyaKoza.Enabled = flg
                Me.chkJisyaTanto.Enabled = flg
                Me.chkJisyaMemo.Enabled = flg
                Me.chkFBFuriirai.Enabled = flg
                Me.chkFBKozafurikae.Enabled = flg
                '家主情報
                Me.chkOwKoza.Enabled = flg
                Me.chkOwEvent.Enabled = flg
                Me.chkOwMemo.Enabled = flg
                '契約者情報
                Me.chkKysKoza.Enabled = flg
                Me.chkKysMemo.Enabled = flg
                Me.chkKysSyogoKana.Enabled = flg
                Me.chkKysHosyonin.Enabled = flg
                '物件情報
                Me.chkBkSyosai.Enabled = flg
                Me.chkBkSyo.Enabled = flg
                Me.chkBkGomi.Enabled = flg
                Me.chkBkKenri.Enabled = flg
                Me.chkBkKotu.Enabled = flg
                Me.chkBkSetudo.Enabled = flg
                Me.chkBkSyuhen.Enabled = flg
                Me.chkBkSzeniji.Enabled = flg
                Me.chkBkMemo.Enabled = flg
                Me.chkBkKagi.Enabled = flg
                Me.chkBkHendo.Enabled = flg
                Me.chkBkKinrincyusyajo.Enabled = flg
                Me.chkBkSansyofile.Enabled = flg
                '部屋情報
                Me.chkHySyosai.Enabled = flg
                Me.chkHySyo.Enabled = flg
                Me.chkHyParking.Enabled = flg
                Me.chkHyTokuyaku.Enabled = flg
                Me.chkHyKagi.Enabled = flg
                Me.chkHyMadoriutiwake.Enabled = flg
                Me.chkHyMenseki.Enabled = flg
                Me.chkHySetubi.Enabled = flg
                Me.chkHyNkinkomk.Enabled = flg
                Me.chkHyHendo.Enabled = flg
                Me.chkHyMemo.Enabled = flg
                Me.chkHyCommonsalespoint.Enabled = flg
                Me.chkHyConfirm.Enabled = flg
                Me.chkHyKenri.Enabled = flg
                Me.chkHySansyofile.Enabled = flg
                Me.chkHyGenjotanka.Enabled = flg
                Me.chkHySzeniji.Enabled = flg
                '送金ルール情報
                Me.chkSoruleSosaki.Enabled = flg
                Me.chkSoruleNkin.Enabled = flg
                Me.chkSoruleKojo.Enabled = flg
                '契約情報
                Me.chkKyRireki.Enabled = flg
                Me.chkKyCar.Enabled = flg
                Me.chkKyKys.Enabled = flg
                Me.chkKyHosyonin.Enabled = flg
                Me.chkKyNyukyo.Enabled = flg
                Me.chkKyTokuyaku.Enabled = flg
                Me.chkKyHoken.Enabled = flg
                Me.chkKyMemo.Enabled = flg
                Me.chkKyNkinkomk.Enabled = flg
                Me.chkKyNkinkomkNx.Enabled = flg
                Me.chkKyHendo.Enabled = flg
                Me.chkKyKojoRule.Enabled = flg
                Me.chkKySorule.Enabled = flg
                Me.chkKyKai.Enabled = flg
                '20160531 鍵情報移行処理の修正 -del sta
                ''2016.04.06 契約鍵情報の移行処理追加 -add
                'Me.chkKyKagi.Enabled = flg
                '20160531 鍵情報移行処理の修正 -del end

            End If

        End Sub

        ''' <summary>
        ''' 汎用コンバーターのみ表示するチェックボックスのチェックをOFFにする処理
        ''' </summary>
        ''' <remarks>
        ''' ・(汎用)
        ''' </remarks>
        Private Sub Set_BaseCVItemToList()

            '業者情報
            list_baseCVchkitem.Add(Me.chkGyCyukaiMemo)
            list_baseCVchkitem.Add(Me.chkGyHokenKoza)
            list_baseCVchkitem.Add(Me.chkGyHokenMemo)
            list_baseCVchkitem.Add(Me.chkGySyuzenMemo)
            list_baseCVchkitem.Add(Me.chkGyYatinhosyoMemo)

            '物件情報
            list_baseCVchkitem.Add(Me.chkBkHendo)
            list_baseCVchkitem.Add(Me.chkBkKinrincyusyajo)
            list_baseCVchkitem.Add(Me.chkBkSansyofile)

            '部屋情報
            list_baseCVchkitem.Add(Me.chkHyCommonsalespoint)
            list_baseCVchkitem.Add(Me.chkHyConfirm)
            list_baseCVchkitem.Add(Me.chkHyKenri)
            list_baseCVchkitem.Add(Me.chkHySansyofile)
            list_baseCVchkitem.Add(Me.chkHyGenjotanka)

            '契約情報
            '20160531 鍵情報移行処理の修正 -del
            'list_baseCVchkitem.Add(Me.chkKyKagi)

        End Sub

        ''' <summary>
        ''' 既存コンバーターのみ表示するチェックボックスのチェックをOFFにする処理 '20160905 汎用の場合の移行対象外項目(既存のみ移行対象)選定の修正 -add
        ''' </summary>
        ''' <remarks>
        ''' ・(汎用)
        ''' </remarks>
        Private Sub Set_ExistCVItemToList()

            '各基本情報_自社                                        '20160915 汎用移行対象からANSERを除外 -add
            list_existCVchkitem.Add(Me.chkFBNsSyutoku)
            list_existCVchkitem.Add(Me.chkMstANSERAccpoint)
            list_existCVchkitem.Add(Me.chkMstANSERArea)
            list_existCVchkitem.Add(Me.chkFBANSERSetuzoku)
            list_existCVchkitem.Add(Me.chkFBNsSyutoku)          	'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkFBFuritesuryo)        	'20160926 選定した移行項目をプログラムへ反映する修正 -add

            '各基本情報_家主   										'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkOwEvent)

            '各マスタ情報      										'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkMstHendo)
            list_existCVchkitem.Add(Me.chkMstHendoitiran)
            list_existCVchkitem.Add(Me.chkMstKagititle)

            '物件情報
            list_existCVchkitem.Add(Me.chkBkHendo)
            'list_existCVchkitem.Add(Me.chkBkKinrincyusyajo)        '20160926 選定した移行項目をプログラムへ反映する修正 -del
            list_existCVchkitem.Add(Me.chkOpOw)
            list_existCVchkitem.Add(Me.chkOpKys)
            list_existCVchkitem.Add(Me.chkBkGomi)       			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkBkKenri)      			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkBkSetudo)     			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkBkSyuhen)     			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkBkSzeniji)    			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkBkSansyofile) 			'20160926 選定した移行項目をプログラムへ反映する修正 -add

            '部屋情報
            list_existCVchkitem.Add(Me.chkHyMenseki)
            list_existCVchkitem.Add(Me.chkHyKenri)
            list_existCVchkitem.Add(Me.chkHyConfirm)    			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkHySzeniji)    			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkHySansyofile) 			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkHyGenjotanka) 			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkHyHendo)      			'20160926 選定した移行項目をプログラムへ反映する修正 -add

            '契約情報
            list_existCVchkitem.Add(Me.chkKySorule)
            list_existCVchkitem.Add(Me.chkKyKojoRule)
            list_existCVchkitem.Add(Me.chkKyKai)
            list_existCVchkitem.Add(Me.chkKySzen)
            list_existCVchkitem.Add(Me.chkKySzenmeisai)
            list_existCVchkitem.Add(Me.CheckBox14)                  '契約変動費親メーター情報
            list_existCVchkitem.Add(Me.CheckBox24)                  '契約解約確認事項情報
            list_existCVchkitem.Add(Me.CheckBox23)                  '契約同時契約情報
            list_existCVchkitem.Add(Me.CheckBox18)                  '契約原状回復目安単価情報
            list_existCVchkitem.Add(Me.CheckBox16)                  '契約敷金保証金随時処理情報
            list_existCVchkitem.Add(Me.CheckBox15)                  '契約関連ファイル情報
            list_existCVchkitem.Add(Me.CheckBox5)                   '契約空室待ち情報
            list_existCVchkitem.Add(Me.chkKyHendo)					'20160926 選定した移行項目をプログラムへ反映する修正 -add

            '請求情報
            list_existCVchkitem.Add(Me.chkSqSq)
            list_existCVchkitem.Add(Me.chkSqHendokensin)
            list_existCVchkitem.Add(Me.chkSqSqKojo)
            list_existCVchkitem.Add(Me.chkSqKajyo)      			'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(Me.chkSqUnyotaino)  			'20160926 選定した移行項目をプログラムへ反映する修正 -add

        End Sub

#End Region

#Region "紐付設定画面関連処理"

        ''' <summary>
        ''' 紐付項目書込処理
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' ・紐付項目チェックの機能追加時に移行対象項目を設定する
        ''' ・現時点では全項目を移行するようにしておく
        ''' </remarks>
        Private Function Set_RelItem() As Boolean

            Dim rtn As Boolean = True

            '紐付項目取得
            Dim list_relitem As New List(Of String) From {"物件分類マスタ", "部屋分類マスタ", "入金区分マスタ", "契約分類マスタ", "入金項目マスタ", "設備マスタ"}        '20160829 設備の新規挿入処理を追加 "設備マスタ"を追加

            For Each relitem In list_relitem

                '正常終了フラグ
                Dim normalflg As Boolean = True

                '処理Repositoryの生成
                Dim obj_rep As Object = Nothing
                obj_rep = Me.Get_ObjRep_RelItem(relitem)

                If obj_rep IsNot Nothing Then
                    normalflg = obj_rep.Set_RelItem(sqlcnnv10, relitem)
                End If

                If normalflg = False Then
                    rtn = False
                    Return rtn
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 紐付項目書込Repositoryの生成 
        ''' </summary>
        ''' <param name="relitem"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_ObjRep_RelItem(ByVal relitem As String)

            Dim obj_rep As Object = Nothing

            Select Case relitem
                Case "物件分類マスタ" : obj_rep = New Njc.Repository.M_bk_rui_Repository
                Case "部屋分類マスタ" : obj_rep = New Njc.Repository.M_hy_rui_Repository
                Case "入金区分マスタ" : obj_rep = New Njc.Repository.M_nkbn_Repository
                Case "契約分類マスタ" : obj_rep = New Njc.Repository.M_ky_rui_Repository
                Case "入金項目マスタ" : obj_rep = New Njc.Repository.M_nkin_Repository
                Case "設備マスタ" : obj_rep = New Njc.Repository.M_setubi_rel_Repository    '20160829 設備の新規挿入処理を追加 -add

            End Select

            Return obj_rep

        End Function

        ''' <summary>
        ''' 紐付画面呼出処理　'20160707 本体と紐付ツールの中断を同期させる処理の追加 Sub → Function へ変更
        ''' </summary>
        ''' <remarks></remarks>
        Private Function Rel_ExeCall(ByVal cvitemtorel As String) As Boolean

            Dim rowcnt As Integer = 0
            Dim rtn_cancelflg As Boolean = False        									'20160707 本体と紐付ツールの中断を同期させる処理の追加 -add 中断フラグ返却用

            'ログ/状況出力
            Dim tmp_sql_relsta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_RELSTA), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_relsta, rowcnt)
            Call Me.Set_Situation(SITUATION_RELSTA, 0)

            '紐付ツールパス取得(Exe実行用)
            Dim exepath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim exedir As String = IO.Path.GetDirectoryName(exepath)
            Dim tmp_strpath As String = EtcMethod.Set_Path(exedir, DIR_RELEXEDIR_NAME)
            Dim relexepath As String = EtcMethod.Set_Path(tmp_strpath, DIR_RELEXE_NAME)


            '--------------------------------------------------
            ' 紐付設定画面呼出処理
            '--------------------------------------------------
            Dim tmpstre As String = "0" & " " & CNVNO & " "
            
            'V7接続情報
            Dim coninfov7 As String = Me.optV7Authent1.Checked & "," & _
                                      fstmodelv7.ServerName.Trim & "," & _
                                      fstmodelv7.InitialCatalog.Trim & "," & _
                                      fstmodelv7.User.Trim & "," & _
                                      fstmodelv7.Pass.Trim

            '10接続情報
            Dim coninfov10 As String = Me.optV10Authent1.Checked & "," & _
                                       fstmodelv10.ServerName.Trim & "," & _
                                       fstmodelv10.InitialCatalog.Trim & "," & _
                                       fstmodelv10.User.Trim & "," & _
                                       fstmodelv10.Pass.Trim & "," & _
                                       fstmodelv10.TimeOut

            '紐付設定画面PG強制表示位置セット
            Dim tmp_relloc As String = Me.Left.ToString & "," & Me.Top.ToString

            '紐付設定画面へ渡す連結パラメータ (接続情報 + 紐付けファイル格納先 + ログファイル格納先)
            Dim cmd_fk8db_cnv As String = """" & coninfov7 & """" & " " & _
                                          """" & coninfov10 & """" & " " & _
                                          """" & RelationDirPath & """" & " " & _
                                          """" & MiddleDirPath & """" & " " & _
                                          cvitemtorel
            
            cmd_fk8db_cnv = tmpstre & cmd_fk8db_cnv
            cmd_fk8db_cnv = cmd_fk8db_cnv & " " & tmp_relloc                            	'20160829 紐付ツール起動位置修正_本体→紐付 -add
            cmd_fk8db_cnv = cmd_fk8db_cnv & " " & """" & BaseMidDirPath & """"          	'20161004 自社口座の口座種別取得処理の修正 -add

            '◎紐付設定画面起動(連結パラメータ渡し)
            Dim obj_ExeStart As System.Diagnostics.Process = System.Diagnostics.Process.Start(relexepath, cmd_fk8db_cnv)
            obj_ExeStart.WaitForExit()

            '20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg sta
            ''ログ/状況出力
            'Dim tmp_sql_relend As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, LOG_SYORIKOMK_RELEND), False)
            'DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_relend, rowcnt)
            'Call Me.Set_Situation(SITUATION_RELEND, 0)

            '10本体アクティブ化
            Me.Activate()                                                                   '20160829_紐付完了後の本体をアクティブにする修正 -add

            '----- 全体用・個別用プログレスバー更新 -----
            Dim obj_pgb As New ProgressBarManager
            Dim pgbcnt As Integer = 1
            Call obj_pgb.pgbsettingPart(pgbcnt)
            Call obj_com.ProgressOutPut(pgbcnt, pgbcnt)
            Call obj_pgb.pgbsettingTotal(pgbcnt)
            Call obj_com.ProgressOutPut(pgbcnt, pgbcnt, True)
            Me.tabPageJikko.Refresh()

            '紐付設定PGの終了戻り値(紐付設定画面で固定値を設定)を取得 (0.正常終了, 1.中断/キャンセル終了, 99.異常終了)        '20161108 レビュー結果：「紐付設定画面にて固定値設定」のような内容の記載をして下さい 20161108_2 レビュー結果戻り修正 コメントを追記しました。
            Dim exitcode As Integer = obj_ExeStart.ExitCode
            Dim logstr As String = ""
            Select Case exitcode
                Case 0                                                                              '20161108 レビュー結果：エラーの場合もキャンセルフラグを立てている？そうであれば、他のキャンセルフラグの箇所は問題ない？
                    '                                                                                   → 20161108_2 レビュー結果戻り修正
                    '                                                                                        紐付ツール移行の処理をスキップするためにエラーの場合でもキャンセルフラグを立てました。
                    '                                                                                        エラー終了かキャンセル終了かは戻り値「relexeerr」で判別するように修正したので問題ないと思います。
                    '                                                                                        ※エラー終了とキャンセル終了で検証しました。
                    logstr = LOG_SYORIKOMK_RELEND
                    rtn_cancelflg = False
                    '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add
                    relexeerr = exitcode
                Case 1
                    logstr = LOG_SYORIKOMK_RELCANCEL
                    rtn_cancelflg = True
                    '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add
                    relexeerr = exitcode
                    '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add sta
                Case 99
                    logstr = LOG_SYORIKOMK_RELCANCEL
                    rtn_cancelflg = True
                    relexeerr = exitcode
                    '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add end
            End Select

            'ログ/状況出力
            Dim tmp_sql_relend As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, logstr), False)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_relend, rowcnt)
            Call Me.Set_Situation(SITUATION_RELEND, 0)

            '返却
            Return rtn_cancelflg
            '20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg end

        End Function

#End Region

#Region "ログ出力処理"

        ''' <summary>
        ''' CSVファイル出力時にヘッダを挿入する       
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_LogHeader(ByVal filepath As String, ByVal headerstr As String)

            Dim strtype As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")     '文字コード設定
            Dim strread As New StreamReader(filepath, strtype)                                      'ログファイルを開く
            Dim tmp_path As String = Path.GetTempFileName()                                         '仮ファイル作成
            Dim strwrite As New StreamWriter(tmp_path, False, strtype)                              '仮ファイルを開く

            '仮ファイルへヘッダー書込み
            strwrite.WriteLine(headerstr)

            '内容読込
            While strread.Peek() > -1
                Dim line As String = strread.ReadLine()
                strwrite.WriteLine(line)
            End While

            '終了処理
            strread.Close()
            strwrite.Close()

            '仮ファイルと入れ替え
            System.IO.File.Copy(tmp_path, filepath, True)
            System.IO.File.Delete(tmp_path)

        End Sub

        ''' <summary>
        ''' ログ出力初期設定 
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_LogInit(ByVal midchkflg As Boolean)

            Dim rowcnt As Integer = 0
            Dim taisyoname As String = ""

            If midchkflg Then
                taisyoname = LOG_TMP_MIDCHKTABLENAME
            Else
                taisyoname = LOG_TMP_TABLENAME
            End If

            'ログ格納用テーブル初期化(DROP)
            Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(taisyoname, True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_drop, rowcnt)

            'ログ格納用テーブル生成
            Dim tmp_sql_create As String = LogSetting.Get_LogTblCreateQry(midchkflg)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_create, rowcnt)

        End Sub

        ''' <summary>
        ''' ログ出力設定 
        ''' </summary>
        ''' <param name="logdropflg"></param>
        ''' <remarks></remarks>
        Private Sub Set_LogFile(ByVal logdropflg As Boolean)

            Dim rowcnt As Integer = 0

            'ログファイル出力
            If Not EtcMethod.Chk_FileExist(LogFilePath) Then
                'ログファイル(CSV)のデータ部出力
                Dim tmp_sql As String = LogSetting.Get_LogTblSelectQry(False)
                Dim rtn As Boolean = True
                Dim errstr As String = ""
                rtn = FileMethod.TblView_Output_CSV(sqlcnnv10, LogFilePath, "", "", errstr, tmp_sql, True)

                'ヘッダーを加えて加工
                '20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg sta
                'Call Me.Set_LogHeader(LogFilePath, LOG_HEADER_TOTAL)
                Dim tmp_header As String = LOG_HEADER_TOTAL
                If Dev_CVFlg = False Then
                    tmp_header = tmp_header.Replace(",対象TBL名", "")
                End If
                Call Me.Set_LogHeader(LogFilePath, tmp_header)
                '20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg end
            End If

            'ログ格納用テーブル削除(DROP)
            If logdropflg Then
                '中間ファイルチェックログ
                Dim tmp_sql_midchklogdrop As String = DBQuery.Qry_DropInfo(LOG_TMP_MIDCHKTABLENAME, True)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_midchklogdrop, rowcnt)
                'コンバートログ
                Dim tmp_sql_logdrop As String = DBQuery.Qry_DropInfo(LOG_TMP_TABLENAME, True)
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_logdrop, rowcnt)
            End If

        End Sub

#End Region

#Region "開発用処理"

        ''' <summary>
        ''' 開発用タブ表示ON/OFF
        ''' </summary>
        ''' <param name="sender">Object</param>
        ''' <param name="e">Event</param>
        ''' <remarks>
        ''' ・開発用タブ表示時に押下…前回表示タブを表示
        ''' ・上記以外時に押下…開発用タブを表示
        '''   ※UI(タブINDEX)を変更した場合、ここを修正する
        ''' </remarks>
        Private Sub btnDevTabChange_Click(sender As Object, e As EventArgs) Handles btnDevTabChange.Click

            Call Chg_TabSelect()

        End Sub

        ''' <summary>
        ''' 開発用タブ表示制御
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Chg_TabSelect()
            Dim seltab As String = IIf(Me.tabCtrlMain.SelectedTab.Name = "", "", Me.tabCtrlMain.SelectedTab.Name)
            Dim tmptaihitab As Integer
            Dim devtab As Integer = 0               '開発用タブINDEX

            '①現開いているタブINDEX値を退避
            Select Case seltab
                Case "tabPageStart"                 'メイン.開始TAB
                    tmptaihitab = 1
                Case "tabPageSession"               'メイン.DB接続設定TAB
                    tmptaihitab = 2
                Case "tabPageSyoki"                 'メイン.初期設定TAB
                    tmptaihitab = 3
                Case "tabPageMenu"                  'メイン.作業選択TAB
                    tmptaihitab = 4
                Case "tabPageJizen"                 'メイン.事前作業TAB
                    tmptaihitab = 5
                Case "tabPageJigo"                  'メイン.事後作業TAB
                    tmptaihitab = 6
                Case "tabPageHojyo"                 'メイン.補助機能TAB
                    tmptaihitab = 7
                Case "tabPageHajimeni"              'コンバーター.はじめにTAB
                    tmptaihitab = 8
                Case "tabPageSelect"                'コンバーター.項目選択TAB
                    tmptaihitab = 9
                Case "tabPageJikko"                 'コンバーター.コンバート実行TAB
                    tmptaihitab = 10
                Case "tabPageEndOK"                 'コンバーター.終了(正常)TAB
                    tmptaihitab = 11
                Case "tabPageEndError"              'コンバーター.終了(エラー)TAB
                    tmptaihitab = 12
                Case "tabPageEndCancel"             'コンバーター.終了(中断)TAB
                    tmptaihitab = 13
                Case "tabPageIkkatu"                '(仮)コンバーター.一括設定TAB
                    tmptaihitab = 14
                Case "tabPageHanyoJizen"            '(仮)コンバーター.汎用CVK用TAB
                    tmptaihitab = 15
                Case Else
                    tmptaihitab = devtab
            End Select

            '②現開いているタブ毎の制御
            Select Case tmptaihitab
                Case devtab
                    '※開発用タブが開かれている場合は前回のタブを表示
                    Select Case taihitab
                        Case 1
                            tabCtrlMain.SelectedTab = tabPageStart
                        Case 2
                            tabCtrlMain.SelectedTab = tabPageSession
                        Case 3
                            tabCtrlMain.SelectedTab = tabPageSyoki
                        Case 4
                            tabCtrlMain.SelectedTab = tabPageMenu
                        Case 5
                            tabCtrlMain.SelectedTab = tabPageJizen
                        Case 6
                            tabCtrlMain.SelectedTab = tabPageJigo
                        Case 7
                            tabCtrlMain.SelectedTab = tabPageHojyo
                        Case 8
                            tabCtrlMain.SelectedTab = tabPageHajimeni
                        Case 9
                            tabCtrlMain.SelectedTab = tabPageSelect
                        Case 10
                            tabCtrlMain.SelectedTab = tabPageJikko
                        Case 11
                            tabCtrlMain.SelectedTab = tabPageEndOK
                        Case 12
                            tabCtrlMain.SelectedTab = tabPageEndError
                        Case 13
                            tabCtrlMain.SelectedTab = tabPageEndCancel
                        Case 14
                            tabCtrlMain.SelectedTab = tabPageIkkatu
                        Case 15
                            tabCtrlMain.SelectedTab = tabPageHanyoJizen
                        Case Else

                    End Select
                    tabPageManager.ChangeTabPageVisible(taihitab, True)
                    tabPageManager.ChangeTabPageVisible(devtab, False)
                    Call Chg_DevBtn(True)
                Case Else
                    '※現開発用タブ以外が開かれている場合は開発用タブを表示
                    tabCtrlMain.SelectedTab = tabPageDev
                    tabPageManager.ChangeTabPageVisible(devtab, True)
                    tabPageManager.ChangeTabPageVisible(tmptaihitab, False)
                    Call Chg_DevBtn(False)
            End Select
            taihitab = tmptaihitab

        End Sub

        ''' <summary>
        ''' 開発用タブ表示状態のボタン制御(表示)
        ''' </summary>
        ''' <param name="condi">状態…True.表示、False.非表示</param>
        ''' <remarks></remarks>
        Public Sub Chg_DevBtn(ByVal condi As Boolean)

            '「戻る」ボタンへ元の情報格納
            Me.btnBack.Enabled = condi
            '「次へ」ボタンへ元の情報格納
            Me.btnNext.Enabled = condi
            '「中止」ボタンへ元の情報格納
            Me.btnEnd.Enabled = condi

        End Sub

#End Region

#Region "コンバート後データ調整処理"

        ''' <summary>
        ''' データ調整を行うメソッド 20160929 データ調整用メソッドの作成 -add
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub DataCond()

            Dim tmp_cnt As Integer = 0

            '契約者口座情報更新処理
            Dim list_kyscondqry As List(Of String) = DataCondModule.Get_List_Qry_KysKozaCond()
            For Each tmp_sql In list_kyscondqry
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
            Next
            '自社口座情報のゆうちょ→銀行変換処理
            Call DataCondModule.Get_List_Qry_JisyaKozaCond(Me.sqlcnnv10)
            '適用税率更新処理
            Call DataCondModule.Set_ZeiMst(sqlcnnv10)
            'メモタイトル一括更新処理(汎用CVのみ)
            If CNVNO = ConvertTypes._汎用 Then
                Dim list_bikotitlecondqry As List(Of String) = DataCondModule.Get_List_Qry_BikoTitleCond()
                For Each tmp_sql In list_bikotitlecondqry
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
                Next
            End If
            '契約情報のデータ調整
            Call DataCondModule.Set_KydataCond(Me.sqlcnnv10)
            '契約者口座情報のデータ調整
            Call DataCondModule.Set_KysdataKozaCond(Me.sqlcnnv10)
            '当月分の毎月入金項目データ調整
            Call DataCondModule.Set_KydataNkinCond(Me.sqlcnnv10)
            '自社web連動用テーブル・フィールドのデータ調整
            Call DataCondModule.Set_WmpTable(Me.sqlcnnv10)
        End Sub

#End Region

#Region "事前/事後/補助作業関連共通処理 (リスト出力など)"

        ''' <summary>
        ''' 事前作業処理 '20160707 事前作業リスト出力処理の追加 メソッド名 Set_Jizen → Set_List_Output
        ''' </summary>
        ''' <param name="listkomk"></param>
        ''' <remarks></remarks>
        Private Sub Set_List_Output(ByVal listkomk As String)

            Dim tmp_komk As String = ""         'リストの項目名                              
            Dim tmp_sql As String = ""          '抽出クエリ格納用
            Dim tmp_headergrp As String = ""    '抽出クエリ作成時にヘッダーも取得しておく
            Dim tmp_filename As String = ""     '出力ファイル名
            Dim obj_cnn As System.Data.SqlClient.SqlConnection = sqlcnnv7                      '20160711 未使用データ削除_家主 未使用データ抽出に伴い10DBからも抽出するため追加 -add
            Dim postfilename As String = "_" & Replace((Replace(Replace((String.Format(Now)), ":", ""), "/", "")), " ", "") & ".csv"

            '事前作業リスト格納先フォルダ有無確認
            Dim tmp_jizenlistdir As String = Me.txtJizenListPath.Text
            If EtcMethod.Chk_DirExist(tmp_jizenlistdir) = False Then
                MsgResult = MessageBox.Show(MSG_ERR_DIR_LIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            Dim listdirpath As String = tmp_jizenlistdir

            '押下されたボタン名から各情報を設定する
            Select Case listkomk

                '--------------------------------------------------
                ' 事前作業リスト
                '--------------------------------------------------
                Case "btnListJizenAzu"                                                          '20160707 事前作業リスト出力エラー修正 -chg：btnSupAzuList → btnListJizenAzu
                    tmp_komk = Me.chkJizenAzu.Text.Replace("処理", "")
                    tmp_sql = Me.Get_UseQry_AzuList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenKai"                                                          '20160707 事前作業リスト出力処理の追加_解約未処理 -add
                    tmp_komk = Me.chkJizenKai.Text.Replace("データ処理", "データ")
                    tmp_sql = Me.Get_UseQry_KaiList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenMinus"                                                        '20160707 事前作業リスト出力処理の追加_マイナス請求 -add
                    tmp_komk = Me.chkJizenMinus.Text.Replace("調整", "")
                    tmp_sql = Me.Get_UseQry_MinusList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenBkhourei"                                                     '20160707 事前作業リスト出力処理の追加_物件法令 -add
                    tmp_komk = Me.chkJizenBkhourei.Text.Replace("調整", "")
                    tmp_sql = Me.Get_UseQry_BkhoureiList(tmp_headergrp, 3)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenHasseiOw"                                                     '20160707 事前作業リスト出力処理の追加_家主控除発生月 -add
                    tmp_komk = Me.chkJizenHasseiOw.Text.Replace("調整", "")
                    tmp_sql = Me.Get_UseQry_HasseiOwList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenHasseiHen"                                                    '20160707 事前作業リスト出力処理の追加_変動費発生月 -add
                    tmp_komk = Me.chkJizenHasseiHen.Text.Replace("調整", "")
                    tmp_sql = Me.Get_UseQry_HasseiHenList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenYanuso"                                                       '20160707 事前作業リスト出力処理の追加_送金先未設定 -add
                    tmp_komk = Me.chkJizenYanuso.Text.Replace("データの確認・調整", "")
                    tmp_sql = Me.Get_UseQry_YanusoList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenSzenKyshutan"                                                 '20160707 事前作業リスト出力処理の追加_修繕負担者 -add
                    tmp_komk = Me.chkJizenSzenKyshutan.Text.Replace("データの確認", "")
                    tmp_sql = Me.Get_UseQry_SzenKyshutanList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenHyKagi"                                                       '20160725 事前作業へ鍵情報の追加 -add
                    tmp_komk = "部屋鍵情報"
                    tmp_sql = Me.Get_UseQry_HyKagiList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenKyKagi"                                                       '20160725 事前作業へ鍵情報の追加 -add
                    tmp_komk = "契約鍵情報"
                    tmp_sql = Me.Get_UseQry_KyKagiList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenNonJisyaKoza"                                                 '20160805 改善対応 -add
                    tmp_komk = "口座実情報未登録"
                    tmp_sql = Me.Get_UseQry_NonJisyaKozaList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenSimeSokin"                                                    '20160829 事前作業_送金予定日リスト出力機能を追加 -add
                    tmp_komk = "締日、送金日の間隔が1ヶ月以上"
                    tmp_sql = Me.Get_UseQry_SimeSokinList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListJizenKozameigikana"                                                '20160829 口座名義カナチェック機能の追加 -add
                    Call Me.Get_UseQry_Kozameigikana_DropStPr()                                 '抽出用のストアドを削除
                    Call Me.Get_UseQry_Kozameigikana_CreateStPr()                               '抽出用のストアドを事前作成
                    tmp_komk = "使用不可文字を含んだ口座名義カナ"
                    tmp_sql = Me.Get_UseQry_KozameigikanaList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename

                    '--------------------------------------------------
                    '事後作業リスト
                    '--------------------------------------------------
                    '※構築中

                    '--------------------------------------------------
                    '検証作業リスト
                    '--------------------------------------------------
                Case "btnListKiOpOw"                                                            '20160711 未使用データ削除_家主 -add
                    tmp_komk = "未使用家主データ"
                    tmp_sql = Me.Set_NotUseOwdata(1, tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                    obj_cnn = sqlcnnv10
                Case "btnListKiOpKys"                                                           '20160711 未使用データ削除_契約者 -add
                    tmp_komk = "未使用契約者データ"
                    tmp_sql = Me.Set_NotUseKysdata(1, tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                    obj_cnn = sqlcnnv10
                Case "btnListBunkatumisyu"                                                      '20160707 検証作業リスト出力処理の追加_分割入金未収分 -add
                    tmp_komk = Me.chkBunkatumisyu.Text.Replace("データの確認", "")
                    tmp_sql = Me.Get_UseQry_BunkatumisyuList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListKojyosh"                                                           '20160707 検証作業リスト出力処理の追加_控除支払 -add
                    tmp_komk = Me.chkKojyosh.Text.Replace("データの確認", "")
                    tmp_sql = Me.Get_UseQry_KojyoshList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
                Case "btnListSzenKysSorit"                                                      '20160707 検証作業リスト出力処理の追加_修繕項目毎契約者送金率 -add
                    tmp_komk = Me.chkSzenKysSorit.Text.Replace("データの確認", "")
                    tmp_sql = Me.Get_UseQry_SzenKysSoritList(tmp_headergrp)
                    tmp_filename = tmp_komk & postfilename
            End Select

            '出力確認メッセージ
            Dim tmp_msg As String = tmp_komk & MSG_LIST_OUTSTA
            MsgResult = MessageBox.Show(tmp_msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                Exit Sub
            End If

            '出力処理
            Dim normalflg As Boolean = True
            Dim errstr As String = ""
            Dim csvlistfilepath As String = EtcMethod.Set_Path(listdirpath, tmp_filename)
            normalflg = FileMethod.TblView_Output_CSV(obj_cnn, csvlistfilepath, "", "", errstr, tmp_sql, True)

            '終了処理
            If normalflg Then
                'ヘッダーを加えて加工
                Call Me.Set_LogHeader(csvlistfilepath, tmp_headergrp)
                MsgResult = MessageBox.Show(MSG_LIST_OUTEND, "成功", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                System.Diagnostics.Process.Start(csvlistfilepath)
            Else
                MsgResult = MessageBox.Show(MSG_LIST_OUTEND_ERR, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            '20160829 口座名義カナチェック機能の追加 -add sta
            If listkomk = "btnListJizenKozameigikana" Then
                Call Me.Get_UseQry_Kozameigikana_DropStPr()                                     '抽出用のストアドを削除
            End If
            '20160829 口座名義カナチェック機能の追加 -add end
        End Sub

#End Region

#Region "事前作業関連処理"

        ''' <summary>
        ''' 事前作業処理_件数表示 '20160707 事前作業リストの件数表示処理の追加 -add
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_List_Cnt_Jizen()

            Dim tmp_header As String = ""                                                       '使用しないが引数として必要であるため仮で用意する
            Dim cnt_outputstr As String = ""
            Dim tmp_sql As String = ""

            '解約未処理
            tmp_sql = Me.Get_UseQry_KaiList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenKai.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '預り金
            tmp_sql = Me.Get_UseQry_AzuList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenAzu.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            'マイナス金額
            tmp_sql = Me.Get_UseQry_MinusList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenMinus.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '物件法令1_標準以外
            tmp_sql = Me.Get_UseQry_BkhoureiList(tmp_header, 1, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenBkhoureiNotDef.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '物件法令2_標準で移行先が存在しない
            tmp_sql = Me.Get_UseQry_BkhoureiList(tmp_header, 2, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenBkhoureiNotCV.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '家主固定控除発生月
            tmp_sql = Me.Get_UseQry_HasseiOwList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenHasseiOw.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '随時変動費発生月
            tmp_sql = Me.Get_UseQry_HasseiHenList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenHasseiHen.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '送金先未設定
            tmp_sql = Me.Get_UseQry_YanusoList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenYanuso.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '修繕対象負担者設定契約者
            tmp_sql = Me.Get_UseQry_SzenKyshutanList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenSzenKyshutan.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '部屋鍵情報                                                                          '20160725 事前作業へ鍵情報の追加 -add
            tmp_sql = Me.Get_UseQry_HyKagiList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenHyKagi.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '契約鍵情報                                                                          '20160725 事前作業へ鍵情報の追加 -add
            tmp_sql = Me.Get_UseQry_KyKagiList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenKyKagi.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '口座実情報未登録
            tmp_sql = Me.Get_UseQry_NonJisyaKozaList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenNonJisyaKoza.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '締、送金日の間隔が1ヶ月以上                                                         '20160829 事前作業_送金予定日リスト出力機能を追加 -add
            tmp_sql = Me.Get_UseQry_SimeSokinList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenSimeSokin.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '使用不可文字が含まれている口座名義カナ                                            　'20160829 口座名義カナチェック機能の追加 -add
            Call Me.Get_UseQry_Kozameigikana_DropStPr()                                          '抽出用のストアドを削除
            Call Me.Get_UseQry_Kozameigikana_CreateStPr()                                        '抽出用のストアドを事前作成
            tmp_sql = Me.Get_UseQry_KozameigikanaList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntJizenKozameigikana.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"
            Call Me.Get_UseQry_Kozameigikana_DropStPr()                                          '抽出用のストアドを削除

        End Sub


        ''' <summary>
        ''' 預り金抽出クエリ
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_AzuList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,物件名,部屋番号,契約No,該当年月,契約者No,契約者名"

            Dim tmp_sql As String = ""
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[部屋番号],[契約No],[該当年月],[契約者No] ")

            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 SQ.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,BK.bk_name AS [物件名] "
            tmp_sql = tmp_sql & " 		,SQ.hy_no AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,SQ.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,SQ.gt_ym AS [該当年月] "
            tmp_sql = tmp_sql & " 		,SQ.kys_no AS [契約者No] "
            tmp_sql = tmp_sql & " 		,KYS.kys_name AS [契約者名] "
            tmp_sql = tmp_sql & " 	FROM sq_meisai SQ "
            tmp_sql = tmp_sql & " 	LEFT JOIN bk_mst BK ON BK.bk_no=SQ.bk_no "
            tmp_sql = tmp_sql & " 	LEFT JOIN kys_mst KYS ON KYS.kys_no=SQ.kys_no "
            tmp_sql = tmp_sql & " 	WHERE SQ.nkin_no = 6020 "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 解約未処理抽出クエリ '20160707 事前作業リスト出力処理の追加_解約未処理
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_KaiList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,部屋No,契約No,更新No,契約開始日,契約終了日,更新有無"

            Dim tmp_sql As String = ""
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY 物件No,部屋No,契約No,更新No ")

            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	/*20160725 解約未処理情報件数表示時の抽出クエリ修正 chg sta*/ "
            tmp_sql = tmp_sql & " 	/* "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 KK.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,KK.hy_no AS [部屋No] "
            tmp_sql = tmp_sql & " 		,KK.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,KK.ko_no AS [更新No] "
            tmp_sql = tmp_sql & " 		,KK.ky_start_ymd AS [契約開始日] "
            tmp_sql = tmp_sql & " 		,KK.ky_end_ymd AS [契約終了日] "
            tmp_sql = tmp_sql & " 		,CASE nx_kosin_kbn "
            tmp_sql = tmp_sql & " 			/*1:自動更新 2:更新有り 3:更新無し(不可)*/ "
            tmp_sql = tmp_sql & " 			WHEN 1 THEN '自動更新' "
            tmp_sql = tmp_sql & " 			WHEN 3 THEN '更新しない' "
            tmp_sql = tmp_sql & " 			ELSE '' "
            tmp_sql = tmp_sql & " 		 END						AS '更新有無' "
            tmp_sql = tmp_sql & " 	FROM ky_kosinkai KK "
            tmp_sql = tmp_sql & " 	WHERE "
            tmp_sql = tmp_sql & " 		KK.ko_no<>999						/*解約データではない(解約未処理の可能性ありを含む)*/ "
            tmp_sql = tmp_sql & " 		AND KK.ky_end_ymd<'2016/01/01'		/*契約終了データ(運用開始日が契約終了日を超えている)*/ "
            tmp_sql = tmp_sql & " 		AND NOT KK.ky_end_ymd IS NULL		/*仮契約ではない*/ "
            tmp_sql = tmp_sql & " 		AND NOT EXISTS "
            tmp_sql = tmp_sql & " 			(					/*請求データが作成されていない(契約が終了している)*/ "
            tmp_sql = tmp_sql & " 				SELECT * FROM sq_meisai SM  "
            tmp_sql = tmp_sql & " 				WHERE  "
            tmp_sql = tmp_sql & " 					( "
            tmp_sql = tmp_sql & " 						ko_no=999																				/*解約データ*/ "
            tmp_sql = tmp_sql & " 						OR (CONVERT(VARCHAR,SM.gt_ym,112)/100)<=(CONVERT(VARCHAR,KK.ky_end_ymd,112)/100)		/*契約終了日以前の該当日付*/ "
            tmp_sql = tmp_sql & " 					) "
            tmp_sql = tmp_sql & " 				AND SM.bk_no=KK.bk_no AND SM.hy_no=KK.hy_no AND SM.ky_no=KK.ky_no AND SM.ko_no=KK.ko_no			/*紐付け設定*/ "
            tmp_sql = tmp_sql & " 			) "
            tmp_sql = tmp_sql & " 	*/ "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 NOTKAIYAKU.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,NOTKAIYAKU.hy_no AS [部屋No] "
            tmp_sql = tmp_sql & " 		,NOTKAIYAKU.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,NOTKAIYAKU.ko_no AS [更新No] "
            tmp_sql = tmp_sql & " 		,NOTKAIYAKU.ky_start_ymd AS [契約開始日] "
            tmp_sql = tmp_sql & " 		,NOTKAIYAKU.ky_end_ymd AS [契約終了日] "
            tmp_sql = tmp_sql & " 		,CASE nx_kosin_kbn "
            tmp_sql = tmp_sql & " 			/*1:自動更新 2:更新有り 3:更新無し(不可)*/ "
            tmp_sql = tmp_sql & " 			WHEN 1 THEN '自動更新' "
            tmp_sql = tmp_sql & " 			WHEN 3 THEN '更新しない' "
            tmp_sql = tmp_sql & " 			ELSE '' "
            tmp_sql = tmp_sql & " 		 END AS [更新有無] "
            tmp_sql = tmp_sql & " 	FROM "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		/*解約処理されていないデータを抽出 sta*/ "
            tmp_sql = tmp_sql & " 		SELECT * FROM ky_kosinkai AS KYK "
            tmp_sql = tmp_sql & " 		WHERE EXISTS "
            tmp_sql = tmp_sql & " 		( "
            tmp_sql = tmp_sql & " 			/*解約未処理データのキーを抽出 sta*/ "
            tmp_sql = tmp_sql & " 			SELECT * FROM "
            tmp_sql = tmp_sql & " 			( "
            tmp_sql = tmp_sql & " 				SELECT * FROM "
            tmp_sql = tmp_sql & " 				( "
            tmp_sql = tmp_sql & " 					SELECT "
            tmp_sql = tmp_sql & " 						 bk_no "
            tmp_sql = tmp_sql & " 						,hy_no "
            tmp_sql = tmp_sql & " 						,ky_no "
            tmp_sql = tmp_sql & " 						,MAX(ko_no) AS ko_no "
            tmp_sql = tmp_sql & " 					FROM ky_kosinkai "
            tmp_sql = tmp_sql & " 					GROUP BY bk_no,hy_no,ky_no "
            tmp_sql = tmp_sql & " 					) AS VW "
            tmp_sql = tmp_sql & " 				WHERE ko_no <> 999 "
            tmp_sql = tmp_sql & " 			) AS VW "
            tmp_sql = tmp_sql & " 			/*解約未処理データのキーを抽出 end*/ "
            tmp_sql = tmp_sql & " 			WHERE KYK.bk_no = VW.bk_no "
            tmp_sql = tmp_sql & " 			AND   KYK.hy_no = VW.hy_no "
            tmp_sql = tmp_sql & " 			AND   KYK.ky_no = VW.ky_no "
            tmp_sql = tmp_sql & " 			AND   KYK.ko_no = VW.ko_no "
            tmp_sql = tmp_sql & " 		) "
            tmp_sql = tmp_sql & " 		/*解約処理されていないデータを抽出 end*/ "
            tmp_sql = tmp_sql & " 	) AS NOTKAIYAKU "
            tmp_sql = tmp_sql & " 	WHERE NOTKAIYAKU.ky_end_ymd IS NOT NULL						/*仮契約ではない*/ "
            tmp_sql = tmp_sql & " 	/*20160905 運用開始年月置換文字列修正 chg sta*/ "
            tmp_sql = tmp_sql & " 	/*AND   NOTKAIYAKU.ky_end_ymd < '20160701'					/*契約終了データ(運用開始日が契約終了日を超えている)運用開始年月置換用文字列*/*/ "
            tmp_sql = tmp_sql & " 	AND   NOTKAIYAKU.ky_end_ymd < '運用開始年月置換用文字列'	/*契約終了データ(運用開始日が契約終了日を超えている)運用開始年月置換用文字列*/ "
            tmp_sql = tmp_sql & " 	/*20160905 運用開始年月置換文字列修正 chg end*/ "
            tmp_sql = tmp_sql & " 	AND "
            tmp_sql = tmp_sql & " 		( "
            tmp_sql = tmp_sql & " 			SELECT COUNT(*) FROM sq_meisai AS SQ "
            tmp_sql = tmp_sql & " 			WHERE SQ.bk_no = NOTKAIYAKU.bk_no "
            tmp_sql = tmp_sql & " 			AND   SQ.hy_no = NOTKAIYAKU.hy_no "
            tmp_sql = tmp_sql & " 			AND   SQ.ky_no = NOTKAIYAKU.ky_no "
            tmp_sql = tmp_sql & " 			AND   SQ.ko_no = NOTKAIYAKU.ko_no "
            tmp_sql = tmp_sql & " 			AND   NOTKAIYAKU.ky_end_ymd < SQ.gt_ym "
            tmp_sql = tmp_sql & " 		) = 0													/*契約終了日以降の請求データが存在しない*/ "
            tmp_sql = tmp_sql & " 	/*20160725 解約未処理情報件数表示時の抽出クエリ修正 chg end*/ "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr
            '20160725 解約未処理情報件数表示時の抽出クエリ修正 -add sta
            Dim tmp_unyoymd As String = Me.txtUnyoYYYYMM.Text & "/01"
            tmp_sql = tmp_sql.Replace("運用開始年月置換用文字列", tmp_unyoymd)
            '20160725 解約未処理情報件数表示時の抽出クエリ修正 -add end
            Return tmp_sql

        End Function

        ''' <summary>
        ''' マイナス請求抽出クエリ '20160707 事前作業リスト出力処理の追加_マイナス請求
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_MinusList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,物件名,部屋番号,契約No,更新No,入金項目コード,入金区分,入金項目名,請求額,送金額"

            Dim tmp_sql As String = ""
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY 物件No,部屋番号,契約No,更新No,入金区分,入金項目コード ")

            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 BK.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,BK.bk_name AS [物件名] "
            tmp_sql = tmp_sql & " 		,KS.hy_no AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,KS.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,KS.ko_no AS [更新No] "
            tmp_sql = tmp_sql & " 		,KS.nkin_no AS [入金項目コード] "
            tmp_sql = tmp_sql & " 		/*,KS.nkin_kbn AS [入金区分]*/ "
            tmp_sql = tmp_sql & " 		,CASE KS.nkin_kbn "
            tmp_sql = tmp_sql & " 			WHEN 1 THEN '1.入金項目-1.毎月' "
            tmp_sql = tmp_sql & " 			WHEN 2 THEN '1.入金項目-2.契約時' "
            tmp_sql = tmp_sql & " 			WHEN 3 THEN '1.入金項目-3.更新時' "
            tmp_sql = tmp_sql & " 			WHEN 4 THEN '1.入金項目-4.解約時' "
            tmp_sql = tmp_sql & " 			WHEN 5 THEN '1.入金項目-5.随時・変動' "
            tmp_sql = tmp_sql & " 			WHEN 6 THEN '1.入金項目-6.その他請求' "
            tmp_sql = tmp_sql & " 			WHEN 7 THEN '1.入金項目-7.控除' "
            tmp_sql = tmp_sql & " 			WHEN 8 THEN '1.入金項目-8.精算・修繕' "
            tmp_sql = tmp_sql & " 			ELSE '' "
            tmp_sql = tmp_sql & " 		END AS [入金区分] "
            tmp_sql = tmp_sql & " 		,MN.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 		,KS.sq_gak AS [請求額] "
            tmp_sql = tmp_sql & " 		,KS.so_gak AS [送金額] "
            tmp_sql = tmp_sql & " 		/*,**/ "
            tmp_sql = tmp_sql & " 	FROM ky_sqdata KS "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin MN ON MN.nkin_no=KS.nkin_no "
            tmp_sql = tmp_sql & " 	LEFT JOIN bk_mst BK ON BK.bk_no=KS.bk_no "
            tmp_sql = tmp_sql & " 	WHERE  "
            tmp_sql = tmp_sql & " 		(sq_gak<0 OR so_gak<0) "
            tmp_sql = tmp_sql & " 		AND NOT KS.nkin_no IN ('4010','4020')	/*4010.敷金戻し, 4020.保証金戻し*/ "
            tmp_sql = tmp_sql & " 		AND NOT KS.nkin_no='4090'				/*4090.敷金差額*/ "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	/*契約請求データ(次回分用)*/ "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 BK.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,BK.bk_name AS [物件名] "
            tmp_sql = tmp_sql & " 		,KS.hy_no AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,KS.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,KS.ko_no AS [更新No] "
            tmp_sql = tmp_sql & " 		,KS.nkin_no AS [入金項目コード] "
            tmp_sql = tmp_sql & " 		/*,KS.nkin_kbn AS [入金区分]*/ "
            tmp_sql = tmp_sql & " 		,CASE KS.nkin_kbn "
            tmp_sql = tmp_sql & " 			WHEN 1 THEN '2.次回更新-1.毎月' "
            tmp_sql = tmp_sql & " 			WHEN 2 THEN '2.次回更新-2.契約時' "
            tmp_sql = tmp_sql & " 			WHEN 3 THEN '2.次回更新-3.更新時' "
            tmp_sql = tmp_sql & " 			WHEN 4 THEN '2.次回更新-4.解約時' "
            tmp_sql = tmp_sql & " 			WHEN 5 THEN '2.次回更新-5.随時・変動' "
            tmp_sql = tmp_sql & " 			WHEN 6 THEN '2.次回更新-6.その他請求' "
            tmp_sql = tmp_sql & " 			WHEN 7 THEN '2.次回更新-7.控除' "
            tmp_sql = tmp_sql & " 			WHEN 8 THEN '2.次回更新-8.精算・修繕' "
            tmp_sql = tmp_sql & " 			ELSE '' "
            tmp_sql = tmp_sql & " 		END AS [入金区分] "
            tmp_sql = tmp_sql & " 		,MN.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 		,KS.sq_gak AS [請求額] "
            tmp_sql = tmp_sql & " 		,KS.so_gak AS [送金額] "
            tmp_sql = tmp_sql & " 		/*,**/ "
            tmp_sql = tmp_sql & " 	FROM ky_sqdata_nx KS "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin MN ON MN.nkin_no=KS.nkin_no "
            tmp_sql = tmp_sql & " 	LEFT JOIN bk_mst BK ON BK.bk_no=KS.bk_no "
            tmp_sql = tmp_sql & " 	WHERE  "
            tmp_sql = tmp_sql & " 		(sq_gak<0 OR so_gak<0) "
            tmp_sql = tmp_sql & " 		AND NOT KS.nkin_no IN ('4010','4020')	/*4010.敷金戻し, 4020.保証金戻し*/ "
            tmp_sql = tmp_sql & " 		AND NOT KS.nkin_no='4090'				/*4090.敷金差額*/ "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 物件法令抽出クエリ '20160707 事前作業リスト出力処理の追加_物件法令
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cnttype"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_BkhoureiList(ByRef headergrp As String, Optional ByVal cnttype As Integer = 0, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,行No,タイトル名,内容,判定"

            Dim tmp_sql As String = ""
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[行No] ")
            Dim wherestr As String = ""

            Select Case cnttype
                Case 0
                    wherestr = ""
                Case 1
                    wherestr = " WHERE [判定区分] IN (1,4) "
                Case 2
                    wherestr = " WHERE [判定区分] IN (3) "
                Case 3
                    wherestr = " WHERE [判定区分] IN (1,3,4) "
            End Select

            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 [物件No] "
            tmp_sql = tmp_sql & " 		,[行No] "
            tmp_sql = tmp_sql & " 		,[タイトル名] "
            tmp_sql = tmp_sql & " 		,[内容] "
            tmp_sql = tmp_sql & " 		,[判定結果] "
            tmp_sql = tmp_sql & " 	FROM "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		SELECT "
            tmp_sql = tmp_sql & " 			 TOTAL.* "
            tmp_sql = tmp_sql & " 			,CASE "
            tmp_sql = tmp_sql & " 				WHEN [判定結果] = 'タイトルが標準以外の値になっているため移行できません' THEN 1 "
            tmp_sql = tmp_sql & " 				WHEN [判定結果] = '物件権利情報への移行時に判別して移行します' THEN 2 "
            tmp_sql = tmp_sql & " 				WHEN [判定結果] = '革命10に移行先が存在しません' THEN 3 "
            tmp_sql = tmp_sql & " 				WHEN [判定結果] = '内容が標準以外の値になっているため移行できません' THEN 4 "
            tmp_sql = tmp_sql & " 				WHEN [判定結果] = 'コンバート時に内容が標準と同形式か判別して移行します' THEN 2 "
            tmp_sql = tmp_sql & " 				ELSE 0 "
            tmp_sql = tmp_sql & " 			 END AS [判定区分] "
            tmp_sql = tmp_sql & " 		FROM "
            tmp_sql = tmp_sql & " 		( "
            tmp_sql = tmp_sql & " 			/*タイトルが編集されているデータ sta*/ "
            tmp_sql = tmp_sql & " 			SELECT * FROM "
            tmp_sql = tmp_sql & " 			( "
            tmp_sql = tmp_sql & " 				SELECT  "
            tmp_sql = tmp_sql & " 					 bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 					,BKJ.jyuyo_no AS [行No] "
            tmp_sql = tmp_sql & " 					,jyuyo_name AS [タイトル名] "
            tmp_sql = tmp_sql & " 					,jyuyo_lstname AS [内容] "
            tmp_sql = tmp_sql & " 					,'タイトルが標準以外の値になっているため移行できません' AS [判定結果] "
            tmp_sql = tmp_sql & " 				FROM bk_jyuyo AS BKJ "
            tmp_sql = tmp_sql & " 				LEFT JOIN "
            tmp_sql = tmp_sql & " 					( "
            tmp_sql = tmp_sql & " 						SELECT * FROM "
            tmp_sql = tmp_sql & " 						( "
            tmp_sql = tmp_sql & " 							SELECT * FROM m_jyuyo "
            tmp_sql = tmp_sql & " 							WHERE jyuyo_kbn = 1 "
            tmp_sql = tmp_sql & " 						) AS VW "
            tmp_sql = tmp_sql & " 					) AS JYUYOTITLE "
            tmp_sql = tmp_sql & " 				ON BKJ.jyuyo_no = JYUYOTITLE.jyuyo_no "
            tmp_sql = tmp_sql & " 				WHERE jyuyo_name NOT IN "
            tmp_sql = tmp_sql & " 									( "
            tmp_sql = tmp_sql & " 										/*標準だが移行先が存在しない sta*/ "
            tmp_sql = tmp_sql & " 										 '所有権以外の権利の有無' "
            tmp_sql = tmp_sql & " 										,'所有権以外の権利の種類' "
            tmp_sql = tmp_sql & " 										,'建物の登録名義人調査日' "
            tmp_sql = tmp_sql & " 										,'建物の登録名義人と貸主が異なる理由' "
            tmp_sql = tmp_sql & " 										,'未完成物件' "
            tmp_sql = tmp_sql & " 										,'損害賠償額の予定又は違約金に関する定め' "
            tmp_sql = tmp_sql & " 										,'保全措置を行う機関' "
            tmp_sql = tmp_sql & " 										,'備考（契約の種類・期間・更新等）' "
            tmp_sql = tmp_sql & " 										,'公正証書にする場合の費用負担' "
            tmp_sql = tmp_sql & " 										,'備考(定期借家契約の場合)' "
            tmp_sql = tmp_sql & " 										,'敷金等の清算に関する事項' "
            tmp_sql = tmp_sql & " 										,'ﾏﾝｼｮﾝの管理の適正化の推進に関する法律' "
            tmp_sql = tmp_sql & " 										/*標準だが移行先が存在しない end*/ "
            tmp_sql = tmp_sql & " 										/*標準で移行先が存在する sta*/ "
            tmp_sql = tmp_sql & " 										,'所有権にかかる権利の有無' "
            tmp_sql = tmp_sql & " 										,'所有権にかかる権利の種類' "
            tmp_sql = tmp_sql & " 										,'法令に基づく制限の概要（法令名）' "
            tmp_sql = tmp_sql & " 										,'法令に基づく制限の概要' "
            tmp_sql = tmp_sql & " 										,'敷地利用関係の種類(敷地が借地の場合)' "
            tmp_sql = tmp_sql & " 										,'契約期間(敷地が賃借の場合）' "
            tmp_sql = tmp_sql & " 										,'備考(敷地が借地の場合)' "
            tmp_sql = tmp_sql & " 										,'宅地造成等規正法' "
            tmp_sql = tmp_sql & " 										,'土砂災害防止対策推進法' "
            tmp_sql = tmp_sql & " 										,'津波防災地域づくりに関する法律' "
            tmp_sql = tmp_sql & " 										/*標準で移行先が存在する end*/ "
            tmp_sql = tmp_sql & " 									) "
            tmp_sql = tmp_sql & " 			) AS NOTBASE "
            tmp_sql = tmp_sql & " 			/*タイトルが編集されているデータ end*/ "
            tmp_sql = tmp_sql & " 			UNION "
            tmp_sql = tmp_sql & " 			/*標準タイトルだが移行先が存在しないデータ sta*/ "
            tmp_sql = tmp_sql & " 			SELECT * FROM "
            tmp_sql = tmp_sql & " 			( "
            tmp_sql = tmp_sql & " 				SELECT  "
            tmp_sql = tmp_sql & " 					 bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 					,BKJ.jyuyo_no AS [行No] "
            tmp_sql = tmp_sql & " 					,jyuyo_name AS [タイトル名] "
            tmp_sql = tmp_sql & " 					,jyuyo_lstname AS [内容]	 "
            tmp_sql = tmp_sql & " 					,CASE "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '所有権以外の権利の有無' THEN '物件権利情報への移行時に判別して移行します' "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '所有権以外の権利の種類' THEN '物件権利情報への移行時に判別して移行します' "
            tmp_sql = tmp_sql & " 						ELSE '革命10に移行先が存在しません' "
            tmp_sql = tmp_sql & " 					 END AS [判定結果] "
            tmp_sql = tmp_sql & " 				FROM bk_jyuyo AS BKJ "
            tmp_sql = tmp_sql & " 				LEFT JOIN "
            tmp_sql = tmp_sql & " 					( "
            tmp_sql = tmp_sql & " 						SELECT * FROM "
            tmp_sql = tmp_sql & " 						( "
            tmp_sql = tmp_sql & " 							SELECT * FROM m_jyuyo "
            tmp_sql = tmp_sql & " 							WHERE jyuyo_kbn = 1 "
            tmp_sql = tmp_sql & " 						) AS VW "
            tmp_sql = tmp_sql & " 					) AS JYUYOTITLE "
            tmp_sql = tmp_sql & " 				ON BKJ.jyuyo_no = JYUYOTITLE.jyuyo_no "
            tmp_sql = tmp_sql & " 				WHERE jyuyo_name IN "
            tmp_sql = tmp_sql & " 									( "
            tmp_sql = tmp_sql & " 										 '所有権以外の権利の有無' "
            tmp_sql = tmp_sql & " 										,'所有権以外の権利の種類' "
            tmp_sql = tmp_sql & " 										,'建物の登録名義人調査日' "
            tmp_sql = tmp_sql & " 										,'建物の登録名義人と貸主が異なる理由' "
            tmp_sql = tmp_sql & " 										,'未完成物件' "
            tmp_sql = tmp_sql & " 										,'損害賠償額の予定又は違約金に関する定め' "
            tmp_sql = tmp_sql & " 										,'保全措置を行う機関' "
            tmp_sql = tmp_sql & " 										,'備考（契約の種類・期間・更新等）' "
            tmp_sql = tmp_sql & " 										,'公正証書にする場合の費用負担' "
            tmp_sql = tmp_sql & " 										,'備考(定期借家契約の場合)' "
            tmp_sql = tmp_sql & " 										,'敷金等の清算に関する事項' "
            tmp_sql = tmp_sql & " 										,'ﾏﾝｼｮﾝの管理の適正化の推進に関する法律' "
            tmp_sql = tmp_sql & " 									) "
            tmp_sql = tmp_sql & " 			) AS NOTCV "
            tmp_sql = tmp_sql & " 			/*標準タイトルだが移行先が存在しないデータ end*/ "
            tmp_sql = tmp_sql & " 			UNION "
            tmp_sql = tmp_sql & " 			/*標準タイトルだが内容が不適合のデータ sta*/ "
            tmp_sql = tmp_sql & " 			SELECT * FROM "
            tmp_sql = tmp_sql & " 			( "
            tmp_sql = tmp_sql & " 				SELECT "
            tmp_sql = tmp_sql & " 					 bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 					,jyuyo_no AS [行No] "
            tmp_sql = tmp_sql & " 					,jyuyo_name AS [タイトル名] "
            tmp_sql = tmp_sql & " 					,jyuyo_lstname AS [内容] "
            tmp_sql = tmp_sql & " 					,CASE "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '所有権にかかる権利の有無' THEN "
            tmp_sql = tmp_sql & " 							CASE "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '有' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '無' THEN '' "
            tmp_sql = tmp_sql & " 								ELSE '内容が標準以外の値になっているため移行できません' "
            tmp_sql = tmp_sql & " 							END "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '所有権にかかる権利の種類' THEN "
            tmp_sql = tmp_sql & " 							CASE "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '仮登記(所有権移転)' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '仮登記(所有権移転請求権)' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '仮差押' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '仮処分' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '差押(含む参加差押)' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '買戻特約' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '予告登記' THEN '' "
            tmp_sql = tmp_sql & " 								ELSE '内容が標準以外の値になっているため移行できません' "
            tmp_sql = tmp_sql & " 							END "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '法令に基づく制限の概要（法令名）' THEN "
            tmp_sql = tmp_sql & " 							CASE "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '新住宅市街地開発法' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '新都市基盤整備法' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '流通業務市街地整備法' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '農地法' THEN '' "
            tmp_sql = tmp_sql & " 								ELSE '内容が標準以外の値になっているため移行できません' "
            tmp_sql = tmp_sql & " 							END "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '敷地利用関係の種類(敷地が借地の場合)' THEN "
            tmp_sql = tmp_sql & " 							CASE "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '借地権(旧法)' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '借地権(新法)' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '定期借地権' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '使用賃借' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '一時賃貸借' THEN '' "
            tmp_sql = tmp_sql & " 								ELSE '内容が標準以外の値になっているため移行できません' "
            tmp_sql = tmp_sql & " 							END "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '契約期間(敷地が賃借の場合）' THEN 'コンバート時に内容が標準と同形式か判別して移行します' "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '宅地造成等規正法' THEN "
            tmp_sql = tmp_sql & " 							CASE "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '造成宅地防災区域内' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '造成宅地防災区域外' THEN '' "
            tmp_sql = tmp_sql & " 								ELSE '内容が標準以外の値になっているため移行できません' "
            tmp_sql = tmp_sql & " 							END "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '土砂災害防止対策推進法' THEN "
            tmp_sql = tmp_sql & " 							CASE "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '土砂災害警戒区域内' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '土砂災害警戒区域外' THEN '' "
            tmp_sql = tmp_sql & " 								ELSE '内容が標準以外の値になっているため移行できません' "
            tmp_sql = tmp_sql & " 							END "
            tmp_sql = tmp_sql & " 						WHEN jyuyo_name = '津波防災地域づくりに関する法律' THEN "
            tmp_sql = tmp_sql & " 							CASE "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '津波災害警戒区域内' THEN '' "
            tmp_sql = tmp_sql & " 								WHEN jyuyo_lstname = '津波災害警戒区域外' THEN '' "
            tmp_sql = tmp_sql & " 								ELSE '内容が標準以外の値になっているため移行できません' "
            tmp_sql = tmp_sql & " 							END "
            tmp_sql = tmp_sql & " 					 END AS [判定結果] "
            tmp_sql = tmp_sql & " 				FROM "
            tmp_sql = tmp_sql & " 				( "
            tmp_sql = tmp_sql & " 					SELECT  "
            tmp_sql = tmp_sql & " 						 bk_no "
            tmp_sql = tmp_sql & " 						,BKJ.jyuyo_no "
            tmp_sql = tmp_sql & " 						,jyuyo_name "
            tmp_sql = tmp_sql & " 						,jyuyo_lstname "
            tmp_sql = tmp_sql & " 					FROM bk_jyuyo AS BKJ "
            tmp_sql = tmp_sql & " 					LEFT JOIN "
            tmp_sql = tmp_sql & " 						( "
            tmp_sql = tmp_sql & " 							SELECT * FROM "
            tmp_sql = tmp_sql & " 							( "
            tmp_sql = tmp_sql & " 								SELECT * FROM m_jyuyo "
            tmp_sql = tmp_sql & " 								WHERE jyuyo_kbn = 1 "
            tmp_sql = tmp_sql & " 							) AS VW "
            tmp_sql = tmp_sql & " 						) AS JYUYOTITLE "
            tmp_sql = tmp_sql & " 					ON BKJ.jyuyo_no = JYUYOTITLE.jyuyo_no "
            tmp_sql = tmp_sql & " 					WHERE jyuyo_name IN "
            tmp_sql = tmp_sql & " 										( "
            tmp_sql = tmp_sql & " 											 '所有権にかかる権利の有無' "
            tmp_sql = tmp_sql & " 											,'所有権にかかる権利の種類' "
            tmp_sql = tmp_sql & " 											,'法令に基づく制限の概要（法令名）' "
            tmp_sql = tmp_sql & " 											/*,'法令に基づく制限の概要'*/ "
            tmp_sql = tmp_sql & " 											,'敷地利用関係の種類(敷地が借地の場合)' "
            tmp_sql = tmp_sql & " 											,'契約期間(敷地が賃借の場合）' "
            tmp_sql = tmp_sql & " 											/*,'備考(敷地が借地の場合)'*/ "
            tmp_sql = tmp_sql & " 											,'宅地造成等規正法' "
            tmp_sql = tmp_sql & " 											,'土砂災害防止対策推進法' "
            tmp_sql = tmp_sql & " 											,'津波防災地域づくりに関する法律'	 "
            tmp_sql = tmp_sql & " 										) "
            tmp_sql = tmp_sql & " 					AND jyuyo_lstname <> '' "
            tmp_sql = tmp_sql & " 				) AS VW "
            tmp_sql = tmp_sql & " 			) AS NAIYO "
            tmp_sql = tmp_sql & " 		/*標準タイトルだが内容が不適合のデータ end*/ "
            tmp_sql = tmp_sql & " 		) AS TOTAL "
            tmp_sql = tmp_sql & " 		WHERE [判定結果] <> '' "
            tmp_sql = tmp_sql & " 	) AS VW1 "
            tmp_sql = tmp_sql & wherestr
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 家主控除発生月抽出クエリ '20160707 事前作業リスト出力処理の追加_家主控除発生月
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_HasseiOwList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,控除管理No,入金項目No,入金項目名,開始年月,発生間隔"

            Dim tmp_sql As String = ""
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[控除管理No],[入金項目No] ")

            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 MEI.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,MEI.kjkn_no AS [控除管理No] "
            tmp_sql = tmp_sql & " 		,MEI.nkin_no AS [入金項目No] "
            tmp_sql = tmp_sql & " 		,KOM.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 		,KOJ.start_ym AS [開始年月] "
            tmp_sql = tmp_sql & " 		,MEI.hasei_mm AS [発生間隔] "
            tmp_sql = tmp_sql & " 	FROM bk_kojyomei MEI "
            tmp_sql = tmp_sql & " 	LEFT JOIN bk_kojyo KOJ ON KOJ.bk_no=MEI.bk_no AND KOJ.kjkn_no=MEI.kjkn_no "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin KOM ON KOM.nkin_no=MEI.nkin_no "
            tmp_sql = tmp_sql & " 	WHERE MEI.[hasei_mm]>12 OR MEI.[hasei_mm] IN ('5','7','8','9','10','11') "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 変動費発生月抽出クエリ '20160707 事前作業リスト出力処理の追加_変動費発生月
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_HasseiHenList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,部屋番号,契約No,更新No,月区分,入金項目コード,入金項目名,発生間隔"

            Dim tmp_sql As String = ""
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[部屋番号],[契約No],[更新No],[月区分],[入金項目コード] ")

            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 KYS.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,KYS.hy_no AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,KYS.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,KYS.ko_no AS [更新No] "
            tmp_sql = tmp_sql & " 		,KYS.tuki_kbn AS [月区分] "
            tmp_sql = tmp_sql & " 		,KYS.nkin_no AS [入金項目コード] "
            tmp_sql = tmp_sql & " 		,MN.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 		,KYS.zj_sqkan AS [発生間隔]  "
            tmp_sql = tmp_sql & " 	FROM ky_sqdata KYS "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin MN ON MN.nkin_no=KYS.nkin_no "
            tmp_sql = tmp_sql & " 	WHERE KYS.zj_sqkan>12 OR KYS.zj_sqkan IN ('5','7','8','9','10','11') "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 送金先未設定抽出クエリ '20160707 事前作業リスト出力処理の追加_送金先未設定
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_YanusoList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,物件名"

            Dim tmp_sql As String = ""
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No] ")

            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT DISTINCT "
            tmp_sql = tmp_sql & " 		 BKM.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,BKM.bk_name AS [物件名] "
            tmp_sql = tmp_sql & " 		/* "
            tmp_sql = tmp_sql & " 		,BKS.bk_no "
            tmp_sql = tmp_sql & " 		,BKS.so_no "
            tmp_sql = tmp_sql & " 		,OWS.so_name "
            tmp_sql = tmp_sql & " 		,BKS.kn_no "
            tmp_sql = tmp_sql & " 		,BKS.so_no "
            tmp_sql = tmp_sql & " 		,BKS.sokoza_no "
            tmp_sql = tmp_sql & " 		*/ "
            tmp_sql = tmp_sql & " 	FROM bk_mst BKM "
            tmp_sql = tmp_sql & " 	LEFT JOIN bk_sokin BKS ON BKS.bk_no=BKM.bk_no "
            tmp_sql = tmp_sql & " 	/*LEFT JOIN m_yanu_so OWS ON OWS.so_no=BKS.so_no*/ "
            tmp_sql = tmp_sql & " 	WHERE BKS.bk_no IS NULL "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 修繕負担者抽出クエリ '20160707 事前作業リスト出力処理の追加_修繕負担者
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_SzenKyshutanList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,部屋番号,リフォームID,入金項目コード,入金項目名"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[部屋番号],[リフォームID],[入金項目コード] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 RI.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,CASE RI.hy_composite "
            tmp_sql = tmp_sql & " 			WHEN '999999999' THEN '-' "
            tmp_sql = tmp_sql & " 			ELSE RI.hy_composite "
            tmp_sql = tmp_sql & " 		 END AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,RI.reform_id AS [リフォームID] "
            tmp_sql = tmp_sql & " 		,RI.nkin_no AS [入金項目コード] "
            tmp_sql = tmp_sql & " 		,MN.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 	FROM reform_item RI "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin MN ON MN.nkin_no=RI.nkin_no "
            tmp_sql = tmp_sql & " 	WHERE hy_composite='999999999' AND total_kari>0 "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 部屋鍵情報抽出クエリ '20160725 事前作業へ鍵情報の追加
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_HyKagiList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,部屋番号,鍵タイトル,鍵名No,鍵内容,鍵本数"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[部屋番号],[鍵名No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 HKAGI.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,HKAGI.hy_no AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,MBIKO.biko_titl AS [鍵タイトル] "
            tmp_sql = tmp_sql & " 		,HKAGI.kagimei_no AS [鍵名No] "
            tmp_sql = tmp_sql & " 		,HKAGI.kagi_no AS [鍵内容] "
            tmp_sql = tmp_sql & " 		,HKAGI.kagi_honsu AS [鍵本数]  "
            tmp_sql = tmp_sql & " 	FROM hy_kagi HKAGI "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_biko MBIKO ON MBIKO.biko_no=HKAGI.kagimei_no "
            tmp_sql = tmp_sql & " 	WHERE MBIKO.biko_kbn=4 AND NOT (HKAGI.kagi_no IS NULL OR HKAGI.kagi_honsu IS NULL) "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 契約鍵情報抽出クエリ '20160725 事前作業へ鍵情報の追加
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_KyKagiList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,部屋番号,契約No,鍵タイトル,鍵名No,鍵内容,鍵本数"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[部屋番号],[契約No],[鍵名No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 KKAGI.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,KKAGI.hy_no AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,KKAGI.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,MBIKO.biko_titl AS [鍵タイトル] "
            tmp_sql = tmp_sql & " 		,KKAGI.kagimei_no AS [鍵名No] "
            tmp_sql = tmp_sql & " 		,KKAGI.kagi_no AS [鍵内容] "
            tmp_sql = tmp_sql & " 		,KKAGI.kagi_honsu AS [鍵本数]  "
            tmp_sql = tmp_sql & " 	FROM ky_kagi KKAGI "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_biko MBIKO ON MBIKO.biko_no=KKAGI.kagimei_no "
            tmp_sql = tmp_sql & " 	WHERE biko_kbn=4 AND NOT (KKAGI.kagi_no IS NULL OR KKAGI.kagi_honsu IS NULL) "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 口座実情報未登録データ抽出クエリ '20160829 事前作業_口座実情報未登録データ抽出処理を追加 -add
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_NonJisyaKozaList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "口座取得元,口座設定No,口座設定名称"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [口座取得元],[口座設定No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 [口座取得元] "
            tmp_sql = tmp_sql & " 		,fkom_no AS [口座設定No] "
            tmp_sql = tmp_sql & " 		,fkom_name AS [口座設定名称] "
            tmp_sql = tmp_sql & " 		/*,**/ "
            tmp_sql = tmp_sql & " 	FROM "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		SELECT "
            tmp_sql = tmp_sql & " 			 '振込依頼人マスタ' AS [口座取得元] "
            tmp_sql = tmp_sql & " 			,fkom_no "
            tmp_sql = tmp_sql & " 			,fkom_name "
            tmp_sql = tmp_sql & " 			,kinyu_no "
            tmp_sql = tmp_sql & " 			,ten_no "
            tmp_sql = tmp_sql & " 			,kosyu_no "
            tmp_sql = tmp_sql & " 			,koza_no "
            tmp_sql = tmp_sql & " 			,koza_meigi "
            tmp_sql = tmp_sql & " 			,koza_kana "
            tmp_sql = tmp_sql & " 			,'' AS post_kigo "
            tmp_sql = tmp_sql & " 			,'' AS post_n "
            tmp_sql = tmp_sql & " 			,'' AS post_bango "
            tmp_sql = tmp_sql & " 			,'' AS biko "
            tmp_sql = tmp_sql & " 		FROM m_furi_irai "
            tmp_sql = tmp_sql & " 		UNION "
            tmp_sql = tmp_sql & " 		SELECT "
            tmp_sql = tmp_sql & " 			 '振込先口座マスタ' AS [口座取得元] "
            tmp_sql = tmp_sql & " 			,fkom_no "
            tmp_sql = tmp_sql & " 			,fkom_name "
            tmp_sql = tmp_sql & " 			,kinyu_no "
            tmp_sql = tmp_sql & " 			,ten_no "
            tmp_sql = tmp_sql & " 			,kosyu_no "
            tmp_sql = tmp_sql & " 			,koza_no "
            tmp_sql = tmp_sql & " 			,koza_meigi "
            tmp_sql = tmp_sql & " 			,koza_kana "
            tmp_sql = tmp_sql & " 			,post_kigo "
            tmp_sql = tmp_sql & " 			,post_n "
            tmp_sql = tmp_sql & " 			,post_bango "
            tmp_sql = tmp_sql & " 			,biko "
            tmp_sql = tmp_sql & " 		FROM m_furi_koza AS MFK "
            tmp_sql = tmp_sql & " 		/*家主口座と一致しないデータに対して紐付けを行う*/ "
            tmp_sql = tmp_sql & " 		WHERE NOT EXISTS "
            tmp_sql = tmp_sql & " 			( "
            tmp_sql = tmp_sql & " 				SELECT * FROM m_yanu_so_koza AS OWKOZA "
            tmp_sql = tmp_sql & " 				WHERE MFK.kinyu_no = OWKOZA.kinyu_no "
            tmp_sql = tmp_sql & " 				AND   MFK.ten_no = OWKOZA.ten_no "
            tmp_sql = tmp_sql & " 				AND   MFK.kosyu_no = OWKOZA.kosyu_no "
            tmp_sql = tmp_sql & " 				AND   MFK.koza_no = OWKOZA.koza_no "
            tmp_sql = tmp_sql & " 			) "
            tmp_sql = tmp_sql & " 		UNION "
            tmp_sql = tmp_sql & " 		SELECT "
            tmp_sql = tmp_sql & " 			 '口座振替マスタ' AS [口座取得元] "
            tmp_sql = tmp_sql & " 			,fkae_no "
            tmp_sql = tmp_sql & " 			,fkae_name "
            tmp_sql = tmp_sql & " 			,kinyu_no "
            tmp_sql = tmp_sql & " 			,ten_no "
            tmp_sql = tmp_sql & " 			,kosyu_no "
            tmp_sql = tmp_sql & " 			,koza_no "
            tmp_sql = tmp_sql & " 			,koza_meigi "
            tmp_sql = tmp_sql & " 			,koza_kana "
            tmp_sql = tmp_sql & " 			,'' AS post_kigo "
            tmp_sql = tmp_sql & " 			,'' AS post_n "
            tmp_sql = tmp_sql & " 			,'' AS post_bango "
            tmp_sql = tmp_sql & " 			,'' AS biko "
            tmp_sql = tmp_sql & " 		FROM m_koza_furi "
            tmp_sql = tmp_sql & " 	) AS VW  "
            tmp_sql = tmp_sql & " 	WHERE ISNULL(kinyu_no,0) = 0 OR ISNULL(ten_no,0) = 0 OR ISNULL(koza_no,0) = 0 "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 締、送金日の間隔が1ヶ月以上のデータ抽出クエリ '20160829 事前作業_送金予定日リスト出力機能を追加
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_SimeSokinList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,管理No,物件名,締区分1,締日1,送金区分1,送金日1,締区分2,締日2,送金区分2,送金日2,締区分3,締日3,送金区分3,送金日3,締区分4,締日4,送金区分4,送金日4,締区分5,締日5,送金区分5,送金日5,締日エラー内容,送金日エラー内容1,送金日エラー内容2"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[管理No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,kn_no AS [管理No] "
            tmp_sql = tmp_sql & " 		,bk_name AS [物件名] "
            tmp_sql = tmp_sql & " 		,CASE WHEN sime_kbn1 = -2 THEN '前々月' WHEN sime_kbn1 = -1 THEN '前月' WHEN sime_kbn1 = 0 THEN '当月' WHEN sime_kbn1 = 1 THEN '翌月' WHEN sime_kbn1 = 2 THEN '翌々月' ELSE '' END AS [締区分1] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,sime1),'') AS [締日1] "
            tmp_sql = tmp_sql & " 		,CASE WHEN som_kbn1 = 0 THEN '当月' WHEN som_kbn1 = 1 THEN '翌月' WHEN som_kbn1 = 2 THEN '翌々月' WHEN som_kbn1 = 3 THEN '翌3々月' WHEN som_kbn1 = 4 THEN '翌4々月' WHEN som_kbn1 = 5 THEN '翌5々月' ELSE '' END AS [送金区分1] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,so_dd1),'') AS [送金日1] "
            tmp_sql = tmp_sql & " 		,CASE WHEN sime_kbn2 = -2 THEN '前々月' WHEN sime_kbn2 = -1 THEN '前月' WHEN sime_kbn2 = 0 THEN '当月' WHEN sime_kbn2 = 1 THEN '翌月' WHEN sime_kbn2 = 2 THEN '翌々月' ELSE '' END AS [締区分2] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,sime2),'') AS [締日2] "
            tmp_sql = tmp_sql & " 		,CASE WHEN som_kbn2 = 0 THEN '当月' WHEN som_kbn2 = 1 THEN '翌月' WHEN som_kbn2 = 2 THEN '翌々月' WHEN som_kbn2 = 3 THEN '翌3々月' WHEN som_kbn2 = 4 THEN '翌4々月' WHEN som_kbn2 = 5 THEN '翌5々月' ELSE '' END AS [送金区分2] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,so_dd2),'') AS [送金日2] "
            tmp_sql = tmp_sql & " 		,CASE WHEN sime_kbn3 = -2 THEN '前々月' WHEN sime_kbn3 = -1 THEN '前月' WHEN sime_kbn3 = 0 THEN '当月' WHEN sime_kbn3 = 1 THEN '翌月' WHEN sime_kbn3 = 2 THEN '翌々月' ELSE '' END AS [締区分3] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,sime3),'') AS [締日3] "
            tmp_sql = tmp_sql & " 		,CASE WHEN som_kbn3 = 0 THEN '当月' WHEN som_kbn3 = 1 THEN '翌月' WHEN som_kbn3 = 2 THEN '翌々月' WHEN som_kbn3 = 3 THEN '翌3々月' WHEN som_kbn3 = 4 THEN '翌4々月' WHEN som_kbn3 = 5 THEN '翌5々月' ELSE '' END AS [送金区分3] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,so_dd3),'') AS [送金日3] "
            tmp_sql = tmp_sql & " 		,CASE WHEN sime_kbn4 = -2 THEN '前々月' WHEN sime_kbn4 = -1 THEN '前月' WHEN sime_kbn4 = 0 THEN '当月' WHEN sime_kbn4 = 1 THEN '翌月' WHEN sime_kbn4 = 2 THEN '翌々月' ELSE '' END AS [締区分4] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,sime4),'') AS [締日4] "
            tmp_sql = tmp_sql & " 		,CASE WHEN som_kbn4 = 0 THEN '当月' WHEN som_kbn4 = 1 THEN '翌月' WHEN som_kbn4 = 2 THEN '翌々月' WHEN som_kbn4 = 3 THEN '翌3々月' WHEN som_kbn4 = 4 THEN '翌4々月' WHEN som_kbn4 = 5 THEN '翌5々月' ELSE '' END AS [送金区分4] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,so_dd4),'') AS [送金日4] "
            tmp_sql = tmp_sql & " 		,CASE WHEN sime_kbn5 = -2 THEN '前々月' WHEN sime_kbn5 = -1 THEN '前月' WHEN sime_kbn5 = 0 THEN '当月' WHEN sime_kbn5 = 1 THEN '翌月' WHEN sime_kbn5 = 2 THEN '翌々月' ELSE '' END AS [締区分5] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,sime5),'') AS [締日5] "
            tmp_sql = tmp_sql & " 		,CASE WHEN som_kbn5 = 0 THEN '当月' WHEN som_kbn5 = 1 THEN '翌月' WHEN som_kbn5 = 2 THEN '翌々月' WHEN som_kbn5 = 3 THEN '翌3々月' WHEN som_kbn5 = 4 THEN '翌4々月' WHEN som_kbn5 = 5 THEN '翌5々月' ELSE '' END AS [送金区分5] "
            tmp_sql = tmp_sql & " 		,ISNULL(CONVERT(VARCHAR,so_dd5),'') AS [送金日5] "
            tmp_sql = tmp_sql & " 		,[締日エラー内容] "
            tmp_sql = tmp_sql & " 		,[送金日エラー内容1] "
            tmp_sql = tmp_sql & " 		,[送金日エラー内容2] "
            tmp_sql = tmp_sql & " 	FROM "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		SELECT "
            tmp_sql = tmp_sql & " 			 KANRI1.bk_no "
            tmp_sql = tmp_sql & " 			,KANRI1.kn_no "
            tmp_sql = tmp_sql & " 			,BK.bk_name "
            tmp_sql = tmp_sql & " 			,KANRI1.sime_kbn1,KANRI1.sime1,KANRI1.som_kbn1,KANRI1.so_dd1 "
            tmp_sql = tmp_sql & " 			,KANRI2.sime_kbn2,KANRI2.sime2,KANRI2.som_kbn2,KANRI2.so_dd2 "
            tmp_sql = tmp_sql & " 			,KANRI3.sime_kbn3,KANRI3.sime3,KANRI3.som_kbn3,KANRI3.so_dd3 "
            tmp_sql = tmp_sql & " 			,KANRI4.sime_kbn4,KANRI4.sime4,KANRI4.som_kbn4,KANRI4.so_dd4 "
            tmp_sql = tmp_sql & " 			,KANRI5.sime_kbn5,KANRI5.sime5,KANRI5.som_kbn5,KANRI5.so_dd5 "
            tmp_sql = tmp_sql & " 			,CASE "
            tmp_sql = tmp_sql & " 				WHEN (KANRI2.sime_kbn2 - KANRI1.sime_kbn1 = 1 AND KANRI2.sime2 > KANRI1.sime1) OR (KANRI2.sime_kbn2 - KANRI1.sime_kbn1 >= 2) THEN '締日1と締日2が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				WHEN (KANRI3.sime_kbn3 - KANRI2.sime_kbn2 = 1 AND KANRI3.sime3 > KANRI2.sime2) OR (KANRI3.sime_kbn3 - KANRI2.sime_kbn2 >= 2) THEN '締日2と締日3が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				WHEN (KANRI4.sime_kbn4 - KANRI3.sime_kbn3 = 1 AND KANRI4.sime4 > KANRI3.sime3) OR (KANRI4.sime_kbn4 - KANRI3.sime_kbn3 >= 2) THEN '締日3と締日4が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				WHEN (KANRI5.sime_kbn5 - KANRI4.sime_kbn4 = 1 AND KANRI5.sime5 > KANRI4.sime4) OR (KANRI5.sime_kbn5 - KANRI4.sime_kbn4 >= 2) THEN '締日4と締日5が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				ELSE '' "
            tmp_sql = tmp_sql & " 			 END AS [締日エラー内容] "
            tmp_sql = tmp_sql & " 			,CASE "
            tmp_sql = tmp_sql & " 				WHEN (KANRI2.som_kbn2 - KANRI1.som_kbn1 = 1 AND KANRI2.so_dd2 > KANRI1.so_dd1) OR (KANRI2.som_kbn2 - KANRI1.som_kbn1 >= 2) THEN '送金日1と送金日2が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				WHEN (KANRI3.som_kbn3 - KANRI2.som_kbn2 = 1 AND KANRI3.so_dd3 > KANRI2.so_dd2) OR (KANRI3.som_kbn3 - KANRI2.som_kbn2 >= 2) THEN '送金日2と送金日3が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				WHEN (KANRI4.som_kbn4 - KANRI3.som_kbn3 = 1 AND KANRI4.so_dd4 > KANRI3.so_dd3) OR (KANRI4.som_kbn4 - KANRI3.som_kbn3 >= 2) THEN '送金日3と送金日4が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				WHEN (KANRI5.som_kbn5 - KANRI4.som_kbn4 = 1 AND KANRI5.so_dd5 > KANRI4.so_dd4) OR (KANRI5.som_kbn5 - KANRI4.som_kbn4 >= 2) THEN '送金日4と送金日5が1ヶ月以上' "
            tmp_sql = tmp_sql & " 				ELSE '' "
            tmp_sql = tmp_sql & " 			 END AS [送金日エラー内容1] "
            tmp_sql = tmp_sql & " 			/*20160901 不正データ抽出条件を追加 add sta*/ "
            tmp_sql = tmp_sql & " 			,CASE "
            tmp_sql = tmp_sql & " 				WHEN KANRI1.som_kbn1 = 0 AND KANRI1.sime1 > so_dd1 THEN '送金設定1の締、送金日が逆転' "
            tmp_sql = tmp_sql & " 				WHEN KANRI2.som_kbn2 = 0 AND KANRI2.sime2 > so_dd1 THEN '送金設定2の締、送金日が逆転' "
            tmp_sql = tmp_sql & " 				WHEN KANRI3.som_kbn3 = 0 AND KANRI3.sime3 > so_dd1 THEN '送金設定3の締、送金日が逆転' "
            tmp_sql = tmp_sql & " 				WHEN KANRI4.som_kbn4 = 0 AND KANRI4.sime4 > so_dd1 THEN '送金設定4の締、送金日が逆転' "
            tmp_sql = tmp_sql & " 				WHEN KANRI5.som_kbn5 = 0 AND KANRI5.sime5 > so_dd1 THEN '送金設定5の締、送金日が逆転' "
            tmp_sql = tmp_sql & " 				ELSE ''	 "
            tmp_sql = tmp_sql & " 			 END AS [送金日エラー内容2] "
            tmp_sql = tmp_sql & " 			/*20160901 不正データ抽出条件を追加 add end*/ "
            tmp_sql = tmp_sql & " 		FROM bk_kanri AS KANRI1 "
            tmp_sql = tmp_sql & " 		LEFT JOIN (SELECT bk_no,kn_no,sime_kbn2,sime2,som_kbn2,so_dd2 FROM bk_kanri) AS KANRI2 ON KANRI1.bk_no = KANRI2.bk_no "
            tmp_sql = tmp_sql & " 		LEFT JOIN (SELECT bk_no,kn_no,sime_kbn3,sime3,som_kbn3,so_dd3 FROM bk_kanri) AS KANRI3 ON KANRI1.bk_no = KANRI3.bk_no "
            tmp_sql = tmp_sql & " 		LEFT JOIN (SELECT bk_no,kn_no,sime_kbn4,sime4,som_kbn4,so_dd4 FROM bk_kanri) AS KANRI4 ON KANRI1.bk_no = KANRI4.bk_no "
            tmp_sql = tmp_sql & " 		LEFT JOIN (SELECT bk_no,kn_no,sime_kbn5,sime5,som_kbn5,so_dd5 FROM bk_kanri) AS KANRI5 ON KANRI1.bk_no = KANRI5.bk_no "
            tmp_sql = tmp_sql & " 		LEFT JOIN (SELECT bk_no,bk_name FROM bk_mst) AS BK ON KANRI1.bk_no = BK.bk_no "
            tmp_sql = tmp_sql & " 	) AS VW "
            tmp_sql = tmp_sql & " 	WHERE [締日エラー内容] <> '' OR [送金日エラー内容1] <> '' OR [送金日エラー内容2] <> '' "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 使用不可文字を判定するストアドプロシージャ削除 '20160829 口座名義カナチェック機能の追加
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Get_UseQry_Kozameigikana_DropStPr()

            Dim tmp_sql As String = DBQuery.Qry_DropFnInfo("fn_Chk_KozaKana")
            Dim tmp_cnt As Integer = 0
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sql, tmp_cnt)

        End Sub

        ''' <summary>
        ''' 使用不可文字を判定するストアドプロシージャ作成 '20160829 口座名義カナチェック機能の追加
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Get_UseQry_Kozameigikana_CreateStPr()

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " CREATE FUNCTION dbo.[fn_Chk_KozaKana](@value NVARCHAR(MAX)) RETURNS VARCHAR(MAX) "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " BEGIN "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	DECLARE @result VARCHAR(MAX) "
            tmp_sql = tmp_sql & " 	DECLARE @cntii int "
            tmp_sql = tmp_sql & " 	DECLARE @uniCode int "
            tmp_sql = tmp_sql & " 	DECLARE @tmp_str VARCHAR(300) "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " 	DECLARE @errmsg_zenkaku VARCHAR(MAX) "
            tmp_sql = tmp_sql & " 	DECLARE @errmsg_enablestr VARCHAR(MAX) "
            tmp_sql = tmp_sql & " 		 "
            tmp_sql = tmp_sql & " 	SET @result = '' "
            tmp_sql = tmp_sql & " 	SET @errmsg_zenkaku = '全角文字列が含まれています' "
            tmp_sql = tmp_sql & " 	SET @errmsg_enablestr = '使用不可文字「」が含まれています' "
            tmp_sql = tmp_sql & " 	SET @cntii = 1 "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " 	WHILE (@cntii <= LEN(@value)) "
            tmp_sql = tmp_sql & " 	BEGIN "
            tmp_sql = tmp_sql & " 		 "
            tmp_sql = tmp_sql & " 		SET @tmp_str = SUBSTRING(@value,@cntii,1) "
            tmp_sql = tmp_sql & " 		/*SET @uniCode = UNICODE(SUBSTRING(@value,@cntii,1))*/ "
            tmp_sql = tmp_sql & " 		 "
            tmp_sql = tmp_sql & " 		/*使用不可文字チェック*/ "
            tmp_sql = tmp_sql & " 		IF      REPLACE(@tmp_str,'@','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「@」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'`','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「`」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'!','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「!」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'" & """" & " ','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「" & """" & "」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'#','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「#」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'$','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「$」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'%','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「%」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'&','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「&」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'''','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「''''」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'*','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「*」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'+','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「+」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,':','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「:」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,';','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「;」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'[','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「[」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'{','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「{」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'<','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「<」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'|','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「|」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'=','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「=」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,']','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「]」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'}','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「}」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'>','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「>」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'^','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「^」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'~','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「~」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'?','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「?」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'_','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「_」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'｡','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「｡」') "
            tmp_sql = tmp_sql & " 		ELSE IF REPLACE(@tmp_str,'､','') = '' SET @result = REPLACE(@errmsg_enablestr,'「」','「､」') "
            tmp_sql = tmp_sql & " 		 "
            tmp_sql = tmp_sql & " 		/*全角文字チェック*/ "
            tmp_sql = tmp_sql & " 		/*ELSE IF NOT (@uniCode >= 65280) AND (@uniCode <= 65519) SET @result = @errmsg_zenkaku*/ "
            tmp_sql = tmp_sql & " 		ELSE IF DATALENGTH(@tmp_str) >= 2 SET @result = @errmsg_zenkaku "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 		/*判定結果*/ "
            tmp_sql = tmp_sql & " 		IF @result = '' "
            tmp_sql = tmp_sql & " 			SET @cntii = @cntii + 1 "
            tmp_sql = tmp_sql & " 		ELSE "
            tmp_sql = tmp_sql & " 			BREAK "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " 	END "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " 	RETURN @result "
            tmp_sql = tmp_sql & " 	 "
            tmp_sql = tmp_sql & " END "

            Dim tmp_cnt As Integer = 0
            DBExec.Exec_NonQuery(sqlcnnv7, tmp_sql, tmp_cnt)

        End Sub

        ''' <summary>
        ''' 使用不可文字を含んだ口座名義カナデータ抽出クエリ '20160829 口座名義カナチェック機能の追加
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_KozameigikanaList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "取得元区分,No,名称,口座No,口座名義カナ,エラー内容"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [取得元区分],[No],[口座No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	/*20160902 レビュー後指摘修正 add sta*/ "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '契約者口座情報' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,kys_no AS [No] "
            tmp_sql = tmp_sql & " 		,kys_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM kys_mst "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	/*20160902 レビュー後指摘修正 add end*/ "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '契約者照合用カナ1' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,kys_no AS [No] "
            tmp_sql = tmp_sql & " 		,kys_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,syogo_kana1 AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](syogo_kana1) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM kys_mst "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '契約者照合用カナ2' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,kys_no AS [No] "
            tmp_sql = tmp_sql & " 		,kys_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,syogo_kana2 AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](syogo_kana2) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM kys_mst "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '契約者照合用カナ3' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,kys_no AS [No] "
            tmp_sql = tmp_sql & " 		,kys_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,syogo_kana3 AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](syogo_kana3) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM kys_mst "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '家主口座情報' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,SOKOZA.so_no AS [No] "
            tmp_sql = tmp_sql & " 		,SO.so_name AS [名称] "
            tmp_sql = tmp_sql & " 		,CONVERT(VARCHAR,sokoza_no) AS [口座No] "
            tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM m_yanu_so_koza AS SOKOZA "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_yanu_so AS SO ON SOKOZA.so_no = SO.so_no "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '仲介・管理業者口座情報' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,gy_no AS [No] "
            tmp_sql = tmp_sql & " 		,gy_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM m_gy "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '家賃入金口座情報' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,fkom_no AS [No] "
            tmp_sql = tmp_sql & " 		,fkom_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM m_furi_koza "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '口座振替情報_名義カナ' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,fkae_no AS [No] "
            tmp_sql = tmp_sql & " 		,fkae_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM m_koza_furi "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '口座振替情報_振込依頼人カナ' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,fkae_no AS [No] "
            tmp_sql = tmp_sql & " 		,fkae_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,fkae_irkana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM m_koza_furi "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '振込依頼人情報_名義カナ' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,fkom_no AS [No] "
            tmp_sql = tmp_sql & " 		,fkom_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM m_furi_irai "
            tmp_sql = tmp_sql & " 	UNION "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 '振込依頼人情報_振込依頼人カナ' AS [取得元区分] "
            tmp_sql = tmp_sql & " 		,fkom_no AS [No] "
            tmp_sql = tmp_sql & " 		,fkom_name AS [名称] "
            tmp_sql = tmp_sql & " 		,'-' AS [口座No] "
            tmp_sql = tmp_sql & " 		,fkom_irkana AS [口座名義カナ] "
            tmp_sql = tmp_sql & " 		,DBO.[fn_Chk_KozaKana](koza_kana) AS [エラー内容] "
            tmp_sql = tmp_sql & " 	FROM m_furi_irai "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & " WHERE [エラー内容] <> '' "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

#End Region

#Region "事後作業関連処理"

#End Region

#Region "補助作業関連処理"

        ''' <summary>
        ''' 検証作業処理_件数表示 '20160707 検証作業リスト出力処理の追加
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_List_Cnt_Hojyo()

            Dim tmp_header As String = ""                     '使用しないが引数として必要であるため仮で用意する
            Dim cnt_outputstr As String = ""
            Dim tmp_sql As String = ""

            '未使用家主                                       '20160711 未使用データ削除_家主 -add
            tmp_sql = Me.Set_NotUseOwdata(0)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv10)
            End If
            Me.lblCntKiOpOw.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '未使用契約者                                     '20160711 未使用データ削除_契約者 -add
            tmp_sql = Me.Set_NotUseKysdata(0)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv10)
            End If
            Me.lblCntKiOpKys.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '分割入金未収分
            tmp_sql = Me.Get_UseQry_BunkatumisyuList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntBunkatumisyu.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '控除支払データ(自社登録分)
            tmp_sql = Me.Get_UseQry_KojyoshList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntKojyosh.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '修繕項目毎契約者送金率
            tmp_sql = Me.Get_UseQry_SzenKysSoritList(tmp_header, True)
            If tmp_sql <> "" Then
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv7)
            End If
            Me.lblCntSzenKysSorit.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            tmp_sql = ""
            cnt_outputstr = "0"

            '20160720 紐付設定ファイル出力機能の追加 -add sta
            '紐付設定ファイル
            Dim list_tmp As New List(Of String)
            Dim tmp_relitem() As String = relitemstr.Split("-")
            For cntii = 0 To UBound(tmp_relitem)
                If tmp_relitem(cntii).Trim <> "" Then
                    list_tmp.Add(tmp_relitem(cntii))
                End If
            Next
            Me.lblCntRelationSet.Text = list_tmp.Count.ToString("#,0") & " 件"
            '20160720 紐付設定ファイル出力機能の追加 -add end 

        End Sub

        ''' <summary>
        ''' 検証作業処理_件数表示(リネーム) '20160711 関連ファイルパスの件数表示処理
        ''' </summary>
        ''' <param name="searchstr"></param>
        ''' <remarks></remarks>
        Private Sub Set_List_Cnt_Hojyo_Rename(ByVal searchstr As String)

            Dim tmp_header As String = ""                       '使用しないが引数として必要であるため仮で用意する
            Dim cnt_outputstr As String = ""
            Dim tmp_sql As String = ""

            If searchstr = "" Then
                Exit Sub
            End If

            'クレーム関連ファイルパス
            If Me.chkRelRenameClaim.Checked Then
                '20160720 リネーム機能の修正 -chg sta
                '先頭から一致する文字列を抽出するように修正
                'tmp_sql = "SELECT COUNT(*) FROM (SELECT DISTINCT fullpath FROM claimdata_relfile WHERE fullpath LIKE '%" & searchstr & "%') AS VW"
                tmp_sql = "SELECT COUNT(*) FROM (SELECT DISTINCT fullpath FROM claimdata_relfile WHERE fullpath LIKE '" & searchstr & "%') AS VW"
                '20160720 リネーム機能の修正 -chg end
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv10)
                Me.lblRelRenameClaim.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            End If
            tmp_sql = ""
            cnt_outputstr = "0"

            '修繕関連ファイルパス
            If Me.chkRelRenameSzen.Checked Then
                '20160720 リネーム機能の修正 -chg sta
                '先頭から一致する文字列を抽出するように修正
                'tmp_sql = "SELECT COUNT(*) FROM (SELECT DISTINCT fullpath FROM szendata_relfile WHERE fullpath LIKE '%" & searchstr & "%') AS VW"
                tmp_sql = "SELECT COUNT(*) FROM (SELECT DISTINCT fullpath FROM szendata_relfile WHERE fullpath LIKE '" & searchstr & "%') AS VW"
                '20160720 リネーム機能の修正 -chg end
                cnt_outputstr = DBExec.Exec_Scalar(tmp_sql, sqlcnnv10)
                Me.lblRelRenameSzen.Text = Int32.Parse(cnt_outputstr).ToString("#,0") & " 件"
            End If
            tmp_sql = ""
            cnt_outputstr = "0"

        End Sub

        ''' <summary>
        ''' 検証作業処理_関連ファイルリネーム処理 '20160707 関連ファイルリネーム処理の追加
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Rename_RelFileName()

            Dim tmp_msgbf As String = "「変換前」"
            Dim tmp_msgaf As String = "「変換後」"
            Dim tmp_msgpostblank As String = "の値を入力して下さい。"
            Dim tmp_msgposterr As String = "に設定されたフォルダが存在しません。再度設定して下さい。"
            Dim tmp_msgmain As String = ""
            Dim bfrelname As String = Me.txtBfRelRename.Text
            Dim afrelname As String = Me.txtAfRelRename.Text

            'クレーム/修繕チェック判別
            If Me.chkRelRenameClaim.Checked = False And Me.chkRelRenameSzen.Checked = False Then
                MsgResult = MessageBox.Show("置換する関連ファイルを選択して下さい。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            '入力値有無判別
            If bfrelname.Trim = "" And afrelname.Trim = "" Then
                tmp_msgmain = tmp_msgbf & tmp_msgaf
            ElseIf bfrelname = "" Then
                tmp_msgmain = tmp_msgbf
            ElseIf afrelname = "" Then
                tmp_msgmain = tmp_msgaf
            End If

            If tmp_msgmain <> "" Then
                MsgResult = MessageBox.Show(tmp_msgmain & tmp_msgpostblank, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            '入力値形式チェック
            If EtcMethod.Chk_DirExist(bfrelname) = False And EtcMethod.Chk_DirExist(afrelname) = False Then
                tmp_msgmain = tmp_msgbf & tmp_msgaf
            ElseIf EtcMethod.Chk_DirExist(bfrelname) = False Then
                tmp_msgmain = tmp_msgbf
            ElseIf EtcMethod.Chk_DirExist(afrelname) = False Then
                tmp_msgmain = tmp_msgaf
            End If

            If tmp_msgmain <> "" Then
                MsgResult = MessageBox.Show(tmp_msgmain & tmp_msgposterr, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            '確認メッセージ
            MsgResult = MessageBox.Show(MSG_RELFILE_RENAMESTA, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                Exit Sub
            End If

            'リネーム処理
            Dim tmp_cnt As Integer = 0
            If Me.chkRelRenameClaim.Checked Then
                '20160720 リネーム機能の修正 -chg sta
                '先頭から一致する場合のみ置換対象とする
                'Dim tmp_sql As String = " UPDATE claimdata_relfile SET fullpath = REPLACE(fullpath,'" & bfrelname & "','" & afrelname & "') "
                Dim tmp_sql As String = " UPDATE claimdata_relfile SET fullpath = REPLACE(fullpath,'" & bfrelname & "','" & afrelname & "') WHERE fullpath LIKE '" & bfrelname & "%' "
                '20160720 リネーム機能の修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
            End If
            If Me.chkRelRenameSzen.Checked Then
                '20160720 リネーム機能の修正 -chg sta
                '先頭から一致する場合のみ置換対象とする
                'Dim tmp_sql As String = " UPDATE szendata_relfile SET fullpath = REPLACE(fullpath,'" & bfrelname & "','" & afrelname & "') "
                Dim tmp_sql As String = " UPDATE szendata_relfile SET fullpath = REPLACE(fullpath,'" & bfrelname & "','" & afrelname & "') WHERE fullpath LIKE '" & bfrelname & "%' "
                '20160720 リネーム機能の修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
            End If

            '完了メッセージ
            MsgResult = MessageBox.Show(MSG_RELFILE_RENAMEEND, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        End Sub

        ''' <summary>
        ''' 未使用データの削除処理 '20160711 未使用データ削除
        ''' </summary>
        ''' <param name="deldata"></param>
        ''' <remarks></remarks>
        Private Sub Main_NotUseData_Delete(ByVal deldata As String)

            '出力対象項目をメッセージ用に成形
            Dim tmp_komk As String = "未使用の" & deldata & "データ"

            '削除確認メッセージ
            MsgResult = MessageBox.Show(tmp_komk & "を削除します。よろしいですか？" & vbCrLf & "(削除したデータは元に戻すことができません)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                Exit Sub
            End If

            '件数抽出・リスト出力・削除処理
            Select Case deldata
                Case "家主"
                    Call Me.Set_NotUseOwdata(2)
                Case "契約者"
                    Call Me.Set_NotUseKysdata(2)
            End Select

            '完了メッセージ
            MsgResult = MessageBox.Show(tmp_komk & "の削除が完了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        End Sub

        ''' <summary>
        ''' 使用されている家主データを仮テーブルへ挿入 '20160711 未使用データ削除_家主
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_OwInsert(ByVal tblname As String)

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " INSERT INTO " & tblname
            tmp_sql = tmp_sql & " 	SELECT ow_no FROM  "
            tmp_sql = tmp_sql & "             	( "
            tmp_sql = tmp_sql & "             		/*物件情報 所有者情報(一棟所有)*/ "
            tmp_sql = tmp_sql & "             		SELECT DISTINCT kasi1_ow_no AS ow_no FROM bkdata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT kasi1_ow_no IS NULL "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT syo1_ow_no AS ow_no FROM bkdata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT syo1_ow_no IS NULL "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT kasi2_ow_no AS ow_no FROM bkdata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT kasi2_ow_no IS NULL "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT syo2_ow_no AS ow_no FROM bkdata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT syo2_ow_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*物件情報 修繕維持管理連絡先情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT syuzenijikanri_kasino AS ow_no FROM bkdata_szeniji "
            tmp_sql = tmp_sql & " 					WHERE NOT syuzenijikanri_kasino IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & "             		/*部屋情報 所有者情報(区分所有)*/ "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT kasi1_ow_no AS ow_no FROM hydata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT kasi1_ow_no IS NULL "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT syo1_ow_no AS ow_no FROM hydata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT syo1_ow_no IS NULL "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT kasi2_ow_no AS ow_no FROM hydata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT kasi2_ow_no IS NULL "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT syo2_ow_no AS ow_no FROM hydata_syo  "
            tmp_sql = tmp_sql & "             		WHERE NOT syo2_ow_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*部屋情報 修繕維持管理連絡先情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT syuzenijikanri_kasino AS ow_no FROM hydata_szeniji "
            tmp_sql = tmp_sql & " 					WHERE NOT syuzenijikanri_kasino IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*契約情報 契約解約情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT szen_owno AS ow_no FROM kydata_kai "
            tmp_sql = tmp_sql & " 					WHERE NOT szen_owno IS NULL "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT szen_owsqsakino AS ow_no FROM kydata_kai "
            tmp_sql = tmp_sql & " 					WHERE NOT szen_owsqsakino IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*請求情報 預り金情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT stakeholder_no AS ow_no FROM azukanri "
            tmp_sql = tmp_sql & " 					WHERE stakeholder_kbn = 200 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*請求情報 出納情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT stakeholder_no AS ow_no FROM suikanri "
            tmp_sql = tmp_sql & " 					WHERE stakeholder_kbn = 200 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*請求情報 滞納情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT sosaki_no AS ow_no FROM unyotainodata "
            tmp_sql = tmp_sql & " 					WHERE NOT sosaki_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*請求情報 請求情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT shsaki_no AS ow_no FROM sqdata "
            tmp_sql = tmp_sql & " 					WHERE NOT shsaki_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*請求情報 家主請求控除情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT sosaki_no AS ow_no FROM kjdata "
            tmp_sql = tmp_sql & " 					WHERE NOT sosaki_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*クレーム情報 クレーム基本情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT ow_no AS ow_no FROM claimdata "
            tmp_sql = tmp_sql & " 					WHERE NOT ow_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*修繕情報 修繕基本情報*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT ow_no AS ow_no FROM szendata "
            tmp_sql = tmp_sql & " 					WHERE NOT ow_no IS NULL "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT ow_sqsakino AS ow_no FROM szendata "
            tmp_sql = tmp_sql & " 					WHERE NOT ow_sqsakino IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & "             		/*家賃入金口座マスタ*/ "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT ow_no FROM m_yatinkoza  "
            tmp_sql = tmp_sql & "             		WHERE NOT ow_no IS NULL "
            tmp_sql = tmp_sql & "             		/*送金ルール 送金先情報*/ "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT so_ow_no AS ow_no FROM sorule_sosaki  "
            tmp_sql = tmp_sql & "             		WHERE NOT so_ow_no IS NULL "
            tmp_sql = tmp_sql & "             		/*送金ルール 送金ルール情報*/ "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT so_no AS ow_no FROM sorule_nk_cmrule  "
            tmp_sql = tmp_sql & "             		WHERE NOT so_no IS NULL "
            tmp_sql = tmp_sql & "             		/*運用開始時未納・滞納金データ*/ "
            tmp_sql = tmp_sql & "             		UNION ALL SELECT DISTINCT sosaki_no AS ow_no FROM unyotainodata  "
            tmp_sql = tmp_sql & "             		WHERE NOT sosaki_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*仕訳データ 念の為*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT stakeholder_no AS ow_no FROM siwakedata "
            tmp_sql = tmp_sql & " 					WHERE stakeholder_kbn = 200 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 					/*口座毎勘定科目マスター 念の為*/ "
            tmp_sql = tmp_sql & " 					UNION ALL SELECT DISTINCT stakeholder_no AS ow_no FROM m_kr_kozakanjo "
            tmp_sql = tmp_sql & " 					WHERE stakeholder_type = 200 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & "             	) AS OWTOTAL "
            tmp_sql = tmp_sql & "             	WHERE ow_no > 0 "
            tmp_sql = tmp_sql & "             	GROUP BY ow_no /*ORDER BY ow_no*/ "

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 未使用家主データ処理(抽出/削除) '20160711 未使用データ削除_家主
        ''' </summary>
        ''' <param name="typeno"></param>
        ''' <param name="headergrp"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_NotUseOwdata(ByVal typeno As Integer, Optional ByRef headergrp As String = "") As String

            Dim rtn_str As String = ""
            Dim tmp_tablename As String = "ow_dellst"
            Dim tmp_cnt As Integer = 0

            '-----------------------------------------------
            ' 使用されている家主データ格納用仮テーブル作成
            '-----------------------------------------------
            '初期化
            Dim tmpdb_dropqry As String = DBQuery.Qry_DropInfo(tmp_tablename, True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_dropqry, tmp_cnt)

            '仮テーブル作成
            Dim tmpdb_createqry As String = " CREATE TABLE " & tmp_tablename & " (ow_no INT NOT NULL) "
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_createqry, tmp_cnt)

            '-----------------------------------------------
            ' 使用されている家主データを格納
            '-----------------------------------------------
            Dim tmpdb_insertqry As String = Me.Get_UseQry_OwInsert(tmp_tablename)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_insertqry, tmp_cnt)

            '-----------------------------------------------
            ' 抽出(件数・リスト出力)/削除
            '-----------------------------------------------
            '20161014 改善対応：返却されない場合がある為修正 -chg sta
            'Select Case typeno
            '    Case 0  '件数抽出
            '        rtn_str = Me.Get_UseQry_OwSelect(tmp_tablename, headergrp, True)
            '        Return rtn_str
            '    Case 1  'リスト出力
            '        rtn_str = Me.Get_UseQry_OwSelect(tmp_tablename, headergrp)
            '        Return rtn_str
            '    Case 2  '削除
            '        Call Me.OwDelete(tmp_tablename)
            'End Select

            ''-----------------------------------------------
            '' 仮テーブル削除
            ''-----------------------------------------------
            'DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_dropqry, tmp_cnt)

            '
            '※修正前ソースについて…
            '・返却値を返却しない場合がある→RETURNを調整
            '・「削除」の場合にのみ仮テーブルを削除しているが、それでよい？
            '→ソースの最後の方でRETURNする
            '・上記の場合はテーブルクリアしてテーブルをドロップする？
            '→件数抽出・リスト出力の場合は仮テーブルは削除しない？
            '→全ての状態において仮テーブルを削除する(下村くん要確認)
            Select Case typeno
                Case 0  '件数抽出
                    rtn_str = Me.Get_UseQry_OwSelect(tmp_tablename, headergrp, True)
                Case 1  'リスト出力
                    rtn_str = Me.Get_UseQry_OwSelect(tmp_tablename, headergrp)
                Case 2  '削除
                    Call Me.OwDelete(tmp_tablename)
                Case Else

            End Select

            '-----------------------------------------------
            ' 仮テーブル削除
            '-----------------------------------------------
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_dropqry, tmp_cnt)

            Return rtn_str
            '20161014 改善対応：返却されない場合がある為修正 -chg end

        End Function

        ''' <summary>
        ''' 未使用家主データの抽出 '20160711 未使用データ削除_家主
        ''' </summary>
        ''' <param name="tmptblname"></param>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_OwSelect(ByVal tmptblname As String, ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "家主No,家主名,郵便番号,住所1,住所2"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [家主No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 OW.ow_no AS [家主No] "
            tmp_sql = tmp_sql & " 		,ow_name AS [家主名] "
            tmp_sql = tmp_sql & " 		,post_code AS [郵便番号] "
            tmp_sql = tmp_sql & " 		,addr1 AS [住所1] "
            tmp_sql = tmp_sql & " 		,addr2 AS [住所2] "
            tmp_sql = tmp_sql & " 	FROM owdata AS OW "
            tmp_sql = tmp_sql & " 	LEFT JOIN ow_dellst AS TMP "
            tmp_sql = tmp_sql & " 	ON OW.ow_no = TMP.ow_no  "
            tmp_sql = tmp_sql & " 	WHERE TMP.ow_no IS NULL "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr
            Return tmp_sql

        End Function

        ''' <summary>
        ''' 未使用家主データを削除 '20160711 未使用データ削除_家主
        ''' </summary>
        ''' <param name="tmptblname"></param>
        ''' <remarks></remarks>
        Private Sub OwDelete(ByVal tmptblname As String)

            Dim tmp_cnt As Integer = 0
            Dim deltablename As New List(Of String) From _
                {"owdata", "owdata_event", "owdata_gazo", "owdata_koza", "owdata_memo", "owdata_relfile"}

            For Each tblname In deltablename

                '削除クエリ作成
                Dim tmpdb_deleteqry As String = ""
                tmpdb_deleteqry = tmpdb_deleteqry & " DELETE FROM " & tblname & " FROM " & tblname & " AS MAIN "
                tmpdb_deleteqry = tmpdb_deleteqry & " LEFT JOIN " & tmptblname & " AS TMP "
                tmpdb_deleteqry = tmpdb_deleteqry & " ON MAIN.ow_no = TMP.ow_no  "
                tmpdb_deleteqry = tmpdb_deleteqry & " WHERE TMP.ow_no IS NULL "

                '削除処理
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_deleteqry, tmp_cnt)

            Next

        End Sub

        ''' <summary>
        ''' 使用されている契約者データを仮テーブルへ挿入 '20160711 未使用データ削除_契約者
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_KysInsert(ByVal tblname As String)

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " INSERT INTO " & tblname
            tmp_sql = tmp_sql & " SELECT kys_no FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	/*契約者データ*/ "
            tmp_sql = tmp_sql & " 	SELECT DISTINCT kys_no FROM kydata_kys  "
            tmp_sql = tmp_sql & " 	WHERE NOT kys_no IS NULL "
            tmp_sql = tmp_sql & " 	/*空室待ちデータ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT kys_no FROM kydata_mati  "
            tmp_sql = tmp_sql & " 	WHERE NOT kys_no IS NULL "
            tmp_sql = tmp_sql & " 	/*契約保証人データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT kys_no FROM kydata_hosyonin  "
            tmp_sql = tmp_sql & " 	WHERE NOT kys_no IS NULL "
            tmp_sql = tmp_sql & " 	/*入居者データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT kys_no FROM kydata_nyukyo  "
            tmp_sql = tmp_sql & " 	WHERE NOT kys_no IS NULL "
            tmp_sql = tmp_sql & " 	/*契約変動費データ(各戸メータータイプ)*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sqsaki_no AS kys_no FROM kydata_hendo  "
            tmp_sql = tmp_sql & " 	WHERE NOT sqsaki_no IS NULL "
            tmp_sql = tmp_sql & " 	/*契約変動費データ(親メータータイプ)*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sqsaki_no AS kys_no FROM kydata_hendoparent  "
            tmp_sql = tmp_sql & " 	WHERE NOT sqsaki_no IS NULL "
            tmp_sql = tmp_sql & " 	/*契約入金データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sqsaki_no AS kys_no FROM kydata_nkin  "
            tmp_sql = tmp_sql & " 	WHERE NOT sqsaki_no IS NULL "
            tmp_sql = tmp_sql & " 	/*次回契約入金データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sqsaki_no AS kys_no FROM kydata_nkin_nx  "
            tmp_sql = tmp_sql & " 	WHERE NOT sqsaki_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*契約解約データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT szen_kysno AS kys_no FROM kydata_kai  "
            tmp_sql = tmp_sql & " 	WHERE NOT szen_kysno IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*預り金データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT stakeholder_no AS kys_no FROM azukanri "
            tmp_sql = tmp_sql & " 	WHERE stakeholder_kbn = 100 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*出納データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT stakeholder_no AS kys_no FROM suikanri "
            tmp_sql = tmp_sql & " 	WHERE stakeholder_kbn = 100 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*変動費検針データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sqsaki_no AS kys_no FROM hendodata_meisai  "
            tmp_sql = tmp_sql & " 	WHERE NOT sqsaki_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*修繕データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT kys_no AS kys_no FROM szendata  "
            tmp_sql = tmp_sql & " 	WHERE NOT kys_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*その他請求データ(etcsq_flg=1)*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sqsaki_no AS kys_no FROM sqdata  "
            tmp_sql = tmp_sql & " 	WHERE NOT sqsaki_no IS NULL /*AND etcsq_flg = 1*/ /*部分入金も移行しているためその他請求フラグを除去*/ "
            tmp_sql = tmp_sql & " 	/*運用開始時未納・滞納金データ*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sqsaki_no AS kys_no FROM unyotainodata  "
            tmp_sql = tmp_sql & " 	WHERE NOT sqsaki_no IS NULL "
            tmp_sql = tmp_sql & " 	/*解約原状回復請求書データ		※保留*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT sq_sakino AS kys_no FROM kydata_kaiszensq  "
            tmp_sql = tmp_sql & " 	WHERE NOT sq_sakino IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*仕訳データ 念の為*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT stakeholder_no AS kys_no FROM siwakedata "
            tmp_sql = tmp_sql & "  	WHERE stakeholder_kbn = 100 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	/*口座毎勘定科目マスター 念の為*/ "
            tmp_sql = tmp_sql & " 	UNION ALL SELECT DISTINCT stakeholder_no AS kys_no FROM m_kr_kozakanjo "
            tmp_sql = tmp_sql & " 	WHERE stakeholder_type = 100 AND NOT stakeholder_no IS NULL "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " ) AS KYSTOTAL "
            tmp_sql = tmp_sql & " WHERE kys_no > 0 "
            tmp_sql = tmp_sql & " GROUP BY kys_no /*ORDER BY kys_no*/ "

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 未使用契約者データ処理(抽出/削除) '20160711 未使用データ削除_契約者
        ''' </summary>
        ''' <param name="typeno"></param>
        ''' <param name="headergrp"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_NotUseKysdata(ByVal typeno As Integer, Optional ByRef headergrp As String = "") As String

            Dim rtn_str As String = ""
            Dim tmp_tablename As String = "kys_dellst"
            Dim tmp_cnt As Integer = 0

            '-----------------------------------------------
            '使用されている契約者データ格納用仮テーブル作成
            '-----------------------------------------------
            '初期化
            Dim tmpdb_dropqry As String = DBQuery.Qry_DropInfo(tmp_tablename, True)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_dropqry, tmp_cnt)

            '仮テーブル作成
            Dim tmpdb_createqry As String = " CREATE TABLE " & tmp_tablename & " (kys_no INT NOT NULL) "
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_createqry, tmp_cnt)

            '-----------------------------------------------
            '使用されている契約者データを格納
            '-----------------------------------------------
            Dim tmpdb_insertqry As String = Me.Get_UseQry_KysInsert(tmp_tablename)
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_insertqry, tmp_cnt)

            '-----------------------------------------------
            '抽出(件数・リスト出力)/削除
            '-----------------------------------------------
            '20161014 改善対応：返却されない場合がある為修正 -chg sta
            'Select Case typeno
            '    Case 0  '件数抽出
            '        rtn_str = Me.Get_UseQry_KysSelect(tmp_tablename, headergrp, True)
            '        Return rtn_str
            '    Case 1  'リスト出力
            '        rtn_str = Me.Get_UseQry_KysSelect(tmp_tablename, headergrp)
            '        Return rtn_str
            '    Case 2  '削除
            '        Call Me.KysDelete(tmp_tablename)
            'End Select

            ''-----------------------------------------------
            ''仮テーブル削除
            ''-----------------------------------------------
            'DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_dropqry, tmp_cnt)
            Select Case typeno
                Case 0  '件数抽出
                    rtn_str = Me.Get_UseQry_KysSelect(tmp_tablename, headergrp, True)
                Case 1  'リスト出力
                    rtn_str = Me.Get_UseQry_KysSelect(tmp_tablename, headergrp)
                Case 2  '削除
                    Call Me.KysDelete(tmp_tablename)
            End Select

            '-----------------------------------------------
            '仮テーブル削除
            '-----------------------------------------------
            DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_dropqry, tmp_cnt)
            Return rtn_str
            '20161014 改善対応：返却されない場合がある為修正 -chg end

        End Function

        ''' <summary>
        ''' 未使用契約者データの抽出 '20160711 未使用データ削除_契約者
        ''' </summary>
        ''' <param name="tmptblname"></param>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_KysSelect(ByVal tmptblname As String, ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "契約者No,契約者名,郵便番号,住所1,住所2"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [契約者No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 KYS.kys_no AS [契約者No] "
            tmp_sql = tmp_sql & " 		,kys_name AS [契約者名] "
            tmp_sql = tmp_sql & " 		,post_code AS [郵便番号] "
            tmp_sql = tmp_sql & " 		,addr1 AS [住所1] "
            tmp_sql = tmp_sql & " 		,addr2 AS [住所2] "
            tmp_sql = tmp_sql & " 	FROM kysdata AS KYS "
            tmp_sql = tmp_sql & " 	LEFT JOIN kys_dellst AS TMP "
            tmp_sql = tmp_sql & " 	ON KYS.kys_no = TMP.kys_no "
            tmp_sql = tmp_sql & " 	WHERE TMP.kys_no IS NULL "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr
            Return tmp_sql

        End Function

        ''' <summary>
        ''' 未使用契約者データを削除 '20160711 未使用データ削除_契約者
        ''' </summary>
        ''' <param name="tmptblname"></param>
        ''' <remarks></remarks>
        Private Sub KysDelete(ByVal tmptblname As String)

            Dim tmp_cnt As Integer = 0
            Dim deltablename As New List(Of String) From _
                {"kysdata", "kysdata_gazo", "kysdata_koza", "kysdata_hosyo", "kysdata_memo", "kysdata_relfile"}

            For Each tblname In deltablename

                '削除クエリ作成
                Dim tmpdb_deleteqry As String = ""
                tmpdb_deleteqry = tmpdb_deleteqry & " DELETE FROM " & tblname & " FROM " & tblname & " AS MAIN "
                tmpdb_deleteqry = tmpdb_deleteqry & " LEFT JOIN " & tmptblname & " AS TMP "
                tmpdb_deleteqry = tmpdb_deleteqry & " ON MAIN.kys_no = TMP.kys_no  "
                tmpdb_deleteqry = tmpdb_deleteqry & " WHERE TMP.kys_no IS NULL "

                '削除処理
                DBExec.Exec_NonQuery(Me.sqlcnnv10, tmpdb_deleteqry, tmp_cnt)

            Next

        End Sub

        ''' <summary>
        ''' 分割入金未収分抽出クエリ '20160707 検証作業リスト出力処理の追加_分割入金未収分
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_BunkatumisyuList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,物件名,部屋No,契約No,更新No,該当年月,入金項目No,入金項目名,分割区分,請求額,請求税額,送金額,送金税額"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[部屋No],[契約No],[更新No],[該当年月],[入金項目No],[分割区分] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 SQ_BBN.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,BK.bk_name AS [物件名] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.hy_no AS [部屋No] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.ky_no AS [契約No] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.ko_no AS [更新No] "
            tmp_sql = tmp_sql & " 		/*,FORMAT(SQ_BBN.gt_ym,'yyyy/MM') AS [該当年月]*/ "
            tmp_sql = tmp_sql & " 		,LEFT(CONVERT(NVARCHAR,SQ_BBN.gt_ym,111),7) AS [該当年月] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.nkin_no AS [入金項目No] "
            tmp_sql = tmp_sql & " 		,MN.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.rec_kbn AS [分割区分] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.sq_gak AS [請求額] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.sq_zeigak AS [請求税額] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.so_gak AS [送金額] "
            tmp_sql = tmp_sql & " 		,SQ_BBN.so_zeigak AS [送金税額] "
            tmp_sql = tmp_sql & " 		/* "
            tmp_sql = tmp_sql & " 		,(sq_gak + sq_zeigak) AS [請求額] "
            tmp_sql = tmp_sql & " 		,(so_gak + so_zeigak) AS [送金額] "
            tmp_sql = tmp_sql & " 		,AAA.max_rec_kbn "
            tmp_sql = tmp_sql & " 		*/ "
            tmp_sql = tmp_sql & " 	FROM sq_meisai AS SQ_BBN "
            tmp_sql = tmp_sql & " 	LEFT JOIN bk_mst AS BK ON SQ_BBN.bk_no = BK.bk_no "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin AS MN ON SQ_BBN.nkin_no = MN.nkin_no "
            tmp_sql = tmp_sql & " 	LEFT JOIN "
            tmp_sql = tmp_sql & " 		( "
            tmp_sql = tmp_sql & " 			SELECT bk_no,hy_no,ky_no,gt_ym,nkin_no,MAX(rec_kbn) AS max_rec_kbn  "
            tmp_sql = tmp_sql & " 			FROM sq_meisai  "
            tmp_sql = tmp_sql & " 			GROUP BY bk_no,hy_no,ky_no,gt_ym,nkin_no "
            tmp_sql = tmp_sql & " 		) AS MAX_MEI  "
            tmp_sql = tmp_sql & " 	ON "
            tmp_sql = tmp_sql & " 		MAX_MEI.bk_no=SQ_BBN.bk_no "
            tmp_sql = tmp_sql & " 		AND MAX_MEI.hy_no=SQ_BBN.hy_no "
            tmp_sql = tmp_sql & " 		AND MAX_MEI.ky_no=SQ_BBN.ky_no "
            tmp_sql = tmp_sql & " 		AND MAX_MEI.gt_ym=SQ_BBN.gt_ym "
            tmp_sql = tmp_sql & " 		AND MAX_MEI.nkin_no=SQ_BBN.nkin_no "
            tmp_sql = tmp_sql & " 	/* "
            tmp_sql = tmp_sql & " 	WHERE "
            tmp_sql = tmp_sql & " 		[sq_gak] + [sq_zeigak] <> [so_gak] + [sq_zeigak] "
            tmp_sql = tmp_sql & " 		AND max_rec_kbn>0 "
            tmp_sql = tmp_sql & " 	*/ "
            tmp_sql = tmp_sql & " 	/*分割区分が0より大きいかつ未入金データを抽出*/ "
            tmp_sql = tmp_sql & " 	WHERE max_rec_kbn > 0 AND nkin_ymd IS NULL "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr
            Return tmp_sql

        End Function

        ''' <summary>
        ''' 控除支払抽出クエリ '20160707 検証作業リスト出力処理の追加_控除支払
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_KojyoshList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,物件名,該当年月,送金予定日,入金項目No,入金項目名,支払額,請求税額"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[該当年月],[送金予定日],[入金項目No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 JISYASIHARAI.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,bk_name AS [物件名] "
            tmp_sql = tmp_sql & " 		,CASE "
            tmp_sql = tmp_sql & " 			WHEN ISNULL(gt_ym,'') <> '' THEN gt_ym "
            tmp_sql = tmp_sql & " 			ELSE DATEADD(DAY,(DAY(so_ymd) - 1 ) * (-1),so_ymd) "
            tmp_sql = tmp_sql & " 		 END AS [該当年月]	 "
            tmp_sql = tmp_sql & " 		,so_ymd AS [送金予定日] "
            tmp_sql = tmp_sql & " 		,JISYASIHARAI.nkin_no AS [入金項目No] "
            tmp_sql = tmp_sql & " 		,MN.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 		,siharai_gak AS [支払額] "
            tmp_sql = tmp_sql & " 		,siharai_zeigak AS [請求税額] "
            tmp_sql = tmp_sql & " 	FROM "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		SELECT * FROM sh_kojo "
            tmp_sql = tmp_sql & " 		WHERE siharai_gak <> 0 "
            tmp_sql = tmp_sql & " 		AND   nkin_no <> 9910 "
            tmp_sql = tmp_sql & " 	) AS JISYASIHARAI "
            tmp_sql = tmp_sql & " 	LEFT JOIN bk_mst AS BK ON JISYASIHARAI.bk_no = BK.bk_no "
            tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin AS MN ON JISYASIHARAI.nkin_no = MN.nkin_no "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 修繕項目毎契約者送金率抽出クエリ '20160707 検証作業リスト出力処理の追加_修繕項目毎契約者送金率
        ''' </summary>
        ''' <param name="headergrp"></param>
        ''' <param name="cntflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_UseQry_SzenKysSoritList(ByRef headergrp As String, Optional ByVal cntflg As Boolean = False) As String

            headergrp = "物件No,部屋番号,リフォームID,行No,入金項目No,入金項目名,送金率"
            Dim selecttaisyo As String = IIf(cntflg, "COUNT(*)", "*")
            Dim orderstr As String = IIf(cntflg, "", " ORDER BY [物件No],[部屋番号],[リフォームID],[行No],[入金項目No] ")

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT " & selecttaisyo & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT  "
            tmp_sql = tmp_sql & " 		 MAIN.bk_no AS [物件No] "
            tmp_sql = tmp_sql & " 		,MAIN.hy_composite AS [部屋番号] "
            tmp_sql = tmp_sql & " 		,MAIN.reform_id AS [リフォームID] "
            tmp_sql = tmp_sql & " 		,MAIN.item_id + 1 AS [行No] "
            tmp_sql = tmp_sql & " 		,MAIN.nkin_no AS [入金項目No] "
            tmp_sql = tmp_sql & " 		,MN.nkin_name AS [入金項目名] "
            tmp_sql = tmp_sql & " 		,MAIN.kari_so_rit AS [送金率] "
            tmp_sql = tmp_sql & " 	FROM reform_item MAIN "
            tmp_sql = tmp_sql & " 	LEFT JOIN M_NKIN AS MN ON MAIN.nkin_no = MN.nkin_no "
            tmp_sql = tmp_sql & " 	WHERE EXISTS "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		SELECT  "
            tmp_sql = tmp_sql & " 			 RI.kari_so_rit  "
            tmp_sql = tmp_sql & " 			,RI.bk_no "
            tmp_sql = tmp_sql & " 			,RI.hy_composite "
            tmp_sql = tmp_sql & " 			,RI.reform_id "
            tmp_sql = tmp_sql & " 			,RI.kari_so_rit  "
            tmp_sql = tmp_sql & " 			,MAXRIT.max_rit "
            tmp_sql = tmp_sql & " 		FROM reform_item RI "
            tmp_sql = tmp_sql & " 		LEFT JOIN "
            tmp_sql = tmp_sql & " 			( "
            tmp_sql = tmp_sql & " 				SELECT  "
            tmp_sql = tmp_sql & " 					 MAX(kari_so_rit) AS max_rit "
            tmp_sql = tmp_sql & " 					,bk_no "
            tmp_sql = tmp_sql & " 					,hy_composite "
            tmp_sql = tmp_sql & " 					,reform_id "
            tmp_sql = tmp_sql & " 				FROM reform_item "
            tmp_sql = tmp_sql & " 				WHERE NOT kari_so_rit IS NULL "
            tmp_sql = tmp_sql & " 				GROUP BY bk_no,hy_composite,reform_id "
            tmp_sql = tmp_sql & " 			) MAXRIT "
            tmp_sql = tmp_sql & " 		ON MAXRIT.bk_no=RI.bk_no AND MAXRIT.hy_composite=RI.hy_composite AND MAXRIT.reform_id=RI.reform_id "
            tmp_sql = tmp_sql & " 		WHERE  "
            tmp_sql = tmp_sql & " 			NOT RI.kari_so_rit IS NULL  "
            tmp_sql = tmp_sql & " 			AND RI.kari_so_rit<>MAXRIT.max_rit "
            tmp_sql = tmp_sql & " 			AND MAIN.bk_no=RI.bk_no AND MAIN.hy_composite=RI.hy_composite AND MAIN.reform_id=RI.reform_id "
            tmp_sql = tmp_sql & " 	) "
            tmp_sql = tmp_sql & " AND MAIN.nkin_no <> 8999 "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & orderstr

            Return tmp_sql

        End Function

#End Region

#Region "コンバート実績を元にした処理"   '20160707 コンバート実績保持の処理追加 -add

        ''' <summary>
        ''' コンバート実績をXMLファイルへ保持しておく処理 '20160707 コンバート実績保持の処理追加 -add
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub MakeFile_CVJisseki()

            '移行項目取得
            Dim list_cvitem As New List(Of String)
            list_cvitem = Me.Get_ListCVChkitemAll(True)

            '※コンバート後なので0件は無いが念の為
            If list_cvitem.Count = 0 Then
                Exit Sub
            End If

            'コンバート実績保存ファイル格納先取得
            Dim cvjissekiinfodirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_INI_NAME)
            Dim cvjissekifilepath As String = ""
            cvjissekifilepath = EtcMethod.Set_Path(cvjissekiinfodirpath, FILE_CVJISSEKINAME)

            'xml作成
            '最初の固定文字列
            Dim strxml_pre As String = "<?xml version='1.0'?>" & _
                                         "<cvjisseki>" & _
                                            "<!--コンバート実績-->"

            '最後の固定文字列
            Dim strxml_post As String = "</cvjisseki>"

            '移行項目の文字列
            Dim strxml_main As String = ""
            Dim tmp_str As String = ""
            Dim cntii As Integer = 1
            For Each cvitem In list_cvitem
                tmp_str = tmp_str & "<cvitem" & cntii.ToString & ">" & cvitem & "</cvitem" & cntii.ToString & ">"
                cntii = cntii + 1
            Next
            strxml_main = tmp_str

            '統合
            Dim strxml_total As String = strxml_pre & strxml_main & strxml_post

            Dim xmlDoc As New System.Xml.XmlDocument

            '文字列からDOMドキュメントを生成
            xmlDoc.LoadXml(strxml_total)
            '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
            'Try
            '    '作成したDOMドキュメントをファイルに保存
            '    xmlDoc.Save(cvjissekifilepath)
            'Catch ex As System.Xml.XmlException
            '    'XMLによる例外をキャッチ
            '    Console.WriteLine(ex.Message)
            'Catch ex As Exception
            '    'その他の例外をキャッチ
            '    Console.WriteLine(ex.Message)
            'End Try
            Try
                '作成したDOMドキュメントをファイルに保存
                xmlDoc.Save(cvjissekifilepath)

                'アクセス拒否
                'Catch ex As UnauthorizedAccessException
                '    MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                '                                "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
                '                                "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '    Return False
                'その他の例外(アクセス拒否を含む)
            Catch ex As Exception               '20161108 レビュー結果：「セットアップフォルダ内～拒否されたため、コンバート実績を～できませんでした。」のように修正
                '20161108_2 レビュー結果戻り修正 -chg sta
                'MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                '                            "コンバート実績を保持することができませんでした。" & vbCrLf & _
                '                            "コンバート実績の保持を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" & _
                '                            "再度セットアップを行ってから、本プログラムを実行して下さい。" & vbCrLf & _
                '                            "※コンバート履歴の保持のみであるためコンバートデータに影響はありません。", _
                '                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されため、" & vbCrLf & _
                                            "コンバート実績を保持することができませんでした。" & vbCrLf & _
                                            "コンバート実績の保持を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" & _
                                            "再度セットアップを行ってから、本プログラムを実行して下さい。" & vbCrLf & _
                                            "※コンバートデータへの影響は無いためこのまま作業を続行しても問題はありません。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '20161108_2 レビュー結果戻り修正 -chg end
            End Try                             '20161108 レビュー結果：「※コンバート履歴～」は除去またはもう少しわかり易い内容で。
            '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end
        End Sub

        ''' <summary>
        ''' コンバート実績の削除(実績を残したxmlファイルを削除) '20160707 コンバート実績保持の処理追加 -add
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub DeleteFile_CVJisseki()

            '保管されているコンバート実績ファイルの読込
            Dim cvjissekiinfodirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_INI_NAME)
            Dim cvjissekifilepath As String = ""
            cvjissekifilepath = EtcMethod.Set_Path(cvjissekiinfodirpath, FILE_CVJISSEKINAME)

            'ファイルが存在しない(コンバート実績が存在しないまたはファイル名が編集されている)場合は処理を抜ける
            If EtcMethod.Chk_FileExist(cvjissekifilepath) = False Then
                Exit Sub
            End If

            'ファイル削除
            '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
            'File.Delete(cvjissekifilepath)
            Try
                '作成したDOMドキュメントをファイルに保存
                File.Delete(cvjissekifilepath)

                'アクセス拒否
                'Catch ex As UnauthorizedAccessException
                '    MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                '                                "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
                '                                "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '    Return False
                'その他の例外(アクセス拒否を含む)
            Catch ex As Exception           '20161108 レビュー結果：MakeFile_CVJissekiと同様
                '20161108_2 レビュー結果戻り修正 -chg sta
                'MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                '                            "コンバート実績を削除することができませんでした。" & vbCrLf & _
                '                            "コンバート実績の削除を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" & _
                '                            "再度セットアップを行ってから、本プログラムを実行して下さい。" & vbCrLf & _
                '                            "※コンバート履歴の削除のみであるためコンバートデータに影響はありません。", _
                '                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されたため、" & vbCrLf & _
                                            "コンバート実績を削除することができませんでした。" & vbCrLf & _
                                            "コンバート実績の削除を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" & _
                                            "再度セットアップを行ってから、本プログラムを実行して下さい。" & vbCrLf & _
                                            "※コンバートデータへの影響は無いためこのまま作業を続行しても問題はありません。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '20161108_2 レビュー結果戻り修正 -chg end
            End Try
            '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end

        End Sub

        ''' <summary>
        ''' コンバート実績読込処理 '20160707 コンバート実績保持の処理追加
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_CVJisseki()

            'ボタンの活性制御
            btnRekiClear.Enabled = True                                 '20160711 コンバート履歴処理の追加 -add
        
            '項目名とチェックボックスを紐付けたハッシュテーブルを生成
            Dim hash_cvitemtochkbox As New Hashtable
            hash_cvitemtochkbox = Get_Hash_CVChkitemAll()

            'チェックボックスのチェックON、着色初期化
            For Each chkbox In hash_cvitemtochkbox
                Dim tmp_chkbox As New CheckBox
                tmp_chkbox = chkbox.Value
                tmp_chkbox.ForeColor = Color.Black
                '20161014 着色処理の修正 -chg sta
                'tmp_chkbox.Checked = True
                Select Case CNVNO
                    Case ConvertTypes._既存ユーザ用
                        tmp_chkbox.Checked = True
                    Case ConvertTypes._汎用
                        tmp_chkbox.Checked = False
                End Select
                '20161014 着色処理の修正 -chg end
            Next

            '保管されているコンバート実績ファイルの読込
            Dim cvjissekiinfodirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_INI_NAME)
            Dim cvjissekifilepath As String = ""
            cvjissekifilepath = EtcMethod.Set_Path(cvjissekiinfodirpath, FILE_CVJISSEKINAME)

            'ファイルが存在しない(コンバート実績が存在しないまたはファイル名が編集されている)場合は処理を抜ける
            If EtcMethod.Chk_FileExist(cvjissekifilepath) = False Then
                'ボタンの活性制御
                btnRekiClear.Enabled = False                            '20160711 コンバート履歴処理の追加 -add
                Exit Sub
            End If

            'xmlファイル読込処理
            Dim xmlreader As System.Xml.XmlReader
            Dim item As String = ""
            Dim data As String = ""
            Dim itemprename As String = "cvitem"
            Dim list_cvjissekiitem As New List(Of String)
            Dim itemcnt As Integer = 1
            xmlreader = System.Xml.XmlReader.Create(cvjissekifilepath)

            While xmlreader.Read
                If xmlreader.NodeType = Xml.XmlNodeType.Element Then
                    'データ取得
                    item = xmlreader.LocalName
                    data = xmlreader.ReadString

                    If item = itemprename & itemcnt.ToString Then
                        list_cvjissekiitem.Add(data)
                        itemcnt = itemcnt + 1
                    End If
                End If
            End While

            xmlreader.Close()

            If list_cvjissekiitem.Count = 0 Then                        '※コンバート実績が残っている段階で0は無いが念の為
                'ボタンの活性制御
                btnRekiClear.Enabled = False                            '20160711 コンバート履歴処理の追加 -add
                Exit Sub
            End If

            'コンバート実績に残っている項目と照合し、一致したチェックボックスのチェックOFFと着色
            For Each cvitem In list_cvjissekiitem
                Dim chkbox As New CheckBox
                If hash_cvitemtochkbox.Contains(cvitem) Then
                    chkbox = hash_cvitemtochkbox(cvitem)
                    chkbox.Checked = False
                    chkbox.ForeColor = Color.Green
                End If
            Next

        End Sub

#End Region

    End Class

End Namespace