
namespace Converter10.Njc.Model
{

    #region 接続情報

    public class DefSQLConnection
    {

        public string ServerName { get; set; }                                    // 接続サーバー名
        public string InitialCatalog { get; set; }                                // カタログ名
        public int NetworkLibrary { get; set; }                               // 接続で使用するネットワークライブラリの種類
        public string User { get; set; }                                          // 認証で使用するユーザ名
        public string Pass { get; set; }                                          // 認証で使用するパスワード
        public string TimeOut { get; set; }                                       // DB接続時タイムアウト設定値(SEC)

        // 2016.03.28 接続情報取得処理修正 -del sta
        // Public Sub New(ByVal flg As Integer)

        // '開発用に接続情報を設定しておく(正式版では空白にすること)          'kakaka 注意：忘れずに空白にすること
        // If flg = 0 Then
        // 'V7用固定設定
        // Me.ServerName = "PC-1KA_SOL\SQL2K8"
        // Me.InitialCatalog = "fk5dtsql"
        // Me.User = "sa"
        // Me.Pass = "p7s2#c1j3n"
        // Else
        // 'V10用固定設定
        // Me.ServerName = "PC-1KA_SOL\SQL2012"
        // Me.InitialCatalog = "fk8db"
        // 'Me.InitialCatalog = "fk8db_cv_stest"
        // Me.User = "sa"
        // Me.Pass = "p7s2#c1j3n"
        // End If

        // Me.TimeOut = "30"

        // End Sub
        // 2016.03.28 接続情報取得処理修正 -del end

        // 2016.02.22 V7DB接続情報取得処理追加 -add sta
        public DefSQLConnection()
        {

            ServerName = "";
            InitialCatalog = "";
            User = "";
            Pass = "";
            TimeOut = "40";                   // 20160514 kakaka test chg (30->15))     '20160829 連動中間ファイル作成処理修正 15→40

        }
        // 2016.02.22 V7DB接続情報取得処理追加 -add end

    }

    #endregion

    #region 移行項目グループ情報

    public class CVItemGrpInfo
    {

        public string MstGrpInfo { get; set; }        // マスタ系
        public string GyGrpInfo { get; set; }         // 業者情報
        public string JisyaGrpInfo { get; set; }      // 自社情報
        public string OwGrpInfo { get; set; }         // 家主情報
        public string KysGrpInfo { get; set; }        // 契約者情報
        public string BkGrpInfo { get; set; }         // 物件情報
        public string HyGrpInfo { get; set; }         // 部屋情報

    }

    #endregion

    #region 移行項目情報

    public class CVItemInfo
    {

        // ----------
        // マスタ系
        // ----------
        public string MstBkbrui { get; set; }
        public string MstBusKotu { get; set; }
        public string MstBus { get; set; }
        public string MstHyrui { get; set; }
        public string MstKagititle { get; set; }
        public string MstKeikaikakunin { get; set; }
        public string MstOwevent { get; set; }
        public string MstKasyoClaimrui { get; set; }
        public string MstTokuyaku { get; set; }
        public string MstYane { get; set; }
        public string MstKeiyakurui { get; set; }
        public string MstHokenrui { get; set; }
        public string MstNkinkbn { get; set; }
        public string MstKinyu { get; set; }
        public string MstKinyuten { get; set; }
        public string MstSchool { get; set; }
        public string MstArea { get; set; }
        public string MstKozasyubetu { get; set; }
        public string MstKozo { get; set; }
        public string MstTorihikitaiyo { get; set; }
        public string MstKeiyakusyubetu { get; set; }
        public string MstNkinkomok { get; set; }
        public string MstSetubi { get; set; }
        public string MstHendo { get; set; }
        public string MstHendoitiran { get; set; }
        public string MstBikotitle { get; set; }      // 20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add
        public string MstBikolst { get; set; }        // 20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add
        public string MstGazotitle { get; set; }      // 20160801 タイトルマスタ統合処理 -add

        // ----------
        // 業者情報
        // ----------
        public string GyCyukaiBase { get; set; }
        public string GyCyukaiKoza { get; set; }
        public string GyCyukaiMemo { get; set; }
        public string GyHokenBase { get; set; }
        public string GyHokenKoza { get; set; }
        public string GyHokenMemo { get; set; }
        public string GyYatinhosyoBase { get; set; }
        public string GyYatinhosyoKoza { get; set; }
        public string GyYatinhosyoMemo { get; set; }
        public string GySyuzenBase { get; set; }
        public string GySyuzenKoza { get; set; }
        public string GySyuzenMemo { get; set; }
        public string GyLifelineBase { get; set; }
        public string GySekoBase { get; set; }
        public string GySisetuBase { get; set; }

        // ----------
        // 自社情報
        // ----------
        public string JisyaBase { get; set; }
        public string JisyaKoza { get; set; }
        public string JisyaTanto { get; set; }        // 20160608 自社担当者情報構築 -add
        public string JisyaMemo { get; set; }
        public string FBFuriirai { get; set; }        // FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        // 20160517 振込手数料情報の新規作成 -add
        public string FBFuritesuryo { get; set; }     // FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        public string FBKozafurikae { get; set; }     // FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        // 20160517 入出金取得情報の新規作成 -add
        public string FBNsSyutoku { get; set; }       // FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        // 20160517 家賃入金口座情報の新規作成 -add
        public string MstYatinKoza { get; set; }      // マスタ系の移行項目だが自社口座情報が関連するためとりあえずこの分類にしておく
        public string MstANSERAccpoint { get; set; }  // 20160704 ANSER情報の移行処理追加 -add
        public string MstANSERArea { get; set; }      // 20160704 ANSER情報の移行処理追加 -add
        public string FBANSERSetuzoku { get; set; }   // 20160704 ANSER情報の移行処理追加 -add

        // ----------
        // 家主情報
        // ----------
        public string OwBase { get; set; }
        public string OwKoza { get; set; }
        public string OwEvent { get; set; }
        public string OwMemo { get; set; }

        // ----------
        // 契約者情報
        // ----------
        public string KysBase { get; set; }
        public string KysKoza { get; set; }
        public string KysMemo { get; set; }
        public string KysSyogoKana { get; set; }
        public string KysHosyonin { get; set; }

        // ----------
        // 物件情報
        // ----------
        public string BkBase { get; set; }
        public string BkSyosai { get; set; }
        public string Bksyo { get; set; }
        public string BkGomi { get; set; }
        public string BkKenri { get; set; }
        public string BkKotu { get; set; }
        public string BkSetudo { get; set; }
        public string BkSyuhen { get; set; }
        public string BkSzeniji { get; set; }
        public string BkMemo { get; set; }
        public string BkKagi { get; set; }                // 汎用分(V7には無い)
        public string BkHendo { get; set; }               // 汎用分(V7には無い)
        public string BkKinrincyusyajo { get; set; }      // 汎用分(V7には無い)
        public string BkSansyofile { get; set; }          // 汎用分(V7には無い)

        // ----------
        // 部屋情報
        // ----------
        public string HyBase { get; set; }
        public string HySyosai { get; set; }
        public string Hysyo { get; set; }
        public string HyParking { get; set; }
        public string HyTokuyaku { get; set; }
        public string HyKagi { get; set; }
        public string HyMadoriutiwake { get; set; }

        public string HyMenseki { get; set; }
        public string HySetubi { get; set; }
        public string HyNkinkomk { get; set; }
        public string HyHendo { get; set; }
        public string HyMemo { get; set; }

        public string HyCommonsalespoint { get; set; }
        public string HyConfirm { get; set; }
        public string HyKenri { get; set; }
        public string HySansyofile { get; set; }
        public string HyGenjotanka { get; set; }
        public string HySzeniji { get; set; }

        // ----------
        // 送金ルール
        // ----------
        public string SoruleBase { get; set; }
        public string SoruleSosaki { get; set; }
        public string SoruleNkin { get; set; }
        public string SoruleKojo { get; set; }

        // ----------
        // 契約情報
        // ----------
        public string KyBase { get; set; }
        public string KyRireki { get; set; }
        public string KyCar { get; set; }
        public string KyKys { get; set; }
        public string KyHosyonin { get; set; }
        public string KyNyukyo { get; set; }
        public string KyTokuyaku { get; set; }
        public string KyMemo { get; set; }
        public string KyHoken { get; set; }
        public string KyNkinkomk { get; set; }
        public string KyNkinkomkNx { get; set; }
        public string KyHendo { get; set; }
        public string KyKojoRule { get; set; }
        public string KySorule { get; set; }
        public string KyKai { get; set; }
        public string KyKagi { get; set; }            // 2016.04.06 契約鍵情報の移行処理追加 -add
        public string KySzen { get; set; }            // 20160621 修繕関連移行処理追加 -add
        public string KySzenmeisai { get; set; }      // 20160621 修繕関連移行処理追加 -add

        // ----------
        // 請求情報
        // ----------
        public string SqKajyo { get; set; }
        public string SqUnyotaino { get; set; }
        public string SqSq { get; set; }
        public string SqHendokensin { get; set; }
        public string SqKoteiKojo { get; set; }
        public string SqSqKojo { get; set; }

        // ----------
        // クレーム修繕情報
        // ----------
        public string ClaimBase { get; set; }
        public string ClaimTaiorireki { get; set; }
        public string ClaimRelfile { get; set; }
        public string SzenBase { get; set; }          // 20160621 修繕関連移行処理追加 -add
        public string SzenSzen { get; set; }          // 20160621 修繕関連移行処理追加 -add
        public string SzenSzenmeisai { get; set; }    // 20160621 修繕関連移行処理追加 -add
        public string SzenClaim { get; set; }         // 20160621 修繕関連移行処理追加 -add
        public string SzenRelfile { get; set; }       // 20160627 修繕関連ファイル移行修正 -add
        public string SzenMemo { get; set; }          // 20160621 修繕関連移行処理追加 -add

        // ----------
        // 初期設定情報
        // ----------
        public string SyskanriBase { get; set; }
        public string SyskanriZei { get; set; }
        public string SyskanriHenkanmoji { get; set; }
        public string SyskanriNkinkomkmerge { get; set; }

        // ----------
        // 物件データ連動情報 '20160720 連動情報構築 -add
        // ----------
        public string RendoSosinBase { get; set; }
        public string RendoSosinJisyaweb { get; set; }
        public string RendoSosinHomes { get; set; }
        public string RendoSosinAthome { get; set; }
        public string RendoSosinSuumo { get; set; }
        public string RendoMapdisp { get; set; }
        public string RendoBtoBgroup { get; set; }
        public string RendoHysosin { get; set; }
        public string RendoHyrui { get; set; }
        public string RendoKokokuSuumo { get; set; }
        public string RendoKokokuAthome { get; set; }
        public string RendoKokokuHomes { get; set; }
        public string RendoKokokuJisyaweb { get; set; }

        public CVItemInfo()
        {

        }

    }

    #endregion

    #region ボタン状態情報


    // 20160705 kakaka del -sta：不要
    // Public Class BtnBackInfo

    // Public Property visibleinfo As Boolean
    // Public Property enableinfo As Boolean
    // Public Property textinfo As String

    // End Class

    // Public Class BtnNextInfo

    // Public Property visibleinfo As Boolean
    // Public Property enableinfo As Boolean
    // Public Property textinfo As String

    // End Class

    // Public Class BtnEndInfo

    // Public Property visibleinfo As Boolean
    // Public Property enableinfo As Boolean
    // Public Property textinfo As String

    // End Class
    // 20160705 kakaka del -end

    #endregion

}