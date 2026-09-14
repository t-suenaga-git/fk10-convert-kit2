Namespace Njc.Model

#Region "接続情報"

    Public Class DefSQLConnection

        Public Property ServerName As String                                    '接続サーバー名
        Public Property InitialCatalog As String                                'カタログ名
        Public Property NetworkLibrary As Integer                               '接続で使用するネットワークライブラリの種類
        Public Property User As String                                          '認証で使用するユーザ名
        Public Property Pass As String                                          '認証で使用するパスワード
        Public Property TimeOut As String                                       'DB接続時タイムアウト設定値(SEC)

        '2016.03.28 接続情報取得処理修正 -del sta
        'Public Sub New(ByVal flg As Integer)

        '    '開発用に接続情報を設定しておく(正式版では空白にすること)          'kakaka 注意：忘れずに空白にすること
        '    If flg = 0 Then
        '        'V7用固定設定
        '        Me.ServerName = "PC-1KA_SOL\SQL2K8"
        '        Me.InitialCatalog = "fk5dtsql"
        '        Me.User = "sa"
        '        Me.Pass = "p7s2#c1j3n"
        '    Else
        '        'V10用固定設定
        '        Me.ServerName = "PC-1KA_SOL\SQL2012"
        '        Me.InitialCatalog = "fk8db"
        '        'Me.InitialCatalog = "fk8db_cv_stest"
        '        Me.User = "sa"
        '        Me.Pass = "p7s2#c1j3n"
        '    End If

        '    Me.TimeOut = "30"

        'End Sub
        '2016.03.28 接続情報取得処理修正 -del end

        '2016.02.22 V7DB接続情報取得処理追加 -add sta
        Public Sub New()

            Me.ServerName = ""
            Me.InitialCatalog = ""
            Me.User = ""
            Me.Pass = ""
            Me.TimeOut = "40"                   '20160514 kakaka test chg (30->15))     '20160829 連動中間ファイル作成処理修正 15→40

        End Sub
        '2016.02.22 V7DB接続情報取得処理追加 -add end

    End Class

#End Region

#Region "移行項目グループ情報"

    Public Class CVItemGrpInfo

        Public Property MstGrpInfo As String        'マスタ系
        Public Property GyGrpInfo As String         '業者情報
        Public Property JisyaGrpInfo As String      '自社情報
        Public Property OwGrpInfo As String         '家主情報
        Public Property KysGrpInfo As String        '契約者情報
        Public Property BkGrpInfo As String         '物件情報
        Public Property HyGrpInfo As String         '部屋情報

    End Class

#End Region

#Region "移行項目情報"

    Public Class CVItemInfo

        '----------
        'マスタ系
        '----------
        Public Property MstBkbrui As String
        Public Property MstBusKotu As String
        Public Property MstBus As String
        Public Property MstHyrui As String
        Public Property MstKagititle As String
        Public Property MstKeikaikakunin As String
        Public Property MstOwevent As String
        Public Property MstKasyoClaimrui As String
        Public Property MstTokuyaku As String
        Public Property MstYane As String
        Public Property MstKeiyakurui As String
        Public Property MstHokenrui As String
        Public Property MstNkinkbn As String
        Public Property MstKinyu As String
        Public Property MstKinyuten As String
        Public Property MstSchool As String
        Public Property MstArea As String
        Public Property MstKozasyubetu As String
        Public Property MstKozo As String
        Public Property MstTorihikitaiyo As String
        Public Property MstKeiyakusyubetu As String
        Public Property MstNkinkomok As String
        Public Property MstSetubi As String
        Public Property MstHendo As String
        Public Property MstHendoitiran As String
        Public Property MstBikotitle As String      '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add
        Public Property MstBikolst As String        '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add
        Public Property MstGazotitle As String      '20160801 タイトルマスタ統合処理 -add

        '----------
        '業者情報
        '----------
        Public Property GyCyukaiBase As String
        Public Property GyCyukaiKoza As String
        Public Property GyCyukaiMemo As String
        Public Property GyHokenBase As String
        Public Property GyHokenKoza As String
        Public Property GyHokenMemo As String
        Public Property GyYatinhosyoBase As String
        Public Property GyYatinhosyoKoza As String
        Public Property GyYatinhosyoMemo As String
        Public Property GySyuzenBase As String
        Public Property GySyuzenKoza As String
        Public Property GySyuzenMemo As String
        Public Property GyLifelineBase As String
        Public Property GySekoBase As String
        Public Property GySisetuBase As String

        '----------
        '自社情報
        '----------
        Public Property JisyaBase As String
        Public Property JisyaKoza As String
        Public Property JisyaTanto As String        '20160608 自社担当者情報構築 -add
        Public Property JisyaMemo As String
        Public Property FBFuriirai As String        'FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        '20160517 振込手数料情報の新規作成 -add
        Public Property FBFuritesuryo As String     'FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        Public Property FBKozafurikae As String     'FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        '20160517 入出金取得情報の新規作成 -add
        Public Property FBNsSyutoku As String       'FBの移行項目だが自社口座情報が必要なためとりあえずこの分類にしておく
        '20160517 家賃入金口座情報の新規作成 -add
        Public Property MstYatinKoza As String      'マスタ系の移行項目だが自社口座情報が関連するためとりあえずこの分類にしておく
        Public Property MstANSERAccpoint As String  '20160704 ANSER情報の移行処理追加 -add
        Public Property MstANSERArea As String      '20160704 ANSER情報の移行処理追加 -add
        Public Property FBANSERSetuzoku As String   '20160704 ANSER情報の移行処理追加 -add

        '----------
        '家主情報
        '----------
        Public Property OwBase As String
        Public Property OwKoza As String
        Public Property OwEvent As String
        Public Property OwMemo As String

        '----------
        '契約者情報
        '----------
        Public Property KysBase As String
        Public Property KysKoza As String
        Public Property KysMemo As String
        Public Property KysSyogoKana As String
        Public Property KysHosyonin As String

        '----------
        '物件情報
        '----------
        Public Property BkBase As String
        Public Property BkSyosai As String
        Public Property Bksyo As String
        Public Property BkGomi As String
        Public Property BkKenri As String
        Public Property BkKotu As String
        Public Property BkSetudo As String
        Public Property BkSyuhen As String
        Public Property BkSzeniji As String
        Public Property BkMemo As String
        Public Property BkKagi As String                '汎用分(V7には無い)
        Public Property BkHendo As String               '汎用分(V7には無い)
        Public Property BkKinrincyusyajo As String      '汎用分(V7には無い)
        Public Property BkSansyofile As String          '汎用分(V7には無い)

        '----------
        '部屋情報
        '----------
        Public Property HyBase As String
        Public Property HySyosai As String
        Public Property Hysyo As String
        Public Property HyParking As String
        Public Property HyTokuyaku As String
        Public Property HyKagi As String
        Public Property HyMadoriutiwake As String

        Public Property HyMenseki As String
        Public Property HySetubi As String
        Public Property HyNkinkomk As String
        Public Property HyHendo As String
        Public Property HyMemo As String

        Public Property HyCommonsalespoint As String
        Public Property HyConfirm As String
        Public Property HyKenri As String
        Public Property HySansyofile As String
        Public Property HyGenjotanka As String
        Public Property HySzeniji As String

        '----------
        '送金ルール
        '----------
        Public Property SoruleBase As String
        Public Property SoruleSosaki As String
        Public Property SoruleNkin As String
        Public Property SoruleKojo As String

        '----------
        '契約情報
        '----------
        Public Property KyBase As String
        Public Property KyRireki As String
        Public Property KyCar As String
        Public Property KyKys As String
        Public Property KyHosyonin As String
        Public Property KyNyukyo As String
        Public Property KyTokuyaku As String
        Public Property KyMemo As String
        Public Property KyHoken As String
        Public Property KyNkinkomk As String
        Public Property KyNkinkomkNx As String
        Public Property KyHendo As String
        Public Property KyKojoRule As String
        Public Property KySorule As String
        Public Property KyKai As String
        Public Property KyKagi As String            '2016.04.06 契約鍵情報の移行処理追加 -add
        Public Property KySzen As String            '20160621 修繕関連移行処理追加 -add
        Public Property KySzenmeisai As String      '20160621 修繕関連移行処理追加 -add

        '----------
        '請求情報
        '----------
        Public Property SqKajyo As String
        Public Property SqUnyotaino As String
        Public Property SqSq As String
        Public Property SqHendokensin As String
        Public Property SqKoteiKojo As String
        Public Property SqSqKojo As String

        '----------
        'クレーム修繕情報
        '----------
        Public Property ClaimBase As String
        Public Property ClaimTaiorireki As String
        Public Property ClaimRelfile As String
        Public Property SzenBase As String          '20160621 修繕関連移行処理追加 -add
        Public Property SzenSzen As String          '20160621 修繕関連移行処理追加 -add
        Public Property SzenSzenmeisai As String    '20160621 修繕関連移行処理追加 -add
        Public Property SzenClaim As String         '20160621 修繕関連移行処理追加 -add
        Public Property SzenRelfile As String       '20160627 修繕関連ファイル移行修正 -add
        Public Property SzenMemo As String          '20160621 修繕関連移行処理追加 -add

        '----------
        '初期設定情報
        '----------
        Public Property SyskanriBase As String
        Public Property SyskanriZei As String
        Public Property SyskanriHenkanmoji As String
        Public Property SyskanriNkinkomkmerge As String

        '----------
        '物件データ連動情報 '20160720 連動情報構築 -add
        '----------
        Public Property RendoSosinBase As String
        Public Property RendoSosinJisyaweb As String
        Public Property RendoSosinHomes As String
        Public Property RendoSosinAthome As String
        Public Property RendoSosinSuumo As String
        Public Property RendoMapdisp As String
        Public Property RendoBtoBgroup As String
        Public Property RendoHysosin As String
        Public Property RendoHyrui As String
        Public Property RendoKokokuSuumo As String
        Public Property RendoKokokuAthome As String
        Public Property RendoKokokuHomes As String
        Public Property RendoKokokuJisyaweb As String

        Public Sub New()

        End Sub

    End Class

#End Region

#Region "ボタン状態情報"


    '20160705 kakaka del -sta：不要
    'Public Class BtnBackInfo

    '    Public Property visibleinfo As Boolean
    '    Public Property enableinfo As Boolean
    '    Public Property textinfo As String

    'End Class

    'Public Class BtnNextInfo

    '    Public Property visibleinfo As Boolean
    '    Public Property enableinfo As Boolean
    '    Public Property textinfo As String

    'End Class

    'Public Class BtnEndInfo

    '    Public Property visibleinfo As Boolean
    '    Public Property enableinfo As Boolean
    '    Public Property textinfo As String

    'End Class
    '20160705 kakaka del -end

#End Region

End Namespace