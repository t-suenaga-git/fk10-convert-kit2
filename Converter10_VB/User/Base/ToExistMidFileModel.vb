Namespace Njc.Model

#Region "部屋駐車場モデル"

    Public Class Hy_cyusyajo_mid_Model

        Public Property ObjCol_物件NO As New Object
        Public Property ObjCol_部屋NO As New Object
        Public Property ObjCol_駐車場の空き数手動設定用 As New Object
        Public Property ObjCol_駐車場の空き有無手動設定用 As New Object
        Public Property ObjCol_駐車場料金区分手動設定用 As New Object
        Public Property ObjCol_駐車場料金手動設定用 As New Object
        Public Property ObjCol_駐車場料金税区分手動設定用 As New Object
        Public Property ObjCol_駐車場の賃貸可能数 As New Object
        Public Property ObjCol_バイクの空き数手動設定用 As New Object
        Public Property ObjCol_バイク駐車場の空き有無手動設定用 As New Object
        Public Property ObjCol_バイク駐車場料金区分手動設定用 As New Object
        Public Property ObjCol_バイク駐車場料金手動設定用 As New Object
        Public Property ObjCol_バイク駐車場料金税区分手動設定用 As New Object
        Public Property ObjCol_バイク駐車場の賃貸可能数 As New Object
        Public Property ObjCol_駐輪場駐車場の空き数手動設定用 As New Object
        Public Property ObjCol_駐輪場駐車場の空き有無手動設定用 As New Object
        Public Property ObjCol_駐輪場駐車場料金区分手動設定用 As New Object
        Public Property ObjCol_駐輪場駐車場料金手動設定用 As New Object
        Public Property ObjCol_駐輪場駐車場料金税区分手動設定用 As New Object
        Public Property ObjCol_駐輪場駐車場の賃貸可能数 As New Object

        Public Property ItemCnt As Integer

    End Class

#End Region

#Region "部屋特約モデル"

    Public Class Hy_tokuyaku_mid_Model

        Public Property ObjCol_物件NO As New Object
        Public Property ObjCol_部屋NO As New Object
        Public Property ObjCol_原状回復特約内容 As New Object
        Public Property ObjCol_入居中修繕特約内容 As New Object
        Public Property ObjCol_その他特約内容 As New Object
        Public Property ItemCnt As Integer

    End Class

#End Region

#Region "部屋契約解約確認事項モデル"

    Public Class Hy_kykaikakuninjiko_mid_Model

        Public Property ObjCol_物件NO As New Object
        Public Property ObjCol_部屋NO As New Object
        Public Property ObjCol_他業者による客付け時広告料の相殺処理募集 As New Object
        Public Property ObjCol_賃料交渉 As New Object
        Public Property ObjCol_案内時の照明器具 As New Object
        Public Property ObjCol_空室の清掃 As New Object
        Public Property ObjCol_案内時の営業車駐車可能スペース As New Object
        Public Property ObjCol_入居申込書 As New Object
        Public Property ObjCol_入居前に必要な工事箇所 As New Object
        Public Property ObjCol_入居時立会い As New Object
        Public Property ObjCol_他業者による客付け時の重要事項説明の確認 As New Object
        Public Property ObjCol_他業者による客付け時広告料の相殺処理契約 As New Object
        Public Property ObjCol_家主への契約金支払い As New Object
        Public Property ObjCol_敷金預かり先 As New Object
        Public Property ObjCol_広告料 As New Object
        Public Property ObjCol_解約違約金予告と短期 As New Object
        Public Property ObjCol_退去時の立会い As New Object
        Public Property ItemCnt As Integer

    End Class

#End Region


#Region "契約契約者情報モデル"

    Public Class Ky_kys_mid_Model

        Public Property ObjCol_物件No As New Object
        Public Property ObjCol_部屋No As New Object
        Public Property ObjCol_契約No As New Object
        Public Property ObjCol_契約管理レコードNo As New Object
        Public Property ObjCol_契約者No1 As New Object
        Public Property ObjCol_入居フラグ1 As New Object
        Public Property ObjCol_契約者No2 As New Object
        Public Property ObjCol_入居フラグ2 As New Object
        Public Property ObjCol_契約者No3 As New Object
        Public Property ObjCol_入居フラグ3 As New Object
        Public Property ItemCnt As Integer

    End Class

#End Region

#Region "契約保証人情報モデル"

    Public Class Ky_hosyonin_mid_Model

        Public Property ObjCol_物件No As New Object
        Public Property ObjCol_部屋No As New Object
        Public Property ObjCol_契約No As New Object
        Public Property ObjCol_契約管理レコードNo As New Object
        Public Property ObjCol_保証人使用契約者No As New Object
        Public Property ObjCol_保証人No1 As New Object
        Public Property ObjCol_保証人No2 As New Object
        Public Property ItemCnt As Integer

    End Class

#End Region

#Region "契約特約情報モデル"

    Public Class Ky_tokuyaku_mid_Model

        Public Property ObjCol_物件No As New Object
        Public Property ObjCol_部屋No As New Object
        Public Property ObjCol_契約No As New Object
        Public Property ObjCol_契約管理レコードNo As New Object
        Public Property ObjCol_原状回復特約内容 As New Object
        Public Property ObjCol_入居中修繕特約内容 As New Object
        Public Property ObjCol_その他特約内容 As New Object
        Public Property ItemCnt As Integer

    End Class

#End Region

End Namespace
