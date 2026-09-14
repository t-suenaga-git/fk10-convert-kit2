
namespace Converter10.Njc.Model
{

    #region 部屋駐車場モデル

    public class Hy_cyusyajo_mid_Model
    {

        public object ObjCol_物件NO { get; set; } = new object();
        public object ObjCol_部屋NO { get; set; } = new object();
        public object ObjCol_駐車場の空き数手動設定用 { get; set; } = new object();
        public object ObjCol_駐車場の空き有無手動設定用 { get; set; } = new object();
        public object ObjCol_駐車場料金区分手動設定用 { get; set; } = new object();
        public object ObjCol_駐車場料金手動設定用 { get; set; } = new object();
        public object ObjCol_駐車場料金税区分手動設定用 { get; set; } = new object();
        public object ObjCol_駐車場の賃貸可能数 { get; set; } = new object();
        public object ObjCol_バイクの空き数手動設定用 { get; set; } = new object();
        public object ObjCol_バイク駐車場の空き有無手動設定用 { get; set; } = new object();
        public object ObjCol_バイク駐車場料金区分手動設定用 { get; set; } = new object();
        public object ObjCol_バイク駐車場料金手動設定用 { get; set; } = new object();
        public object ObjCol_バイク駐車場料金税区分手動設定用 { get; set; } = new object();
        public object ObjCol_バイク駐車場の賃貸可能数 { get; set; } = new object();
        public object ObjCol_駐輪場駐車場の空き数手動設定用 { get; set; } = new object();
        public object ObjCol_駐輪場駐車場の空き有無手動設定用 { get; set; } = new object();
        public object ObjCol_駐輪場駐車場料金区分手動設定用 { get; set; } = new object();
        public object ObjCol_駐輪場駐車場料金手動設定用 { get; set; } = new object();
        public object ObjCol_駐輪場駐車場料金税区分手動設定用 { get; set; } = new object();
        public object ObjCol_駐輪場駐車場の賃貸可能数 { get; set; } = new object();

        public int ItemCnt { get; set; }

    }

    #endregion

    #region 部屋特約モデル

    public class Hy_tokuyaku_mid_Model
    {

        public object ObjCol_物件NO { get; set; } = new object();
        public object ObjCol_部屋NO { get; set; } = new object();
        public object ObjCol_原状回復特約内容 { get; set; } = new object();
        public object ObjCol_入居中修繕特約内容 { get; set; } = new object();
        public object ObjCol_その他特約内容 { get; set; } = new object();
        public int ItemCnt { get; set; }

    }

    #endregion

    #region 部屋契約解約確認事項モデル

    public class Hy_kykaikakuninjiko_mid_Model
    {

        public object ObjCol_物件NO { get; set; } = new object();
        public object ObjCol_部屋NO { get; set; } = new object();
        public object ObjCol_他業者による客付け時広告料の相殺処理募集 { get; set; } = new object();
        public object ObjCol_賃料交渉 { get; set; } = new object();
        public object ObjCol_案内時の照明器具 { get; set; } = new object();
        public object ObjCol_空室の清掃 { get; set; } = new object();
        public object ObjCol_案内時の営業車駐車可能スペース { get; set; } = new object();
        public object ObjCol_入居申込書 { get; set; } = new object();
        public object ObjCol_入居前に必要な工事箇所 { get; set; } = new object();
        public object ObjCol_入居時立会い { get; set; } = new object();
        public object ObjCol_他業者による客付け時の重要事項説明の確認 { get; set; } = new object();
        public object ObjCol_他業者による客付け時広告料の相殺処理契約 { get; set; } = new object();
        public object ObjCol_家主への契約金支払い { get; set; } = new object();
        public object ObjCol_敷金預かり先 { get; set; } = new object();
        public object ObjCol_広告料 { get; set; } = new object();
        public object ObjCol_解約違約金予告と短期 { get; set; } = new object();
        public object ObjCol_退去時の立会い { get; set; } = new object();
        public int ItemCnt { get; set; }

    }

    #endregion


    #region 契約契約者情報モデル

    public class Ky_kys_mid_Model
    {

        public object ObjCol_物件No { get; set; } = new object();
        public object ObjCol_部屋No { get; set; } = new object();
        public object ObjCol_契約No { get; set; } = new object();
        public object ObjCol_契約管理レコードNo { get; set; } = new object();
        public object ObjCol_契約者No1 { get; set; } = new object();
        public object ObjCol_入居フラグ1 { get; set; } = new object();
        public object ObjCol_契約者No2 { get; set; } = new object();
        public object ObjCol_入居フラグ2 { get; set; } = new object();
        public object ObjCol_契約者No3 { get; set; } = new object();
        public object ObjCol_入居フラグ3 { get; set; } = new object();
        public int ItemCnt { get; set; }

    }

    #endregion

    #region 契約保証人情報モデル

    public class Ky_hosyonin_mid_Model
    {

        public object ObjCol_物件No { get; set; } = new object();
        public object ObjCol_部屋No { get; set; } = new object();
        public object ObjCol_契約No { get; set; } = new object();
        public object ObjCol_契約管理レコードNo { get; set; } = new object();
        public object ObjCol_保証人使用契約者No { get; set; } = new object();
        public object ObjCol_保証人No1 { get; set; } = new object();
        public object ObjCol_保証人No2 { get; set; } = new object();
        public int ItemCnt { get; set; }

    }

    #endregion

    #region 契約特約情報モデル

    public class Ky_tokuyaku_mid_Model
    {

        public object ObjCol_物件No { get; set; } = new object();
        public object ObjCol_部屋No { get; set; } = new object();
        public object ObjCol_契約No { get; set; } = new object();
        public object ObjCol_契約管理レコードNo { get; set; } = new object();
        public object ObjCol_原状回復特約内容 { get; set; } = new object();
        public object ObjCol_入居中修繕特約内容 { get; set; } = new object();
        public object ObjCol_その他特約内容 { get; set; } = new object();
        public int ItemCnt { get; set; }

    }

    #endregion

}