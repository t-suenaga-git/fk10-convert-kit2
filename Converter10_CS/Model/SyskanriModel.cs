
namespace Converter10.Njc.Model
{

    #region 初期設定基本情報モデル

    public class Profile_fk_base_Model
    {

        public string Vari_KenNo { get; set; }
        public string Vari_SiNo { get; set; }

    }

    #endregion

    #region 初期設定税編集情報モデル

    public class Profile_fk_zei_Model
    {

        public string Vari_No { get; set; }
        public string Vari_ZeiRit { get; set; }
        public string Vari_StartYmd { get; set; }
        public string Vari_EndYmd { get; set; }
        public string Vari_Biko { get; set; }

    }

    #endregion

    #region 初期設定変換文字情報モデル

    public class Profile_fk_henkanmoji_Model
    {

        public string Vari_No { get; set; }
        public string Vari_Target { get; set; }
        public string Vari_Cnv { get; set; }

    }

    #endregion

    #region 初期設定入金項目集約情報モデル

    public class Profile_fk_nkinkomkmerge_Model
    {

        public string Vari_Identity { get; set; }
        public string Vari_Name { get; set; }
        public string Vari_DaihyoNkinNo { get; set; }
        public Common.SetArrayData Vari_Int { get; set; }
        public int ItemCnt { get; set; }
        // Public Property Vari_Int1 As String
        // Public Property Vari_Int2 As String
        // Public Property Vari_Int3 As String
        // Public Property Vari_Int4 As String
        // Public Property Vari_Int5 As String
        // Public Property Vari_Int6 As String
        // Public Property Vari_Int7 As String
        // Public Property Vari_Int8 As String
        // Public Property Vari_Int9 As String
        // Public Property Vari_Int10 As String
        // Public Property Vari_Int11 As String
        // Public Property Vari_Int12 As String
        // Public Property Vari_Int13 As String
        // Public Property Vari_Int14 As String
        // Public Property Vari_Int15 As String
        // Public Property Vari_Int16 As String
        // Public Property Vari_Int17 As String
        // Public Property Vari_Int18 As String
        // Public Property Vari_Int19 As String
        // Public Property Vari_Int20 As String

        public Profile_fk_nkinkomkmerge_Model(int @int)
        {

            Vari_Int = new Common.SetArrayData(@int);
            ItemCnt = @int + 1;

        }

    }

    #endregion

}