
namespace Converter10
{
    static class CVDBInfoModule
    {

        public const string CVDBINFO_DBNAME = "cv_dbinfo";
        public const string CVDBINFO_HEADER1 = "紐付用TBL名";
        public const string CVDBINFO_HEADER2 = "紐付用フィールド名";
        public const string CVDBINFO_HEADER3 = "CV用種別";
        public const string CVDBINFO_HEADER4 = "CV用項目名";
        public const string CVDBINFO_HEADER5 = "CV用最小値";
        public const string CVDBINFO_HEADER6 = "CV用最大値";
        public const string CVDBINFO_HEADER7 = "デフォルト値";
        public const string CVDBINFO_HEADER8 = "CV用キーNo";
        public const string CVDBINFO_HEADER9 = "CV用必須項目フラグ";
        public const string CVDBINFO_HEADER10 = "有無確認区分";
        public const string CVDBINFO_HEADER11 = "CV備考";

        /// <summary>
    /// コンバート用DB情報格納仮テーブル作成クエリ
    /// </summary>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string Get_CVDBInfo_CreateQry()
        {

            string rtn_qry = "";

            rtn_qry = rtn_qry + " CREATE TABLE " + CVDBINFO_DBNAME;
            rtn_qry = rtn_qry + " 	( ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER1 + " NCHAR(50), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER2 + " NCHAR(50), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER3 + " NCHAR(100), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER4 + " NCHAR(100), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER5 + " NCHAR(50), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER6 + " NCHAR(50), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER7 + " NCHAR(50), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER8 + " NCHAR(10), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER9 + " NCHAR(10), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER10 + " NCHAR(50), ";
            rtn_qry = rtn_qry + CVDBINFO_HEADER11 + " NCHAR(500), ";
            rtn_qry = rtn_qry + "  ";
            rtn_qry = rtn_qry + " 	); ";

            return rtn_qry;

        }

        /// <summary>
    /// 革命10DB情報とCV用情報を結合させて抽出するクエリ
    /// </summary>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string Qry_GetTableInfo()
        {

            string tmp_sql = "";

            tmp_sql = tmp_sql + " SELECT DISTINCT ";
            tmp_sql = tmp_sql + " 	 VW.* ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.CV用種別,'') AS CV用種別 ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.CV用項目名,'') AS CV用項目名 ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.CV用最小値,'') AS CV用最小値 ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.CV用最大値,'') AS CV用最大値 ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.デフォルト値,'') AS デフォルト値 ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.CV用キーNo,'') AS CV用キーNo ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.CV用必須項目フラグ,'') AS CV用必須項目フラグ ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.有無確認区分,'') AS 有無確認区分 ";
            tmp_sql = tmp_sql + " 	,ISNULL(CVDB.CV備考,'') AS CV備考 ";
            tmp_sql = tmp_sql + " FROM ";
            tmp_sql = tmp_sql + " ( ";
            tmp_sql = tmp_sql + " 	SELECT ";
            tmp_sql = tmp_sql + " 		 col.column_id AS [行No] ";
            tmp_sql = tmp_sql + " 		,CASE ";
            tmp_sql = tmp_sql + " 			WHEN index_column_id IS NULL THEN '' ";
            tmp_sql = tmp_sql + " 			ELSE '*' ";
            tmp_sql = tmp_sql + " 		 END AS [主キー] ";
            tmp_sql = tmp_sql + " 		,jpnname.value AS [種別] ";
            tmp_sql = tmp_sql + " 		,obj.NAME AS [TBL名] ";
            tmp_sql = tmp_sql + " 		,fldjpname.value AS [項目名] ";
            tmp_sql = tmp_sql + " 		,col.NAME AS [フィールド名] ";
            tmp_sql = tmp_sql + " 		,type_name(col.user_type_id) AS [属性] ";
            tmp_sql = tmp_sql + " 		,col.max_length AS [サイズ] ";
            tmp_sql = tmp_sql + " 		,CASE is_nullable ";
            tmp_sql = tmp_sql + " 			WHEN '1' THEN '○' ";
            tmp_sql = tmp_sql + " 			ELSE '' ";
            tmp_sql = tmp_sql + " 		 END AS [Null許容] ";
            tmp_sql = tmp_sql + " 		 ,obj.id ";
            tmp_sql = tmp_sql + " 		 /* ";
            tmp_sql = tmp_sql + " 		,fldcomment.value AS [説明] ";
            tmp_sql = tmp_sql + " 		,creater.value AS creator ";
            tmp_sql = tmp_sql + " 		,checker.value AS checker ";
            tmp_sql = tmp_sql + " 		,comment.value  AS [TBL説明] ";
            tmp_sql = tmp_sql + " 		*/ ";
            tmp_sql = tmp_sql + " 	FROM sys.sysobjects AS obj ";
            tmp_sql = tmp_sql + " 	LEFT JOIN sys.extended_properties AS jpnname ON obj.id = jpnname.major_id ";
            tmp_sql = tmp_sql + " 		AND jpnname.NAME = 'MS_Description' ";
            tmp_sql = tmp_sql + " 		AND jpnname.minor_id = 0 ";
            tmp_sql = tmp_sql + " 	LEFT JOIN sys.columns AS col ON obj.id = col.object_id ";
            tmp_sql = tmp_sql + " 	LEFT JOIN sys.extended_properties AS fldjpname ON col.object_id = fldjpname.major_id ";
            tmp_sql = tmp_sql + " 		AND col.column_id = fldjpname.minor_id ";
            tmp_sql = tmp_sql + " 		AND fldjpname.NAME = 'Jpfieldname' ";
            tmp_sql = tmp_sql + " 	LEFT JOIN sys.indexes AS I ON col.object_id = I.object_id ";
            tmp_sql = tmp_sql + " 		AND I.is_primary_key = 1 ";
            tmp_sql = tmp_sql + " 	LEFT JOIN sys.index_columns AS pkey ON I.object_id = pkey.object_id ";
            tmp_sql = tmp_sql + " 		AND col.column_id = pkey.column_id ";
            tmp_sql = tmp_sql + " 		AND I.index_id = pkey.index_id ";
            tmp_sql = tmp_sql + " 	WHERE obj.xtype = 'U' ";
            tmp_sql = tmp_sql + " ) AS VW ";
            tmp_sql = tmp_sql + " LEFT JOIN cv_dbinfo AS CVDB ";
            tmp_sql = tmp_sql + " ON  VW.TBL名 = CVDB.紐付用TBL名 ";
            tmp_sql = tmp_sql + " AND VW.フィールド名 = CVDB.紐付用フィールド名 ";
            // 2016.03.28 接続先不正によるエラー時の処理追加 -add
            tmp_sql = tmp_sql + " WHERE CVDB.CV用種別 <> '' ";          // 接続先が10DB以外の場合は正常に結合できないためここで結合できたか判定するために追加
            tmp_sql = tmp_sql + " ORDER BY [TBL名],id,[行No] ";

            return tmp_sql;

        }

        /// <summary>
    /// コンバート用DB情報格納仮テーブル挿入クエリ
    /// </summary>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string Get_CVDBInfo_InsertQry(string cvitem)
        {

            var rtn_qry = new System.Text.StringBuilder();

            switch (cvitem ?? "")
            {

                case "各マスタ情報-バス交通マスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_buskotu','buskotu_no','バス交通マスタ','バス交通No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_buskotu','buskotu_name','バス交通マスタ','バス会社名','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_buskotu','history','バス交通マスタ','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-バス停マスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_buskotu_stop','buskotu_no','バス停マスタ','バス交通No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_buskotu_stop','sortorder','バス停マスタ','バス停No','1','3000','','2','','','最大値は画面上で直接入力した値から取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_buskotu_stop','keito_name','バス停マスタ','系統名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_buskotu_stop','busstop_name','バス停マスタ','バス停名','','100','','','●','',''); ");
                        break;
                    }
                case "各マスタ情報-学校区マスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','ken_no','学校区マスタ','県No','1','47','','1','','都道府県',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','si_no','学校区マスタ','市No','1101','47382','','2','','市区町村',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','no','学校区マスタ','学校区No','1','100','','','','','最大値が不明なのでとりあえず100で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','add_cyo','学校区マスタ','町','','100','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','add_banti','学校区マスタ','番地','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','syogaku_name','学校区マスタ','小学校名','','100','','5','','','小学校/中学校名はいずれかが設定されていれば可なのでデータが存在するレコードを抽出しておく(このファイル上で必須項目としない)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','cyugaku_name','学校区マスタ','中学校名','','100','','6','','','小学校/中学校名はいずれかが設定されていれば可なのでデータが存在するレコードを抽出しておく(このファイル上で必須項目としない)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','history','学校区マスタ','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_koku_add','add_cyome','学校区マスタ','丁目','','100','','4','','',''); ");
                        break;
                    }
                case "各マスタ情報-鍵タイトルマスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_kagi_title','kagi_kbn','鍵タイトルマスタ','鍵区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_kagi_title','kagi_no','鍵タイトルマスタ','鍵No','1','50','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_kagi_title','kagi_name','鍵タイトルマスタ','鍵名称','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_kagi_title','history','鍵タイトルマスタ','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-クレーム分類設定内容":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_claim_rui','claim_ruikbn','クレーム分類設定内容','箇所・クレーム分類区分','1','2','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_claim_rui','claim_ruino','クレーム分類設定内容','分類No','1','50','','2','','','最大値が不明のため、とりあえず20で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_claim_rui','claim_ruisortorder','クレーム分類設定内容','ソート順','1','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_claim_rui','claim_name','クレーム分類設定内容','名称','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_claim_rui','claim_useflg','クレーム分類設定内容','使用有無','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_claim_rui','history','クレーム分類設定内容','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-エリアマスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_area','area_no','エリアマスタ','エリアNo','1','20','','1','','','最大値が不明のため、とりあえず20で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_area','area_sortorder','エリアマスタ','表示順','1','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_area','area_name','エリアマスタ','エリア名','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_area','area_useflg','エリアマスタ','使用有無','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_area','history','エリアマスタ','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-保険種類マスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hoken_rui','hoken_ruino','保険種類マスタ','保険種類No','1','9999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hoken_rui','hoken_ruiname','保険種類マスタ','保険種類名','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hoken_rui','hoken_ruikana','保険種類マスタ','保険種類カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hoken_rui','biko_kihon','保険種類マスタ','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hoken_rui','history','保険種類マスタ','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hoken_rui','useflg','保険種類マスタ','検索で見つけられるようにする','0','1','1','','','',''); ");
                        break;
                    }
                case "各マスタ情報-契約分類マスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','ky_ruino','契約分類マスタ','契約分類No','1','9999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','ky_ruiname','契約分類マスタ','契約分類名','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','ky_ruibiko','契約分類マスタ','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','ky_ruitutimm','契約分類マスタ','更新・解約通知期間(ヶ月)','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','ky_ruicolor','契約分類マスタ','契約分類色','-2147483648','2147483647','-16777216','','','','汎用の場合は色を限定して文字列で入力。数値範囲はとりあえずintの最大最小を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','teisyaku_flg','契約分類マスタ','定期借地借家権契約扱い有無','1','2','2','','','','2がチェックOFF'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','history','契約分類マスタ','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_ky_rui','useflg','契約分類マスタ','マスタ検索対象フラグ','0','1','1','','','',''); ");
                        break;
                    }
                case "各マスタ情報-特約マスタ":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_tokuyaku','tokuyaku_grpno','特約マスタ','特約区分','1','3','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_tokuyaku','tokuyaku_no','特約マスタ','特約No','1','9999','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_tokuyaku','tokuyaku_title','特約マスタ','特約タイトル','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_tokuyaku','tokuyaku_template','特約マスタ','特約詳細','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_tokuyaku','history','特約マスタ','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-変動費設定内容":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','rule_no','変動費設定内容','ルールNo','1','9999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','rule_name','変動費設定内容','ルール名','','50','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','tani','変動費設定内容','その他単位','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','tanisjis','変動費設定内容','その他単位SJIS','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','ryokin_cnt','変動費設定内容','料金数','1','5','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','hendo_rui','変動費設定内容','分類','1','5','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','hasuusyori_sel','変動費設定内容','端数処理','1','3','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','hasuusyosu_ptn','変動費設定内容','端数処理パターン','1','9','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','koukei','変動費設定内容','口径','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','biko_kihon','変動費設定内容','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','keijo_rui','変動費設定内容','計上分類','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','zei_kbnhayami','変動費設定内容','税区分早見','1','3','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','zei_kbntanka','変動費設定内容','税区分単価','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','history','変動費設定内容','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule','useflg','変動費設定内容','マスタ検索対象フラグ','0','1','1','','','',''); ");
                        break;
                    }
                case "各マスタ情報-変動費料金単価表":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','rule_no','変動費料金単価表','変動費ルールNo','1','9999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','item_no','変動費料金単価表','行No','1','10000','','2','','','画面上では999999まで入力可能だが制御が入っているため登録最大数を最大値に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','siyoryo','変動費料金単価表','使用量','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','ryokin1','変動費料金単価表','料金１','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','ryokin2','変動費料金単価表','料金２','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','ryokin3','変動費料金単価表','料金３','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','ryokin4','変動費料金単価表','料金４','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hendorule_itiran','ryokin5','変動費料金単価表','料金５','0','99999999999','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-備考タイトルマスタ": // 20160621 修繕関連移行処理追加(備考タイトルマスタを含める)
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo','memo_kbn','備考タイトルマスタ','備考区分','1','13','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo','memo_no','備考タイトルマスタ','備考No','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo','memo_sortorder','備考タイトルマスタ','表示順','0','0','','','','','使用されていない可能性あり'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo','memo_name','備考タイトルマスタ','備考タイトル','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo','memo_useflg','備考タイトルマスタ','使用有無','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo','memo_usehojoitems','備考タイトルマスタ','補助アイテム使用有無','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo','history','備考タイトルマスタ','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-備考入力補助リストマスタ": // 20160621 修繕関連移行処理追加(備考タイトルマスタを含める)
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo_lst','memo_kbn','備考入力補助リストマスタ','備考区分','1','13','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo_lst','memo_no','備考入力補助リストマスタ','備考No','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo_lst','memo_lstno','備考入力補助リストマスタ','備考入力補助リストNo','1','20','','3','','','画面上から最大値を取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo_lst','memo_lstname','備考入力補助リストマスタ','備考入力補助リスト内容','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_memo_lst','history','備考入力補助リストマスタ','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "各マスタ情報-画像タイトルマスタ": // 20160720 連動情報構築
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_gazo_title','gazo_kbn','画像タイトルマスタ','画像区分','1','10','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_gazo_title','gazo_no','画像タイトルマスタ','画像No','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_gazo_title','gazo_name','画像タイトルマスタ','画像名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_gazo_title','history','画像タイトルマスタ','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "業者情報-仲介業者基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','gy_fudono','仲介業者基本情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','gy_fudoname','仲介業者基本情報','業者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','gy_fudonamesjis','仲介業者基本情報','業者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','gy_fudokana','仲介業者基本情報','業者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','post_code','仲介業者基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','addr_kenno','仲介業者基本情報','都道府県コード','1','47','','','','都道府県',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','addr_sino','仲介業者基本情報','市区町村コード','1101','47382','','','','市区町村',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','addr_cyo','仲介業者基本情報','町地域','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','addr_cyome','仲介業者基本情報','丁番地','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','addr_cyomeptn','仲介業者基本情報','丁番地選択','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','addr_banti','仲介業者基本情報','街区番号地番','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','addr_etc','仲介業者基本情報','その他','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','tel1','仲介業者基本情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','tel2','仲介業者基本情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','fax','仲介業者基本情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','mobiletel1','仲介業者基本情報','携帯1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','mobiletel2','仲介業者基本情報','携帯2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','daihyo_yakusyoku','仲介業者基本情報','代表者役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','daihyo_name','仲介業者基本情報','代表者名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','tanto_busyoyakusyoku','仲介業者基本情報','担当者部署役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','tanto_name','仲介業者基本情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','tanto_namesjis','仲介業者基本情報','担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','tanto_kana','仲介業者基本情報','担当者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','mail','仲介業者基本情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','mobilemail','仲介業者基本情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','url','仲介業者基本情報','Webアドレス(URL)','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','yusentel_kbn','仲介業者基本情報','優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','yusenmail_kbn','仲介業者基本情報','優先設定(メールアドレス)','1','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','biko_kihon','仲介業者基本情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','sofu_hakkofurikomi','仲介業者基本情報','口座振込通知書発行の許可','1','2','1','','','','1がチェックON'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','menkyo_bango','仲介業者基本情報','免許証番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','menkyo_ymd','仲介業者基本情報','免許年月日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','syunin_bango','仲介業者基本情報','取引主任者登録番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','syunin_name','仲介業者基本情報','取引主任者名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','kanrikyokai_bango','仲介業者基本情報','全国賃貸不動産管理業協会会員番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','keieikanrisi_bango','仲介業者基本情報','賃貸不動産経営管理士登録番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','keieikanrisi_name','仲介業者基本情報','賃貸不動産経営管理士名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','kanrigy_bango','仲介業者基本情報','賃貸住宅管理業者登録番号(仮)','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','useflg','仲介業者基本情報','検索対象フラグ','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','history','仲介業者基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','rowid','仲介業者基本情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','daihyo_namesjis','仲介業者基本情報','代表者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','syunin_namesjis','仲介業者基本情報','取引主任者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','tanto_autoinputflg','仲介業者基本情報','仲介業者契約担当者自動入力フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudo','keieikanrisi_namesjis','仲介業者基本情報','賃貸不動産経営管理士名(SJIS)','','100','','','','',''); ");
                        break;
                    }
                case "業者情報-仲介業者口座情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','gy_fudono','仲介業者口座情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','gy_fudokozano','仲介業者口座情報','口座No','1','1','','2','','','1件のみ入力可能だが口座コードを設定する箇所が存在する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','kinyu_no','仲介業者口座情報','金融機関No','1','9999','','','','金融機関',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','kinyu_tenno','仲介業者口座情報','金融機関支店No','1','9999','','','','金融機関支店',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','koza_syubetu','仲介業者口座情報','口座種別','1','9','1','','','口座種別',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','koza_bango','仲介業者口座情報','口座番号','1','9999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','koza_meigi','仲介業者口座情報','口座名義','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','koza_meigikana','仲介業者口座情報','口座名義カナ','','50','','','','口座名義カナ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','biko_koza','仲介業者口座情報','備考(口座情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','biko_furikomi','仲介業者口座情報','備考(口座振込情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','sgfirai_kbn','仲介業者口座情報','総合振込区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','sgfirai_no','仲介業者口座情報','振込依頼人No','1','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','sgfirai_tesufutankbn','仲介業者口座情報','振込手数料負担区分','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','sgfirai_tesukeisankbn','仲介業者口座情報','振込手数料計算区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','sgfirai_tesukotei1gak','仲介業者口座情報','振込手数料固定額1','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','sgfirai_tesukotei2gak','仲介業者口座情報','振込手数料固定額2','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','biko_sgfirai','仲介業者口座情報','備考(総合振込情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','history','仲介業者口座情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','yucyokoza_kigo1','仲介業者口座情報','ゆうちょ銀行記号1','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','yucyokoza_kigo2','仲介業者口座情報','ゆうちょ銀行記号2','','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudokoza','yucyokoza_bango','仲介業者口座情報','ゆうちょ銀行番号','','8','','','','',''); ");
                        break;
                    }
                case "業者情報-仲介業者メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudomemo','gy_fudono','仲介業者メモ情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudomemo','memo_no','仲介業者メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudomemo','memo','仲介業者メモ情報','内容','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_fudomemo','history','仲介業者メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "業者情報-保険業者基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','gy_hokenno','保険業者基本情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','gy_hokenname','保険業者基本情報','業者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','gy_hokennamesjis','保険業者基本情報','業者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','gy_hokenkana','保険業者基本情報','業者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','post_code','保険業者基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','addr1','保険業者基本情報','住所1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','addr2','保険業者基本情報','住所2','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','tel1','保険業者基本情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','tel2','保険業者基本情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','fax','保険業者基本情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','mobiletel1','保険業者基本情報','携帯1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','mobiletel2','保険業者基本情報','携帯2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','tanto_busyoyakusyoku','保険業者基本情報','担当者部署役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','tanto_name','保険業者基本情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','tanto_namesjis','保険業者基本情報','担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','tanto_kana','保険業者基本情報','担当者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','mail','保険業者基本情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','mobilemail','保険業者基本情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','url','保険業者基本情報','Webアドレス(URL)','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','yusentel_kbn','保険業者基本情報','優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','yusenmail_kbn','保険業者基本情報','優先設定(メールアドレス)','1','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','biko_kihon','保険業者基本情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','useflg','保険業者基本情報','検索対象フラグ','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','history','保険業者基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hoken','rowid','保険業者基本情報','付箋Guid','','16','','','','',''); ");
                        break;
                    }
                case "業者情報-保険業者口座情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','gy_hokenno','保険業者口座情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','gy_hokenkozano','保険業者口座情報','口座No','1','1','','2','','','1件のみ入力可能だが口座コードを設定する箇所が存在する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','kinyu_no','保険業者口座情報','金融機関No','1','9999','','','','金融機関',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','kinyu_tenno','保険業者口座情報','金融機関店No','1','9999','','','','金融機関支店',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','koza_syubetu','保険業者口座情報','口座種別','1','9','1','','','口座種別',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','koza_bango','保険業者口座情報','口座番号','1','9999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','koza_meigi','保険業者口座情報','口座名義','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','koza_meigikana','保険業者口座情報','口座名義カナ','','50','','','','口座名義カナ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','biko_koza','保険業者口座情報','備考(口座情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','biko_furikomi','保険業者口座情報','備考(口座振込情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','sgfirai_kbn','保険業者口座情報','総合振込区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','sgfirai_no','保険業者口座情報','振込依頼人No','1','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','sgfirai_tesufutankbn','保険業者口座情報','振込手数料負担区分','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','sgfirai_tesukeisankbn','保険業者口座情報','振込手数料計算区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','sgfirai_tesukotei1gak','保険業者口座情報','振込手数料固定額1','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','sgfirai_tesukotei2gak','保険業者口座情報','振込手数料固定額2','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','biko_sgfirai','保険業者口座情報','備考(総合振込情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenkoza','history','保険業者口座情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "業者情報-保険業者メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenmemo','gy_hokenno','保険業者メモ情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenmemo','memo_no','保険業者メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenmemo','memo','保険業者メモ情報','内容','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hokenmemo','history','保険業者メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "業者情報-家賃保証業者基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','gy_hosyono','家賃保証業者基本情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','gy_hosyoname','家賃保証業者基本情報','業者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','gy_hosyonamesjis','家賃保証業者基本情報','業者名(shift-jis)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','gy_hosyokana','家賃保証業者基本情報','業者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','post_code','家賃保証業者基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','addr1','家賃保証業者基本情報','住所1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','addr2','家賃保証業者基本情報','住所2','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','tel1','家賃保証業者基本情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','tel2','家賃保証業者基本情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','fax','家賃保証業者基本情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','mobiletel1','家賃保証業者基本情報','携帯1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','mobiletel2','家賃保証業者基本情報','携帯2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','tanto_busyoyakusyoku','家賃保証業者基本情報','担当者部署役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','tanto_name','家賃保証業者基本情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','tanto_namesjis','家賃保証業者基本情報','担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','tanto_kana','家賃保証業者基本情報','担当者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','mail','家賃保証業者基本情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','mobilemail','家賃保証業者基本情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','url','家賃保証業者基本情報','Webアドレス(URL)','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','yusentel_kbn','家賃保証業者基本情報','優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','yusenmail_kbn','家賃保証業者基本情報','優先設定(メールアドレス)','1','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','biko_kihon','家賃保証業者基本情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','useflg','家賃保証業者基本情報','検索対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','history','家賃保証業者基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','rowid','家賃保証業者基本情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyo','rendo_namekbn','家賃保証業者基本情報','連動用家賃保証会社区分','-1','7','-1','','','',''); ");
                        break;
                    }
                case "業者情報-家賃保証業者メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyomemo','gy_hosyono','家賃保証業者メモ情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyomemo','memo_no','家賃保証業者メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyomemo','memo','家賃保証業者メモ情報','内容','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_hosyomemo','history','家賃保証業者メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "業者情報-修繕業者基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','gy_szenno','修繕業者基本情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','gy_szenname','修繕業者基本情報','業者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','gy_szennamesjis','修繕業者基本情報','業者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','gy_szenkana','修繕業者基本情報','業者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','post_code','修繕業者基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','addr1','修繕業者基本情報','住所1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','addr2','修繕業者基本情報','住所2','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','tel1','修繕業者基本情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','tel2','修繕業者基本情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','fax','修繕業者基本情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','mobiletel1','修繕業者基本情報','携帯1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','mobiletel2','修繕業者基本情報','携帯2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','daihyo_yakusyoku','修繕業者基本情報','代表者役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','daihyo_name','修繕業者基本情報','代表者名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','tanto_busyoyakusyoku','修繕業者基本情報','担当者部署役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','tanto_name','修繕業者基本情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','tanto_namesjis','修繕業者基本情報','担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','tanto_kana','修繕業者基本情報','担当者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','mail','修繕業者基本情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','mobilemail','修繕業者基本情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','url','修繕業者基本情報','Webアドレス(URL)','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','yusentel_kbn','修繕業者基本情報','優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','yusenmail_kbn','修繕業者基本情報','優先設定(メールアドレス)','1','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','biko_kihon','修繕業者基本情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','useflg','修繕業者基本情報','検索対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','history','修繕業者基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','rowid','修繕業者基本情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szen','daihyo_namesjis','修繕業者基本情報','代表者名(SJIS)','','100','','','','',''); ");
                        break;
                    }
                case "業者情報-修繕業者口座情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','gy_szenno','修繕業者口座情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','gy_szenkozano','修繕業者口座情報','口座No','1','1','1','2','','','1件のみ入力可能だが口座コードを設定する箇所が存在する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','kinyu_no','修繕業者口座情報','金融機関No','1','9999','','','','金融機関',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','kinyu_tenno','修繕業者口座情報','金融機関店No','1','9999','','','','金融機関支店',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','koza_syubetu','修繕業者口座情報','口座種別','1','9','1','','','口座種別',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','koza_bango','修繕業者口座情報','口座番号','1','9999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','koza_meigi','修繕業者口座情報','口座名義','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','koza_meigikana','修繕業者口座情報','口座名義カナ','','50','','','','口座名義カナ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','yucyokoza_kigo1','修繕業者口座情報','ゆうちょ銀行記号1','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','yucyokoza_kigo2','修繕業者口座情報','ゆうちょ銀行記号2','','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','yucyokoza_bango','修繕業者口座情報','ゆうちょ銀行番号','','8','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','biko_koza','修繕業者口座情報','備考(口座情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','biko_furikomi','修繕業者口座情報','備考(口座振込情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','sgfirai_kbn','修繕業者口座情報','総合振込区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','sgfirai_no','修繕業者口座情報','振込依頼人No','1','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','sgfirai_tesufutankbn','修繕業者口座情報','振込手数料負担区分','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','sgfirai_tesukeisankbn','修繕業者口座情報','振込手数料計算区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','sgfirai_tesukotei1gak','修繕業者口座情報','振込手数料固定額1','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','sgfirai_tesukotei2gak','修繕業者口座情報','振込手数料固定額2','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','biko_sgfirai','修繕業者口座情報','備考(総合振込情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenkoza','history','修繕業者口座情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "業者情報-修繕業者メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenmemo','gy_szenno','修繕業者メモ情報','業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenmemo','memo_no','修繕業者メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenmemo','memo','修繕業者メモ情報','内容','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_szenmemo','history','修繕業者メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "業者情報-ライフライン業者情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','gy_lifelineno','ライフライン業者情報','ライフライン業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','gy_lifelinename','ライフライン業者情報','ライフライン業者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','gy_lifelinenamesjis','ライフライン業者情報','ライフライン業者名(Shift-jis)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','gy_lifelinekana','ライフライン業者情報','ライフライン業者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','eigyosyo','ライフライン業者情報','営業所','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','eigyosyosjis','ライフライン業者情報','営業所(Shift-jis)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','post_code','ライフライン業者情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','addr1','ライフライン業者情報','住所1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','addr2','ライフライン業者情報','住所2','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','tel1','ライフライン業者情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','tel2','ライフライン業者情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','fax','ライフライン業者情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','mobiletel1','ライフライン業者情報','携帯1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','mobiletel2','ライフライン業者情報','携帯2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','yusentel_kbn','ライフライン業者情報','優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','biko_kihon','ライフライン業者情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','useflg','ライフライン業者情報','検索対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','history','ライフライン業者情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','rowid','ライフライン業者情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','lifeline_denkiflg','ライフライン業者情報','電気','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','lifeline_josuidoflg','ライフライン業者情報','上水道','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','lifeline_gasflg','ライフライン業者情報','ガス','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','lifeline_toyuflg','ライフライン業者情報','灯油','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','lifeline_other1flg','ライフライン業者情報','その他','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_lifeline','lifeline_haisuiflg','ライフライン業者情報','排水','0','1','0','','','',''); ");
                        break;
                    }
                case "業者情報-施工業者情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','gy_sekono','施工業者情報','施工業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','gy_sekoname','施工業者情報','施工業者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','gy_sekonamesjis','施工業者情報','施工業者名(Shift-jis)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','gy_sekokana','施工業者情報','施工業者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','eigyosyo','施工業者情報','営業所','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','eigyosyosjis','施工業者情報','営業所(Shift-jis)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','post_code','施工業者情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','addr1','施工業者情報','住所1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','addr2','施工業者情報','住所2','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','tel1','施工業者情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','tel2','施工業者情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','fax','施工業者情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','mobiletel1','施工業者情報','携帯1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','mobiletel2','施工業者情報','携帯2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','tanto_busyoyakusyoku','施工業者情報','担当者部署役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','tanto_name','施工業者情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','tanto_namesjis','施工業者情報','担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','tanto_kana','施工業者情報','担当者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','mail','施工業者情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','mobilemail','施工業者情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','url','施工業者情報','Webアドレス(URL)','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','yusentel_kbn','施工業者情報','優先設定(電話番号)','0','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','yusenmail_kbn','施工業者情報','優先設定(メールアドレス)','0','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','biko_kihon','施工業者情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','useflg','施工業者情報','検索対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','history','施工業者情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_seko','rowid','施工業者情報','付箋Guid','','16','','','','',''); ");
                        break;
                    }
                case "業者情報-施設保守業者情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','gy_sisetuno','施設保守業者情報','施設保守業者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','gy_sisetuname','施設保守業者情報','施設保守業者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','gy_sisetunamesjis','施設保守業者情報','施設保守業者名(Shift-jis)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','gy_sisetukana','施設保守業者情報','施設保守業者名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','eigyosyo','施設保守業者情報','営業所','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','eigyosyosjis','施設保守業者情報','営業所(Shift-jis)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','post_code','施設保守業者情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','addr1','施設保守業者情報','住所1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','addr2','施設保守業者情報','住所2','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','tel1','施設保守業者情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','tel2','施設保守業者情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','fax','施設保守業者情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','mobiletel1','施設保守業者情報','携帯1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','mobiletel2','施設保守業者情報','携帯2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','tanto_busyoyakusyoku','施設保守業者情報','担当者部署役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','tanto_name','施設保守業者情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','tanto_namesjis','施設保守業者情報','担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','tanto_kana','施設保守業者情報','担当者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','mail','施設保守業者情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','mobilemail','施設保守業者情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','url','施設保守業者情報','Webアドレス(URL)','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','yusentel_kbn','施設保守業者情報','優先設定(電話番号)','0','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','yusenmail_kbn','施設保守業者情報','優先設定(メールアドレス)','0','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','biko_kihon','施設保守業者情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','useflg','施設保守業者情報','検索対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','history','施設保守業者情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('gydata_sisetu','rowid','施設保守業者情報','付箋Guid','','16','','','','',''); ");
                        break;
                    }

                case "自社情報-自社基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_no','自社基本情報','自社・支店No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_name','自社基本情報','自社・支店名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_namesjis','自社基本情報','自社・支店名（SJIS）','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_kana','自社基本情報','自社・支店カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_name2','自社基本情報','自社・支店名２','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_name2sjis','自社基本情報','自社・支店名２（SJIS）','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_name3','自社基本情報','自社・支店名３','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','jisya_name3sjis','自社基本情報','自社・支店名３（SJIS）','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','daihyo_yakusyoku','自社基本情報','代表者役職','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','daihyo_name','自社基本情報','代表者名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','post_code','自社基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','addr_kenno','自社基本情報','都道府県番号','1','47','','','','都道府県',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','addr_sino','自社基本情報','市区町村番号','1101','47382','','','','市区町村',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','addr_cyo','自社基本情報','町地域','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','addr_cyome','自社基本情報','丁番地','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','addr_cyomeptn','自社基本情報','丁番地パターン','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','addr_banti','自社基本情報','街区番号地番','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','addr_etc','自社基本情報','その他','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','tel1','自社基本情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','tel2','自社基本情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','mobiletel1','自社基本情報','携帯TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','mobiletel2','自社基本情報','携帯TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','fax','自社基本情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','mail','自社基本情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','mobilemail','自社基本情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','url','自社基本情報','Webアドレス','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','biko_kihon','自社基本情報','備考（基本情報）','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','menkyo_bango','自社基本情報','免許証番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','menkyo_ymd','自社基本情報','免許年月日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','syunin_bango','自社基本情報','取引主任者登録番号（代表）','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','syunin_name','自社基本情報','取引主任者名（代表）','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanrikyokai_bango','自社基本情報','全国賃貸不動産管理業協会会員番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','keieikanrisi_bango','自社基本情報','賃貸不動産経営管理士登録番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','keieikanrisi_name','自社基本情報','賃貸不動産経営管理士名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanrigy_bango','自社基本情報','賃貸住宅管理業者登録番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai1','自社基本情報','加入団体１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai2','自社基本情報','加入団体２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai3','自社基本情報','加入団体３','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai4','自社基本情報','加入団体４','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai5','自社基本情報','加入団体５','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai6','自社基本情報','加入団体６','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai7','自社基本情報','加入団体７','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai8','自社基本情報','加入団体８','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai9','自社基本情報','加入団体９','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','kanyu_dantai10','自社基本情報','加入団体１０','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','history','自社基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','useflg','自社基本情報','マスタ検索対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','daihyo_namesjis','自社基本情報','代表者名（SJIS）','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','syunin_namesjis','自社基本情報','取引主任者名（代表）（SJIS）','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata','keieikanrisi_namesjis','自社基本情報','賃貸不動産経営管理士名（SJIS）','','100','','','','',''); ");
                        break;
                    }
                case "自社情報-自社口座情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','jisya_no','自社口座情報','自社支店','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','jisya_kozano','自社口座情報','口座No','1','999','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','kinyu_no','自社口座情報','金融機関No','1','9999','','','','金融機関_自社口座',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','kinyu_tenno','自社口座情報','金融機関店No','1','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','koza_syubetu','自社口座情報','口座種別','1','9','','','','口座種別',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','koza_bango','自社口座情報','口座番号','1','9999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','koza_meigi','自社口座情報','口座名義','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','koza_meigikana','自社口座情報','口座名義カナ','','100','','','','口座名義カナ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','yucyokoza_kigo1','自社口座情報','ゆうちょ口座記号１','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','yucyokoza_kigo2','自社口座情報','ゆうちょ口座記号２','','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','yucyokoza_bango','自社口座情報','ゆうちょ口座番号','','8','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','biko_koza','自社口座情報','備考(口座情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','history','自社口座情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_koza','useflg','自社口座情報','マスタ検索対象フラグ','0','1','1','','','',''); ");
                        break;
                    }
                case "自社情報-自社担当者情報": // 20160608 自社担当者情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','logonuser_no','自社担当者情報','賃貸革命ログオンユーザNo','1','99999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','logonuser_name','自社担当者情報','賃貸革命ログオンユーザ名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','logonuser_namesjis','自社担当者情報','賃貸革命ログオンユーザ名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','logonuser_kana','自社担当者情報','賃貸革命ログオンユーザ名カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','logongroup_no','自社担当者情報','賃貸革命ログオングループNo','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','tel1','自社担当者情報','TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','tel2','自社担当者情報','TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','mobiletel1','自社担当者情報','携帯電話1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','fax1','自社担当者情報','FAX1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','mailaddress1','自社担当者情報','メールアドレス1','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','mobilemailaddress1','自社担当者情報','携帯メールアドレス1','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','takkenmenkyo','自社担当者情報','宅建免許番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','biko','自社担当者情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','denylogon','自社担当者情報','担当者ログオン禁止フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','noassign','自社担当者情報','使用しないフラグ','0','1','0','','','','0：チェックON'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','password','自社担当者情報','ログオンパスワード','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','behaviourdefine_base64','自社担当者情報','behaviourdefineクラスをシリアライズしてbase64したデータ','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','colordefine_base64','自社担当者情報','colordefineクラスをシリアライズしてbase64したデータ','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','n3spreadcolordefine_base64','自社担当者情報','n3spreadcolordefineクラスをシリアライズしてbase64したデータ','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','menudesigndefine_base64','自社担当者情報','menudesigndefineクラスをシリアライズしてbase64したデータ','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','smtpdefine_base64','自社担当者情報','smtpdefineクラスをシリアライズしてbase64したデータ','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','history','自社担当者情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','sendtargetsiteid','自社担当者情報','送信対象サイトID','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','windowsuser_name','自社担当者情報','windowsログオンユーザー名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','hysearchdefine_base64','自社担当者情報','hysearchdefineクラスをシリアライズしてbase64したデータ','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','csvsetting_base64','自社担当者情報','csvsettingクラスをシリアライズしてbase64したデータ','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_logonuser','dontshowfeedbackagain','自社担当者情報','意見画面表示フラグ','0','1','0','','','',''); ");
                        break;
                    }
                case "自社情報-自社メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_memo','jisya_no','自社メモ情報','自社・支店No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_memo','memo_no','自社メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_memo','memo','自社メモ情報','メモ','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('jisyadata_memo','history','自社メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "自社情報-振込依頼人情報":
                    {
                        // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','sgfirai_no','振込依頼人情報','振込依頼人No','1','9999','','1','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','sgfirai_name','振込依頼人情報','振込依頼人名','','200','','','●','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','jisya_no','振込依頼人情報','自社支店No.','1','999999999','','','','自社_口座用',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','jisya_kozano','振込依頼人情報','自社口座No','1','999','','','','自社口座_口座用',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','sgfirainin_code','振込依頼人情報','振込依頼人コード','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','sgfirainin_kana','振込依頼人情報','振込依頼人カナ','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','biko_kihon','振込依頼人情報','備考(基本情報)','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','sgfzenfmt_no','振込依頼人情報','総合振込全銀フォーマットNo','1','999999999','','','','','最大桁数リストになかったため画面上から最大値を取得'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','file_sosin','振込依頼人情報','送信ファイルパス','','255','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','crlf','振込依頼人情報','改行有無(0:改行しない 1:改行する)','0','1','1','','','','10画面上での設定箇所が不明'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','changewo_flg','振込依頼人情報','ヲ変換フラグ(0:変換しない 1:変換する)','0','1','1','','','','10画面上での設定箇所が不明'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','biko_data','振込依頼人情報','備考(データ)','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','keisandefault_flg','振込依頼人情報','計算時の初期値フラグ(0:初期値としない 1:初期値とする)','0','1','0','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','futancyousei_flg','振込依頼人情報','負担調整フラグ(0:調整しない 1:調整する)','0','1','0','','','','10画面上での設定箇所が不明'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','biko_tesu','振込依頼人情報','備考(手数料)','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','history','振込依頼人情報','履歴','','-1','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','sgfdata_createflg','振込依頼人情報','総合振込データ生成フラグ','0','1','0','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_sgfirai','useflg','振込依頼人情報','マスタ検索対象フラグ','0','1','1','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','sgfirai_no','振込依頼人情報','振込依頼人No','1','9999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','sgfirai_name','振込依頼人情報','振込依頼人名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','jisya_no','振込依頼人情報','自社支店No','1','999999999','','','','自社_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','jisya_kozano','振込依頼人情報','自社口座No','1','999','','','','自社口座_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','sgfirainin_code','振込依頼人情報','振込依頼人コード','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','sgfirainin_kana','振込依頼人情報','振込依頼人カナ','','100','','','','口座名義カナ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','biko','振込依頼人情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','sgfzenfmt_no','振込依頼人情報','総合振込全銀フォーマットNo','1','999999999','','','','FB関連_総合振込','最大桁数リストになかったため画面上から最大値を取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','file_sosin','振込依頼人情報','送信ファイルパス','','255','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','changewo_flg','振込依頼人情報','ヲ変換フラグ(0:変換しない 1:変換する)','0','1','1','','','','10画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','futancyousei_flg','振込依頼人情報','負担調整フラグ(0:調整しない 1:調整する)','0','1','0','','','','10画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','history','振込依頼人情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','sgfdata_createflg','振込依頼人情報','総合振込データ生成フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfirai','useflg','振込依頼人情報','マスタ検索対象フラグ','0','1','1','','','',''); ");
                        break;
                    }
                // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg end
                case "自社情報-振込手数料情報":     // 20160517 振込手数料情報の新規作成
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','sgfirai_no','振込手数料情報','振込依頼人No','1','9999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','rec_no','振込手数料情報','レコードNo','0','11','','2','','','V7の最大値に合わせる (10の行Noは0から始まる)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','from_gak','振込手数料情報','金額From','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','to_gak','振込手数料情報','金額To','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','doukoudouten_densingak','振込手数料情報','同行同支店手数料(電信扱)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','doukoudouten_bunsyogak','振込手数料情報','同行同支店手数料(文書扱)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','doukoutaten_densingak','振込手数料情報','同行他支店手数料(電信扱)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','doukoutaten_bunsyogak','振込手数料情報','同行他支店手数料(文書扱)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','takou_densingak','振込手数料情報','他行手数料(電信扱)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_sgfiraitesu','takou_bunsyogak','振込手数料情報','他行手数料(文書扱)','0','99999999999','','','','',''); ");
                        break;
                    }
                case "自社情報-口座振替情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','fkae_no','口座振替情報','振替情報No','1','8999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','fkae_name','口座振替情報','振替情報名称','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','fkae_kana','口座振替情報','振替情報カナ名称','','100','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 口座振替情報 -del
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_fkaejyoho','servicetype','口座振替情報','サービスタイプ　1:オリコ　2:ジャックス　3:アプラス','0','0','','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','fkae_fb_fkomiraino','口座振替情報','振込依頼人','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','fkae_fb_fkomiraikana','口座振替情報','振込依頼人カナ名','','100','','','','口座名義カナ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','kamei_no','口座振替情報','加盟店No 依頼人番号と共用','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','jlease_nyukinkbn','口座振替情報','ジェイリース　入金区分　1:変更なし・・・・','0','0','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','jisya_no','口座振替情報','自社支店No','1','999999999','0','','','自社_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','fkae_fb_nkinukekozano','口座振替情報','振替先口座No','1','999','0','','','自社口座_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','hikiotosibi','口座振替情報','引落日','0','31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','tesu_gak','口座振替情報','手数料　請求額','0','9999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','biko','口座振替情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','fkae_fb_orgfmtno','口座振替情報','オリジナルフォーマットNo','1','999999999','','','','FB関連_口座振替','最大桁数リストになかったため画面上から最大値を取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','datasort','口座振替情報','データ並び替え　0:物件番号と部屋番号　1:契約者番号','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','keiyakusyano_syuturyoku','口座振替情報','契約者Noを契約者番号に出力する  0:出力しない　1:出力する','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','crlf','口座振替情報','改行コード(CRLF)出力　0:なし　1:あり','-1','4','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','wo_mojihenkan','口座振替情報','「ｦ」→「ｵ」変換　0:しない　1:する','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','file_sosin','口座振替情報','送信ファイルパス名','','500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','file_jyusin','口座振替情報','受信ファイルパス名','','500','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 口座振替情報 -del
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_fkaejyoho','fdsakusei','口座振替情報','2枚ＦＤ作成手順を使用する　0:しない　1:する','1','2','2','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','seikyu_tani','口座振替情報','請求のまとめ方　0:請求先単位　1:契約単位','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','seikyu_taino','口座振替情報','滞納分請求　0:しない　1:する','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','seikyu_tukitani','口座振替情報','滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける','-1','1','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','tesu_nyukinkanri','口座振替情報','手数料の請求入金管理をする　0:しない　1:する','1','1','1','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','tesu_kurikosi','口座振替情報','未入金の手数料は次回請求に加える　0:加えない　1:加える','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','syogorule','口座振替情報','照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合','1','1','1','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','yutyobango','口座振替情報','ゆうちょ銀行金融機関番号指定　不使用の場合はNULL','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','history','口座振替情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','useflg','口座振替情報','マスタ検索対象フラグ','0','1','1','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 口座振替情報 -add sta
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','multi_flg','口座振替情報','マルチヘッダー形式フラグ','0','0','0','','','','10画面上での設定箇所が不明の為とりあえず0を設定'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','yutyo_appendcode','口座振替情報','ゆうちょ銀行付加コード','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','yutyo_usefkomcode','口座振替情報','振替でのゆうちょ銀行コードに振込用変換を利用するフラグ','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','resultfile_notuse','口座振替情報','振替入金処理時に受信ファイルを読み込まないフラグ','1','2','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 口座振替情報 -add end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkaejyoho','saifkae_use','口座振替情報','再振替対応利用フラグ','1','2','2','','','',''); ");  // 20161012 革命10アップデートに伴う修正 -add
                        break;
                    }
                case "自社情報-入出金取得情報":     // 20160517 入出金取得情報の新規作成
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','ns_no','入出金取得情報','入出金取得No','1','9999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','ns_name','入出金取得情報','設定名称','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','ns_kana','入出金取得情報','設定カナ名称','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','jisya_no','入出金取得情報','自社支店No','1','999999999','','','','FB関連_入出金自社No',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','orgfmt_no','入出金取得情報','フォーマット定義名','1','999999999','','','●','FB関連_入出金フォーマット',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','file_jyusin','入出金取得情報','受信ファイル名（フルパス）','','500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','data_syubetu','入出金取得情報','取込データ種別','0','0','','','','','10画面上での設定箇所が不明の為とりあえず0を設定しておく'); ");
                        // 20160519 EXEUpdateに伴う修正 入出金取得情報 -del
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('m_fb_nskinsetting','crlf','入出金取得情報','改行コード有無','-1','4','-1','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','biko','入出金取得情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','history','入出金取得情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','useflg','入出金取得情報','使用フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','yatin_kozano','入出金取得情報','家賃口座No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','file_sosin','入出金取得情報','送信ファイル名（フルパス）','','500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','yokinbunkatusetting','入出金取得情報','預金分割設定','','-1','','','','','京王カスタマイズ'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_nskinsetting','duplicate_check','入出金取得情報','重複チェック','0','1','0','','','',''); ");
                        break;
                    }
                case "自社情報-家賃入金口座情報":     // 20160517 家賃入金口座情報の新規作成
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','yatin_kozano','家賃入金口座情報','家賃入金受付口座No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','yatin_kozaname','家賃入金口座情報','受取人名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','yatin_kozanamesjis','家賃入金口座情報','受取人名Unicode','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','yatin_kozakana','家賃入金口座情報','受取人カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','sqsaki_kbn','家賃入金口座情報','請求先区分','0','900','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','jisya_no','家賃入金口座情報','自社支店No','1','999999999','','','','自社_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','jisya_kozano','家賃入金口座情報','自社口座No','1','999','','','','自社口座_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','ow_no','家賃入金口座情報','家主No','1','999999999','','','','家主_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','ow_kozano','家賃入金口座情報','家主口座No','1','999','','','','家主口座_口座用',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_useflg','家賃入金口座情報','ANSER-SPC有無','0','1','','','','','10の設定箇所が不明の為V7の最大最小値を設けておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_kozasiteikbn','家賃入金口座情報','ANSER-SPC口座指定方式','1','2','','','','','10の設定箇所が不明の為V7の最大最小値を設けておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_syumokukbn','家賃入金口座情報','ANSER-SPC種目付加方法','1','9','','','','','10の設定箇所が不明の為V7の最大最小値を設けておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_ninibango','家賃入金口座情報','ANSER-SPC任意番号','0','99','','','','','10の設定箇所が不明の為V7の最大最小値を設けておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_kanyubango','家賃入金口座情報','ANSER-SPC加入者番号','','12','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_kobetuflg','家賃入金口座情報','ANSER-SPC接続先個別設定フラグ','0','1','','','','','10の設定箇所が不明の為V7の最大最小値を設けておく。V7では0がチェックON'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_kobetukbn','家賃入金口座情報','ANSER-SPC接続方法','1','2','','','','','10の設定箇所が不明の為V7の最大最小値を設けておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_kobetuareano','家賃入金口座情報','ANSER-SPCエリアNo','1','99','','','','','10DBから取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_kobetutikuno','家賃入金口座情報','ANSER-SPC地区No','1','99','','','','','10DBから取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_kobetutel','家賃入金口座情報','ANSER-SPCTEL','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_ansyobango','家賃入金口座情報','ANSER-SPC照会用暗証番号','','4','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','spc_servicecode','家賃入金口座情報','ANSER-SPCサービスコード','','4','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','history','家賃入金口座情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','useflg','家賃入金口座情報','使用フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','yatin_kozabiko','家賃入金口座情報','家賃入金口座備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','headertemplate','家賃入金口座情報','ヘッダーテンプレート','','500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_yatinkoza','trallertemplate','家賃入金口座情報','トレーラーテンプレート','','500','','','','',''); ");
                        break;
                    }
                case "自社情報-ANSERエリア情報": // 20160704 ANSER情報の移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcarea','spc_areano','ANSERエリア情報','SPCエリアNo','1','99','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcarea','spc_areaname','ANSERエリア情報','SPCエリア名称','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcarea','biko_spcarea','ANSERエリア情報','備考','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcarea','history','ANSERエリア情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "自社情報-ANSERアクセスポイント情報": // 20160704 ANSER情報の移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcaccesspoint','spc_tikuno','ANSERアクセスポイント情報','SPC地区No','1','99','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcaccesspoint','spc_areano','ANSERアクセスポイント情報','SPCエリアNo','1','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcaccesspoint','spc_tikuname','ANSERアクセスポイント情報','SPC地区名称','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcaccesspoint','biko','ANSERアクセスポイント情報','備考','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcaccesspoint','history','ANSERアクセスポイント情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "自社情報-ANSER接続情報": // 20160704 ANSER情報の移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_setuzokuno','ANSER接続情報','接続方法No','1','1','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_name','ANSER接続情報','設定名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_kaisensyu','ANSER接続情報','回線種別(1:電話回線 2:ISDN)','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_devtype','ANSER接続情報','開発タイプ','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_devname','ANSER接続情報','開発名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_setuzoku_hoho','ANSER接続情報','接続方法(1:アクセスポイント 2:TEL)','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_areano','ANSER接続情報','エリアNo','1','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_tikuno','ANSER接続情報','地区No','1','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_tel','ANSER接続情報','電話番号','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_gaisen','ANSER接続情報','外線No','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_retry_kaisu','ANSER接続情報','リトライ回数','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_retry_kankaku','ANSER接続情報','リトライの間隔','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_servicecode','ANSER接続情報','利用するサービスコード','0','0','','','','','設定値不明のため0を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_biko','ANSER接続情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','history','ANSER接続情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_spcsetuzoku','spc_crlf','ANSER接続情報','改行有無','1','1','','','','',''); ");
                        break;
                    }
                case "家主情報-家主基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','ow_no','家主基本情報','家主No','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kojinhojin_flg','家主基本情報','個人法人フラグ','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','ow_name','家主基本情報','家主名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','ow_namesjis','家主基本情報','家主名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','ow_kana','家主基本情報','家主カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','keisyo','家主基本情報','宛名敬称','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','post_code','家主基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr_kenno','家主基本情報','都道府県コード','1','47','','','','都道府県',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr_sino','家主基本情報','市区町村コード','1101','47382','','','','市区町村',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr_cyo','家主基本情報','町地域','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr_cyome','家主基本情報','丁番地','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr_cyomeptn','家主基本情報','丁番地選択','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr_banti','家主基本情報','街区番号地番','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr_etc','家主基本情報','その他','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','toukiaddr1','家主基本情報','登記住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','toukiaddr2','家主基本情報','登記住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','tel1','家主基本情報','TEL1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','tel2','家主基本情報','TEL2','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','fax','家主基本情報','FAX','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','mobiletel1','家主基本情報','携帯１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','mobiletel2','家主基本情報','携帯２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','mail','家主基本情報','メールアドレス','','256','','','','URL',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','mobilemail','家主基本情報','携帯メールアドレス','','256','','','','URL',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','yusentel_kbn','家主基本情報','優先電話設定','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','yusenmail_kbn','家主基本情報','優先メール設定','1','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','birthday','家主基本情報','誕生日・設立日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','nensyu','家主基本情報','年収・年商','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','biko_kihon','家主基本情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','gender','家主基本情報','性別','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','honseki','家主基本情報','本籍地','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_name','家主基本情報','勤務先名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_namesjis','家主基本情報','勤務先名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_kana','家主基本情報','勤務先カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_postcode','家主基本情報','勤務先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_addr1','家主基本情報','勤務先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_addr2','家主基本情報','勤務先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_tel1','家主基本情報','勤務先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_tel2','家主基本情報','勤務先TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_fax','家主基本情報','勤務先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_gyosyu','家主基本情報','勤務先業種','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_busyo','家主基本情報','勤務先部署','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_nyuryokuym','家主基本情報','勤務先情報記入年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kinmu_nyusyaym','家主基本情報','勤務先入社年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','url','家主基本情報','URL','','2500','','','','URL',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','gyosyu','家主基本情報','業種','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','daihyo_name','家主基本情報','代表者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','daihyo_namesjis','家主基本情報','代表者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','daihyo_kana','家主基本情報','代表者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','daihyo_yakusyoku','家主基本情報','代表者役職','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','daihyo_atenaflg','家主基本情報','代表者を宛先に含める','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','tanto_name','家主基本情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','tanto_namesjis','家主基本情報','担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','tanto_kana','家主基本情報','担当者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','tanto_busyo','家主基本情報','担当者部署','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','tanto_atenaflg','家主基本情報','担当者を宛先に含める','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sihonkin','家主基本情報','資本金','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','jugyosu','家主基本情報','従業員数','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','nyuryoku_ym','家主基本情報','記入年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','torihikisaki','家主基本情報','主要取引先','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_name','家主基本情報','連絡先名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_namesjis','家主基本情報','連絡先名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_kana','家主基本情報','連絡先カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_keisyo','家主基本情報','連絡先宛名敬称','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_postcode','家主基本情報','連絡先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_addr1','家主基本情報','連絡先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_addr2','家主基本情報','連絡先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_tel1','家主基本情報','連絡先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_tel2','家主基本情報','連絡先TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_fax','家主基本情報','連絡先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_mobiletel1','家主基本情報','連絡先携帯１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_mobiletel2','家主基本情報','連絡先携帯２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_yusentelkbn','家主基本情報','連絡先優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','renraku_aidagara','家主基本情報','連絡先間柄','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','biko_renraku','家主基本情報','連絡先備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_kbn','家主基本情報','書類送付先区分','1','4','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_name','家主基本情報','送付先名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_namesjis','家主基本情報','送付先名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_kana','家主基本情報','送付先カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_keisyo','家主基本情報','送付先宛名敬称','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_postcode','家主基本情報','送付先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_addr1','家主基本情報','送付先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_addr2','家主基本情報','送付先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_tel1','家主基本情報','送付先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_tel2','家主基本情報','送付先TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','sofu_fax','家主基本情報','送付先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','biko_sofu','家主基本情報','送付先備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','event_tantono','家主基本情報','イベント担当者No','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','useflg','家主基本情報','検索対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','findkeyword','家主基本情報','絞り込みキーワード','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','history','家主基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','rowid','家主基本情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','kaikei_jouhou','家主基本情報','会計情報','','20','','','','','画面上での設定箇所が不明'); ");
                        // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add sta
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr1','家主基本情報','住所1','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata','addr2','家主基本情報','住所2','','100','','','','',''); ");
                        break;
                    }
                // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add end

                case "家主情報-家主口座情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','ow_no','家主口座情報','家主No','1','999999999','','1','','',''); ");
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('owdata_koza','ow_kozano','家主口座情報','口座No','1','999','','2','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','ow_kozano','家主口座情報','口座No','1','999999999','','2','','',''); ");     // 2016.04.25 ユーザーテスト用に上限を999999999へ引き上げ
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','kinyu_no','家主口座情報','金融機関No','1','9999','','','','金融機関_家主口座',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','kinyu_tenno','家主口座情報','金融機関支店No','1','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','koza_syubetu','家主口座情報','口座種別','1','9','1','','','口座種別',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','koza_bango','家主口座情報','口座番号','1','9999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','koza_meigi','家主口座情報','口座名義','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','koza_meigikana','家主口座情報','口座名義カナ','','100','','','','口座名義カナ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','yucyokoza_kigo1','家主口座情報','ゆうちょ口座記号１','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','yucyokoza_kigo2','家主口座情報','ゆうちょ口座記号２','','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','yucyokoza_bango','家主口座情報','ゆうちょ口座番号','','8','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','biko_koza','家主口座情報','口座備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','biko_furikomi','家主口座情報','振込情報備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','sgfirai_kbn','家主口座情報','総合振込区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','sgfirai_no','家主口座情報','振込依頼人No','1','9999','','','','振込依頼人',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','sgfirai_tesufutankbn','家主口座情報','振込手数料負担区分','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','sgfirai_tesukeisankbn','家主口座情報','振込手数料計算区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','sgfirai_tesukotei1gak','家主口座情報','振込手数料固定額1','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','sgfirai_tesukotei2gak','家主口座情報','振込手数料固定額2','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','biko_sgfirai','家主口座情報','総合振込情報備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','history','家主口座情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','rowid','家主口座情報','行ID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_koza','koza_printkbn','家主口座情報','家主向け帳票の表示','1','2','1','','','',''); ");
                        break;
                    }
                case "家主情報-家主イベント情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_event','ow_no','家主イベント情報','家主No','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_event','event_kbn','家主イベント情報','イベント区分','1','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_event','event_cnt','家主イベント情報','イベントNo','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_event','event_ymd','家主イベント情報','イベント日付','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_event','event_data','家主イベント情報','イベントデータ','-1','6','-1','','','',''); ");
                        break;
                    }
                case "家主情報-家主メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_memo','ow_no','家主メモ情報','家主No','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_memo','memo_no','家主メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_memo','memo','家主メモ情報','内容','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('owdata_memo','history','家主メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }

                case "契約者情報-契約者基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kys_no','契約者基本情報','契約者No','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kys_name','契約者基本情報','契約者名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kys_namesjis','契約者基本情報','契約者名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kys_kana','契約者基本情報','契約者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kojinhojin_flg','契約者基本情報','個人法人フラグ','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','keisyo','契約者基本情報','宛名敬称','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','post_code','契約者基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','addr1','契約者基本情報','住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','addr2','契約者基本情報','住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tel1','契約者基本情報','TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tel2','契約者基本情報','TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','fax','契約者基本情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','mobiletel1','契約者基本情報','携帯１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','mobiletel2','契約者基本情報','携帯２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','mail','契約者基本情報','メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','mobilemail','契約者基本情報','携帯メールアドレス','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','yusentel_kbn','契約者基本情報','優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','yusenmail_kbn','契約者基本情報','優先設定(メールアドレス)','1','11','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','birthday','契約者基本情報','生年月日・設立日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nensyu','契約者基本情報','年収・年商','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','biko_kihon','契約者基本情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','gender','契約者基本情報','性別','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','honseki','契約者基本情報','本籍地','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_name','契約者基本情報','勤務先名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_namesjis','契約者基本情報','勤務先名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_kana','契約者基本情報','勤務先カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_postcode','契約者基本情報','勤務先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_addr1','契約者基本情報','勤務先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_addr2','契約者基本情報','勤務先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_tel1','契約者基本情報','勤務先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_tel2','契約者基本情報','勤務先TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_fax','契約者基本情報','勤務先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_gyosyu','契約者基本情報','勤務先業種','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_busyo','契約者基本情報','勤務先部署','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_nyuryokuym','契約者基本情報','勤務先情報記入年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','kinmu_nyusyaym','契約者基本情報','勤務先入社年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyukyomae_nyuryokuym','契約者基本情報','入居前連絡先情報記入年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyukyomae_postcode','契約者基本情報','入居前連絡先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyukyomae_addr1','契約者基本情報','入居前連絡先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyukyomae_addr2','契約者基本情報','入居前連絡先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyukyomae_tel1','契約者基本情報','入居前連絡先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyukyomae_tel2','契約者基本情報','入居前連絡先TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyukyomae_fax','契約者基本情報','入居前連絡先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','url','契約者基本情報','Webアドレス(URL)','','2500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','gyosyu','契約者基本情報','業種','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','daihyo_name','契約者基本情報','代表者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','daihyo_namesjis','契約者基本情報','代表者名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','daihyo_kana','契約者基本情報','代表者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','daihyo_yakusyoku','契約者基本情報','代表者役職','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','daihyo_atenaflg','契約者基本情報','代表者を宛先に含める','0','1','0','','','','担当者を宛先～ONでOFFになる(ラジオボタンのような動き)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tanto_name','契約者基本情報','担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tanto_namesjis','契約者基本情報','担当者名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tanto_kana','契約者基本情報','担当者カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tanto_busyo','契約者基本情報','担当者部署','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tanto_yakusyoku','契約者基本情報','担当者役職','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','tanto_atenaflg','契約者基本情報','担当者を宛先に含める','0','1','0','','','','代表者を宛先～ONでOFFになる(ラジオボタンのような動き)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sihonkin','契約者基本情報','資本金','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','jugyosu','契約者基本情報','従業員数','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','nyuryoku_ym','契約者基本情報','記入年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','torihikisaki','契約者基本情報','主要取引先','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_name','契約者基本情報','連絡先名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_namesjis','契約者基本情報','連絡先名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_kana','契約者基本情報','連絡先カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_keisyo','契約者基本情報','連絡先宛名敬称','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_postcode','契約者基本情報','連絡先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_addr1','契約者基本情報','連絡先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_addr2','契約者基本情報','連絡先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_tel1','契約者基本情報','連絡先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_tel2','契約者基本情報','連絡先TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_fax','契約者基本情報','連絡先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_mobiletel1','契約者基本情報','連絡先携帯１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_mobiletel2','契約者基本情報','連絡先携帯２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_yusentelkbn','契約者基本情報','連絡先優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','renraku_aidagara','契約者基本情報','連絡先間柄','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','biko_renraku','契約者基本情報','連絡先備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_kbn','契約者基本情報','書類送付先区分','1','6','6','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_name','契約者基本情報','書類送付先名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_namesjis','契約者基本情報','書類送付先名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_kana','契約者基本情報','書類送付先カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_keisyo','契約者基本情報','書類送付先宛名敬称','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_postcode','契約者基本情報','書類送付先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_addr1','契約者基本情報','書類送付先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_addr2','契約者基本情報','書類送付先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_tel1','契約者基本情報','書類送付先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_tel2','契約者基本情報','書類送付先TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_fax','契約者基本情報','書類送付先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','biko_sofu','契約者基本情報','書類送付先備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_hakkofurikomi','契約者基本情報','振込通知書発行の可否','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_hakkofurikae','契約者基本情報','振替通知書発行の可否','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','sofu_hakkotokusoku','契約者基本情報','督促状発行の可否','1','2','1','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 契約者基本情報 -del sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kysdata','furikomituti_bktaniflg','契約者基本情報','物件単位振込通知フラグ','0','1','0','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kysdata','furikaetuti_bktaniflg','契約者基本情報','物件単位振替通知フラグ','0','1','0','','','',''); ")
                        // 20160519 EXEUpdateに伴う修正 契約者基本情報 -del end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','furikomi_kasouseflg','契約者基本情報','口座振込仮想口座使用フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','hosyo_multipleflg','契約者基本情報','保証人複数フラグ','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','furikae_hosyoumu','契約者基本情報','口座振替の保証有無','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','useflg','契約者基本情報','検索対象フラグ','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','history','契約者基本情報','履歴情報','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','rowid','契約者基本情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata','cvpay_umu','契約者基本情報','コンビニ収納サービス利用有無','0','1','1','','','',''); ");
                        break;
                    }
                case "契約者情報-契約者口座情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','kys_no','契約者口座情報','契約者No','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','kys_kozano','契約者口座情報','契約者口座No','1','3','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','koza_kbn','契約者口座情報','口座区分','1','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','kinyu_no','契約者口座情報','金融期間No','1','9999','','','','金融機関',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','kinyu_tenno','契約者口座情報','金融機関支店No','1','9999','','','','金融機関支店',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','koza_syubetu','契約者口座情報','口座種別','1','9','1','','','口座種別',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','koza_bango','契約者口座情報','口座番号','0','9999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','koza_meigi','契約者口座情報','口座名義','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','koza_meigikana','契約者口座情報','口座名義カナ','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','yucyokoza_kigo1','契約者口座情報','ゆうちょ記号１','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','yucyokoza_kigo2','契約者口座情報','ゆうちょ記号２','','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','yucyokoza_bango','契約者口座情報','ゆうちょ口座番号','','8','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','fkae_no','契約者口座情報','口座振替No','1','8999','','','','口座振替',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','fkae_tesugak','契約者口座情報','口座振替手数料','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','fkae_kysbango','契約者口座情報','口座振替契約者番号','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','biko_koza','契約者口座情報','口座備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','biko_furikomi','契約者口座情報','口座振込備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','sgfirai_kbn','契約者口座情報','総合振込区分','1','2','1','','','','口座No = 3 は NULL'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','sgfirai_no','契約者口座情報','振込依頼人No','1','9999','','','','振込依頼人',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','sgfirai_tesufutankbn','契約者口座情報','振込手数料負担区分','1','2','2','','','','口座No = 3 は NULL'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','sgfirai_tesukeisankbn','契約者口座情報','振込手数料計算区分','1','2','1','','','','口座No = 3 は NULL'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','sgfirai_tesukotei1gak','契約者口座情報','振込手数料固定額1','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','sgfirai_tesukotei2gak','契約者口座情報','振込手数料固定額2','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','biko_sgfirai','契約者口座情報','備考(総合振込情報)','','100','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 契約者口座情報 -add
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','koza_printkbn','契約者口座情報','印刷口座区分','1','2','1','','','',''); ");
                        // 20161012 革命10アップデートに伴う修正 -add
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_koza','koza_yucyoflg','契約者口座情報','金融機関区分(ゆうちょ銀行フラグ)','0','1','0','','','',''); ");
                        break;
                    }
                case "契約者情報-契約者照合用カナ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkomsyogo','sqsaki_kbn','契約者照合用カナ情報','マスター区分','0','900','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkomsyogo','sqsaki_no','契約者照合用カナ情報','契約者No','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkomsyogo','sqsaki_kozano','契約者照合用カナ情報','口座No','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkomsyogo','syogo_no','契約者照合用カナ情報','照合用文字列連番','1','10','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkomsyogo','syogo_priority','契約者照合用カナ情報','照合優先順位（1～）','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_fb_fkomsyogo','fkom_syogomoji','契約者照合用カナ情報','振込照合用文字列','','50','','','●','口座名義カナ',''); ");
                        break;
                    }
                case "契約者情報-契約者保証人情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kys_no','契約者保証人情報','契約者No','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','hosyo_no','契約者保証人情報','保証人No','1','5','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','hosyo_name','契約者保証人情報','保証人名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','hosyo_namesjis','契約者保証人情報','保証人名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','hosyo_kana','契約者保証人情報','保証人カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','keisyo','契約者保証人情報','宛名敬称','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','post_code','契約者保証人情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','addr1','契約者保証人情報','住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','addr2','契約者保証人情報','住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','tel1','契約者保証人情報','TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','tel2','契約者保証人情報','TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','fax','契約者保証人情報','FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','mobiletel1','契約者保証人情報','携帯１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','mobiletel2','契約者保証人情報','携帯２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','yusentel_kbn','契約者保証人情報','優先設定(電話番号)','1','12','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','birthday','契約者保証人情報','誕生日・設立日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','nensyu','契約者保証人情報','年収・年商','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','biko_hosyo','契約者保証人情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','aidagara','契約者保証人情報','間柄','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_name','契約者保証人情報','勤務先名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_namesjis','契約者保証人情報','勤務先名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_kana','契約者保証人情報','勤務先カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_postcode','契約者保証人情報','勤務先郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_addr1','契約者保証人情報','勤務先住所１','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_addr2','契約者保証人情報','勤務先住所２','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_tel1','契約者保証人情報','勤務先TEL１','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_tel2','契約者保証人情報','勤務先TEL２','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_fax','契約者保証人情報','勤務先FAX','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_gyosyu','契約者保証人情報','勤務先業種','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_busyo','契約者保証人情報','勤務先部署','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_nyuryokuym','契約者保証人情報','勤務先情報記入年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_nyusyaym','契約者保証人情報','勤務先入社年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','kinmu_taisyaym','契約者保証人情報','勤務先退社年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_hosyo','biko_kinmu','契約者保証人情報','勤務先備考','','100','','','','',''); ");
                        break;
                    }
                case "契約者情報-契約者メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_memo','kys_no','契約者メモ情報','契約者No','1','999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_memo','memo_no','契約者メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_memo','memo','契約者メモ情報','内容','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kysdata_memo','history','契約者メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }

                case "物件情報-物件基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_guid','物件基本情報','物件ユニークID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_no','物件基本情報','物件NO','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_deleteflg','物件基本情報','物件削除フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','delete_guid','物件基本情報','削除Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','delete_day','物件基本情報','削除日','1900/1/1','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','delete_cnt','物件基本情報','削除復旧回数','0','100','','','','','使用しない項目なのでとりあえず100を最大としておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_name','物件基本情報','物件名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_deletename','物件基本情報','物件名(削除時の退避用)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_kana','物件基本情報','物件カナ','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','tatemono_sikibetu','物件基本情報','建物識別コード','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','post_code','物件基本情報','郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','addr_kenno','物件基本情報','都道府県コード','1','47','','','','都道府県',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','addr_sino','物件基本情報','市区町村コード','1101','47382','','','','市区町村',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','addr_cyo','物件基本情報','町地域','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','addr_cyome','物件基本情報','丁番地','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','addr_cyomeptn','物件基本情報','丁番地選択','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','addr_banti','物件基本情報','街区番号地番','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','addr_etc','物件基本情報','その他','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','history','物件基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','rowid','物件基本情報','行ID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_namesjis','物件基本情報','物件名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','bk_gaibuno','物件基本情報','物件外部No','1','999999999','','','','','物件Noと同じ範囲にしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','delete_cause','物件基本情報','削除理由','','-1','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 物件基本情報 -add
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata','lastupdate','物件基本情報','最終更新日','1900/1/1','2100/12/31','','','','',''); ");
                        break;
                    }
                case "物件情報-物件詳細情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','bk_guid','物件詳細情報','物件ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','bk_ruinokbn','物件詳細情報','物件分類','1','9999','1','','●','物件分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gps_wgsido','物件詳細情報','緯度','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gps_wgskeido','物件詳細情報','経度','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gps_tokyoido','物件詳細情報','緯度(日本測地系)','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gps_tokyokeido','物件詳細情報','経度(日本測地系)','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','bk_dentetuflg','物件詳細情報','電鉄物件フラグ','0','0','0','','','','フラグの意味が不明。設定箇所も不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','syunko_ymd','物件詳細情報','竣工日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kaidate','物件詳細情報','地上階建て','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','tika','物件詳細情報','地下階','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','elevator_flg','物件詳細情報','エレベータフラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','elevator_number','物件詳細情報','エレベータ数','1','99','','','','','エレベーター有の場合に設定可'); ");
                        // 20160829 革命10バージョンアップに伴う修正 -del
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('bkdata_detail','rooftop_flg','物件詳細情報','屋上フラグ','0','1','0','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kozo_nokbn','物件詳細情報','建物構造(基本)-構造','-1','99','-1','','','構造','値が存在するのは11まで。99はその他。V7からは名称を抽出し、紐付け結果からコードを取得する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kozo_yanekbn','物件詳細情報','建物構造(基本)-屋根構造','-1','10','-1','','','','DBにマスタが存在する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','moto_gy_fudono','物件詳細情報','情報元業者NO','1','999999','','','','仲介業者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','biko_kihon','物件詳細情報','備考(基本情報)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kosu_total','物件詳細情報','総戸数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_nobeyuka','物件詳細情報','物件面積-延床面積','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_nobeyukatubo','物件詳細情報','物件面積-延床面積坪数','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_sikiti','物件詳細情報','物件面積-敷地面積','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_sikititubo','物件詳細情報','物件面積-敷地面積坪数','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_parking','物件詳細情報','物件面積-駐車面積','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_parkingtubo','物件詳細情報','物件面積-駐車面積坪数','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_yukatoki','物件詳細情報','物件面積-延床面積(登記)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_yukatokitubo','物件詳細情報','物件面積-延床面積坪数(登記)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_sikititoki','物件詳細情報','物件面積-敷地面積(公簿)','0','999999','','','','','設定箇所が不明のためとりあえず他の同一項目と同じ値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_sikititokitubo','物件詳細情報','物件面積-敷地面積坪数(公簿)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','default_sqtuki','物件詳細情報','入金口座-請求月初期値','1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','keiyakuyou_sime','物件詳細情報','契約書用入金締切日','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','yatin_jisansaki','物件詳細情報','家賃持参先','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','yatin_kozano','物件詳細情報','入金口座-家賃入金口座No','1','999999','','','','家賃入金口座',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','yatin_kykozaflg','物件詳細情報','入金口座-契約金用入金口座の有無','1','1','','','','','共通口座に設定するとNULLになる。→「1」以外の値をNULLにするため最小値を「1」にしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','yatin_kykozano','物件詳細情報','入金口座-契約金用入金口座No','1','999999','','','','家賃入金口座',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kozo_taikakbn','物件詳細情報','耐火構造区分','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gy_sekono','物件詳細情報','施工会社No','1','999999','','','','施工業者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gy_hosyuno','物件詳細情報','保守業者No','1','999999','','','','保守業者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanri_hosiki','物件詳細情報','管理形態-管理方式','-1','3','-1','','','','4：無が存在するが登録すると3：自主管理になるため3を最大にしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanri_gyname','物件詳細情報','管理形態-管理業者名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanri_gytanto','物件詳細情報','管理形態-管理業者担当者','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanri_gytel','物件詳細情報','管理形態-管理業者電話番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanrinin_gyomukeitai','物件詳細情報','管理形態-管理業者業務形態','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanrinin_tel','物件詳細情報','管理形態-管理人電話番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','jisya_no','物件詳細情報','支店No','1','999999999','','','','自社',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kotu_sonota1','物件詳細情報','その他交通１-その他交通','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kotu_sonota1kyori','物件詳細情報','その他交通１-距離','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kotu_sonota2','物件詳細情報','その他交通２-その他交通','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kotu_sonota2kyori','物件詳細情報','その他交通２-距離','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','denki_gy_lifeno','物件詳細情報','ライフライン-電気-公共機関No','1','999999','','','','ライフライン(電気)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','water_gy_lifeno','物件詳細情報','ライフライン-上水-公共機関No','1','999999','','','','ライフライン(上水)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gas_gy_lifeno','物件詳細情報','ライフライン-ガス-公共機関No','1','999999','','','','ライフライン(ガス)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','haisui_gy_lifeno','物件詳細情報','ライフライン-排水-公共機関No','1','999999','','','','ライフライン(排水)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','toyu_gy_lifeno','物件詳細情報','ライフライン-灯油-公共機関No','1','999999','','','','ライフライン(灯油)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','lifeline1_gy_lineno','物件詳細情報','ライフライン-その他１-公共機関No','1','999999','','','','ライフライン(その他1)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','lifeline2_gy_lineno','物件詳細情報','ライフライン-その他２-公共機関No','1','999999','','','','ライフライン(その他2)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','lifeline3_gy_lineno','物件詳細情報','ライフライン-その他３-公共機関No','1','999999','','','','ライフライン(その他3)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kensingyomu_umu','物件詳細情報','検針業務の有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kensinorder_kbn','物件詳細情報','検針登録の並び順','1','4','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parenthendo_useflg','物件詳細情報','親子メーター変動費の使用有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_kanriflg','物件詳細情報','駐車場-付随駐車場有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_car','物件詳細情報','駐車場-自動車台数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_bike','物件詳細情報','駐車場-バイク台数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_bicycle','物件詳細情報','駐車場-自転車台数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_bicyclefreeflg','物件詳細情報','駐車場-自転車利用自由フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kubunsyo','物件詳細情報','所有者-一棟・所有区分','1','2','1','','●','',''); ");
                        // 20160519 EXEUpdateに伴う修正 物件詳細情報 -del
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('bkdata_detail','soymd_basis','物件詳細情報','該当月/送金月ベース決定','0','1','','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_kirokukbn','物件詳細情報','石綿使用調査-調査の有無','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_syokai1flg','物件詳細情報','石綿使用調査-調査結果の問合せ先-所有者','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_syokai2flg','物件詳細情報','石綿使用調査-調査結果の問合せ先-管理組合','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_syokai3flg','物件詳細情報','石綿使用調査-調査結果の問合せ先-管理業者','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_syokai4flg','物件詳細情報','石綿使用調査-調査結果の問合せ先-施工業者','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_gy_sekono','物件詳細情報','石綿使用調査-調査結果の問合せ先-施工業者名','0','0','','','','','画面上で設定する箇所が不明。必要？'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_cyosaymd','物件詳細情報','石綿使用調査-調査年月日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_cyosakikankbn','物件詳細情報','石綿使用調査-実施機関','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_cyosahani','物件詳細情報','石綿使用調査-調査範囲','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','ishiwata_useflg','物件詳細情報','石綿使用調査-使用有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_usearea','物件詳細情報','石綿使用調査-使用箇所','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','isiwata_biko','物件詳細情報','石綿使用調査-備考','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_sindanflg','物件詳細情報','耐震診断-診断有無','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_syokai1flg','物件詳細情報','耐震診断-診断記録の問合せ先-所有者','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_syokai2flg','物件詳細情報','耐震診断-診断記録の問合せ先-管理組合','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_syokai3flg','物件詳細情報','耐震診断-診断記録の問合せ先-管理業者','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_syorui1flg','物件詳細情報','耐震診断-耐震基準適合証明書の写し','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_syorui2flg','物件詳細情報','耐震診断-住宅性能評価所の写し','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_syorui3flg','物件詳細情報','耐震診断-耐震診断結果の写し','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','taisin_biko','物件詳細情報','耐震診断-備考','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_dosyatiiki','物件詳細情報','法令-土砂災害警戒地域内外','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_ruikbn','物件詳細情報','法令-法令分類','-1','4','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_naiyo','物件詳細情報','法令-法令内容','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','sikiti_riyoruikbn','物件詳細情報','敷地利用-敷地利用種類','-1','4','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','sikiti_kystartymd','物件詳細情報','敷地利用-契約期間開始','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','sikiti_kyendymd','物件詳細情報','敷地利用-契約期間終了','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','sikiti_biko','物件詳細情報','敷地利用-備考','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','nyukyoritu_startymd','物件詳細情報','入居率一覧対象開始日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','history','物件詳細情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','hyothergyfudo_flg','物件詳細情報','部屋毎に業者が異なる場合フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','gomi_hosoku','物件詳細情報','ゴミ出しに関する補足情報','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','toki_ymd','物件詳細情報','登記情報の日付','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','syo_kenriflg','物件詳細情報','所有権にかかる権利有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','syo_kenrikbn','物件詳細情報','所有権にかかる権利の種類','-1','7','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','other_kenriflg','物件詳細情報','所有権以外の権利有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_kentiku','物件詳細情報','物件面積-建築面積','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_kentikutubo','物件詳細情報','物件面積-建築面積坪数','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_kentikutoki','物件詳細情報','物件面積-建築面積(登記)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','men_kentikutokitubo','物件詳細情報','物件面積-建築面積坪数(登記)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','bk_logonuser_no','物件詳細情報','自社担当者No','1','99999','','','','担当者','最大桁数設定箇所に記載が無かったため画面上の最大値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanri_gyfax','物件詳細情報','管理形態-管理業者FAX','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','homeelevator_flg','物件詳細情報','エレベータフラグ(家)','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_dosyatokubetutiiki','物件詳細情報','法令-土砂災害特別警戒区域内外','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_zoseitakutitiiki','物件詳細情報','法令-造成宅地防災区域内外','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_tunamitiiki','物件詳細情報','法令-津波災害警戒区域内外','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_dosyatiikibiko','物件詳細情報','法令-土砂災害警戒地域備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_dosyatokubetutiikibiko','物件詳細情報','法令-土砂災害特別警戒区域備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_zoseitakutitiikibiko','物件詳細情報','法令-造成宅地防災区域備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','horei_tunamitiikibiko','物件詳細情報','法令-津波災害警戒区域備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','svbunrui_no','物件詳細情報','サービス分類','-1','0','-1','','','','京王カスタマイズ'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','krbunrui_no','物件詳細情報','会計グループ分類','-1','0','-1','','','','画面上で設定する箇所が不明。必要？'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanrinin_name','物件詳細情報','管理形態-管理人名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kozo_other','物件詳細情報','その他建物構造','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kadoti_flg','物件詳細情報','角地フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','cityplan','物件詳細情報','都市計画','-1','7','-1','','','都市計画',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','yototiki','物件詳細情報','用途地域','-1','14','-1','','','用途地域',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanrinin_namesjis','物件詳細情報','管理形態-管理人名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanri_gynamesjis','物件詳細情報','管理形態-管理業者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','kanri_gytantosjis','物件詳細情報','管理形態-管理業者担当者名(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_caraki','物件詳細情報','駐車場空き台数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_bikeaki','物件詳細情報','バイク置き場空き台数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','parking_bicycleaki','物件詳細情報','駐輪場空き台数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','syogaku_name','物件詳細情報','小学校区','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','syogaku_kyori','物件詳細情報','小学校距離','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','cyugaku_name','物件詳細情報','中学校区','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','cyugaku_kyori','物件詳細情報','中学校距離','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','bk_area_no','物件詳細情報','エリア','1','20','','','','エリア','名称を抽出し、あらかじめ作成されているエリアマスタと結合してコードを取得する'); ");
                        // 20160519 EXEUpdateに伴う修正 物件詳細情報 -add
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_detail','emergencyelevator_flg','物件詳細情報','非常用エレベーターフラグ','0','1','0','','','',''); ");
                        break;
                    }
                case "物件情報-物件所有者情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','bk_guid','物件所有者情報','物件ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','kn_no','物件所有者情報','管理No','1','10','','2','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','sorule_guid','物件所有者情報','送金ルールGuid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','kasi1_ow_no','物件所有者情報','貸主１No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','syo1_ow_no','物件所有者情報','所有者１No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','kasi2_ow_no','物件所有者情報','貸主２No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','syo2_ow_no','物件所有者情報','所有者２No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','syo_startymd','物件所有者情報','所有期間開始','1980/01/01','2100/12/31','1980/01/01','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','syo_endymd','物件所有者情報','所有期間終了','1980/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syo','history','物件所有者情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件情報-物件ゴミ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_dust','bk_guid','物件ゴミ情報','物件GUID(Key)','','16','','1','','','中間ファイルでは物件Noのフィールドのため使用する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_dust','dust_no','物件ゴミ情報','ゴミ情報No','1','20','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_dust','dust_rui','物件ゴミ情報','ゴミ分類','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_dust','dust_youbi','物件ゴミ情報','ゴミ出し曜日','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_dust','history','物件ゴミ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件情報-物件権利情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kenri','bk_guid','物件権利情報','物件GUID(Key)','','16','','1','','','中間ファイルでは物件Noのフィールドのため使用する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kenri','kenri_no','物件権利情報','権利情報No','1','5','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kenri','other_kenrirui','物件権利情報','所有権以外の権利の種類','1','6','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kenri','other_kenribiko','物件権利情報','所有権以外の権利備考','','500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kenri','history','物件権利情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件情報-物件交通情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bk_guid','物件交通情報','物件ユニークID','','16','','1','','','中間ファイルでは物件Noのフィールドのため使用する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','ensen_cnt','物件交通情報','沿線番号','1','100','','2','','','最大値をとりあえず100にしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','ensen_no','物件交通情報','沿線No','1','636','','','','沿線',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','eki_no','物件交通情報','駅No','1001','636002','','','','駅',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','firsttrain_flg','物件交通情報','始発フラグ','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','kyori','物件交通情報','距離','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','toho_min','物件交通情報','徒歩','0','999999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','car_min','物件交通情報','車','0','999999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_min','物件交通情報','バス','0','999999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_company','物件交通情報','バス会社','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_station','物件交通情報','バス停','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_tohomin','物件交通情報','バス停徒歩','0','999999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_kyori','物件交通情報','バス停までの距離','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','history','物件交通情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','setting_kotutype','物件交通情報','設定する交通のタイプ','0','1','1','','','','0:最寄りバス停　1：最寄り駅'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_companytoeki','物件交通情報','バス系統・路線名(駅までバスを利用)','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_stationtoeki','物件交通情報','バス停名(駅までバスを利用)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_kyoritoeki','物件交通情報','バス停(駅までバスを利用)までの道路距離','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','bus_tohomintoeki','物件交通情報','バス停(駅までバスを利用)までの徒歩(分)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kotu','car_kyori','物件交通情報','車距離','','20','','','','',''); ");
                        break;
                    }
                // 20160519 EXEUpdateに伴う修正 物件交通情報 -del
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('bkdata_kotu','kotu_syudan','物件交通情報','交通手段','1','3','1','','','',''); ")
                case "物件情報-物件接道情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','bk_guid','物件接道情報','物件Guid','','','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','setudo_cnt','物件接道情報','接道No','1','4','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','setudo_muki','物件接道情報','方角','1','8','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','setudo_roadpattern','物件接道情報','道路種','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','setudo_haba','物件接道情報','幅員','0','999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','setudo_setudokyori','物件接道情報','接道距離','0','999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','setudo_pointselectflg','物件接道情報','位置指定道路','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_setudo','history','物件接道情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件情報-物件周辺情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syuhen','syuhen_guid','物件周辺情報','周辺Guid(Key)','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syuhen','bk_guid','物件周辺情報','物件Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syuhen','syuhen_cnt','物件周辺情報','周辺No','1','10','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syuhen','sisetu_kbn','物件周辺情報','周辺施設分類','1','33','','','●','','中身はハードコーディングされている(随時追加される可能性あり)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syuhen','sisetu_name','物件周辺情報','周辺施設名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syuhen','sisetu_kyori','物件周辺情報','周辺施設距離','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_syuhen','history','物件周辺情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件情報-物件修繕維持管理連絡先情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','bk_guid','物件修繕維持管理連絡先情報','物件GUID(Key)','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_no','物件修繕維持管理連絡先情報','修繕及び維持管理No','1','10','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_kasyo','物件修繕維持管理連絡先情報','修繕及び維持管理の箇所','','100','','','','修繕維持箇所分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_taisyokbn','物件修繕維持管理連絡先情報','対象区分','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_gyno','物件修繕維持管理連絡先情報','業者no','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_simei','物件修繕維持管理連絡先情報','氏名(商号または名称)(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_address','物件修繕維持管理連絡先情報','住所(主たる事務所の所在地)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_tel','物件修繕維持管理連絡先情報','連絡先電話番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_jisyano','物件修繕維持管理連絡先情報','自社no','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_kasino','物件修繕維持管理連絡先情報','貸主No','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_syono','物件修繕維持管理連絡先情報','所有者No','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_szeniji','syuzenijikanri_simeiu','物件修繕維持管理連絡先情報','氏名(商号または名称)','','400','','','','',''); ");
                        break;
                    }
                case "物件情報-物件鍵情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kagi','bk_guid','物件鍵情報','物件GUID(Key)','','16','','1','','',''); ");
                        // 20161028 物件/部屋鍵取得方法修正 -chg sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('bkdata_kagi','kagi_no','物件鍵情報','鍵No','1','50','','2','','共用鍵タイトル',''); ")
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kagi','kagi_no','物件鍵情報','鍵No','1','50','','2','','',''); ");
                                    break;
                                }
                        }
                        // 20161028 物件/部屋鍵取得方法修正 -chg end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kagi','honsu','物件鍵情報','鍵本数','1','99','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kagi','biko','物件鍵情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kagi','history','物件鍵情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kagi','hokan','物件鍵情報','保管場所','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_kagi','gyshare','物件鍵情報','業者間での情報共有','','100','','','','',''); ");
                        break;
                    }
                case "物件情報-物件メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_memo','bk_guid','物件メモ情報','物件ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_memo','memo_no','物件メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_memo','memo','物件メモ情報','メモ','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_memo','history','物件メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件情報-物件近隣駐車場情報":   // (汎用ツールのみ)
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_parkingother','bk_guid','物件近隣駐車場情報','物件GUID(Key)','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_parkingother','parking_no','物件近隣駐車場情報','駐車場No','1','10','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_parkingother','naiyo','物件近隣駐車場情報','駐車場名','','200','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_parkingother','kyori','物件近隣駐車場情報','距離(m)','','20','','','','','999999までの数値のみ入力可'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_parkingother','gak','物件近隣駐車場情報','料金(月額)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_parkingother','zeikbn','物件近隣駐車場情報','税区分','1','2','-1','','','',''); ");
                        break;
                    }
                case "物件情報-物件参照ファイル情報":  // (汎用ツールのみ)
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_relfile','bk_guid','物件参照ファイル情報','物件Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_relfile','file_no','物件参照ファイル情報','ファイルNo','0','9','','2','','','行Noは0から始まっている。とりあえず10ファイルにしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_relfile','fullpath','物件参照ファイル情報','フルパス','','-1','','','●','パス',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_relfile','addtime','物件参照ファイル情報','追加時間','1900/1/1','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_relfile','biko','物件参照ファイル情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_relfile','history','物件参照ファイル情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件情報-物件変動費親メーター情報":    // (汎用ツールのみ)
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','bk_guid','物件変動費親メーター情報','物件ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','bkhendo_guid','物件変動費親メーター情報','物件変動Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','rec_no','物件変動費親メーター情報','行No','1','10','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','hendo_kbn','物件変動費親メーター情報','メーター分類 (水道/ガス/電気/灯油/その他)','1','5','-1','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','meter_name','物件変動費親メーター情報','メーター名(部屋no・フロア名など)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','nkin_no','物件変動費親メーター情報','変動費入金項目','5000','5999','','','','入金項目(変動費)',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','child_meter','物件変動費親メーター情報','子メーター検針','1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','biko','物件変動費親メーター情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','useflg','物件変動費親メーター情報','使用フラグ','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','history','物件変動費親メーター情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','sq_mmkbn','物件変動費親メーター情報','請求月','1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','tani','物件変動費親メーター情報','単位','','40','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('bkdata_hendo','tanisjis','物件変動費親メーター情報','単位(SJIS)','','10','','','','',''); ");
                        break;
                    }

                case "部屋情報-部屋基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','hy_guid','部屋基本情報','部屋ユニークID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','bk_guid','部屋基本情報','物件ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','hy_no','部屋基本情報','部屋NO','','16','','2','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','hy_gaibuno','部屋基本情報','外部用部屋No','0','0','','','','','中間ファイル作成時に設定しても移行時にエラーでスキップされる場合があるため、部屋情報作成後に連番を一括で更新する。プログラム内ではダミーで0を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','hy_deleteflg','部屋基本情報','部屋削除フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','delete_guid','部屋基本情報','削除Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','delete_day','部屋基本情報','削除日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','delete_cnt','部屋基本情報','削除復旧回数','0','100','','','','','使用しない項目なのでとりあえず100を最大としておく'); ");
                        // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata','hy_ruinokbn','部屋基本情報','部屋分類','1','9999','1','','','部屋分類',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','hy_ruinokbn','部屋基本情報','部屋分類名','1','9999','10','','','部屋分類',''); ");
                        // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','madori_cnt','部屋基本情報','間取り','1','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','madori_typekbn','部屋基本情報','間取り(LDK種類)','-1','10','-1','','','間取','間取の文字列をプログラム内で変換して取得するため中間ファイルでは使用しない'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','madori_biko','部屋基本情報','間取り備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_senyujitu','部屋基本情報','専有実面積(m2)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_senyujitutubo','部屋基本情報','専有実面積(坪)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_yuka','部屋基本情報','床面積(m2)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_yukatubo','部屋基本情報','床面積(坪)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_senyutouki','部屋基本情報','専有登記面積(m2)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_senyutoukitubo','部屋基本情報','専有登記面積(坪)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_toki','部屋基本情報','登記延床面積(m2)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_tokitubo','部屋基本情報','登記延床面積(坪)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_balcony','部屋基本情報','バルコニー面積(m2)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_balconytubo','部屋基本情報','バルコニー面積(坪)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_tempo','部屋基本情報','店舗付き住宅店舗部分面積(m2)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_tempotubo','部屋基本情報','店舗付き住宅店舗部分面積(坪)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_jutaku','部屋基本情報','店舗付き住宅住宅部分面積(m2)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','men_jutakutubo','部屋基本情報','店舗付き住宅住宅部分面積(坪)','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','history','部屋基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','rowid','部屋基本情報','行ID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','delete_cause','部屋基本情報','削除理由','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','kaiyaku_uketukekbn','部屋基本情報','解約受付区分','1','3','1','','','',''); ");
                        // 20161130 部屋情報_解約受付月数のデフォルト値をNULLに変更 -chg sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata','kaiyaku_months','部屋基本情報','解約受付月数','1','99','3','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','kaiyaku_months','部屋基本情報','解約受付月数','1','99','','','','',''); ");
                        // 20161130 部屋情報_解約受付月数のデフォルト値をNULLに変更 -chg end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','kaiyaku_days','部屋基本情報','解約受付日数','1','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','kaiyaku_day','部屋基本情報','解約受付日にち','1','31','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 部屋基本情報 -add sta
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','sort_hy_no','部屋基本情報','ソート用部屋No','','300','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','lastupdate','部屋基本情報','最終更新日','1900/01/01','2100/12/31','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 部屋基本情報 -add end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata','wmp_id','部屋基本情報','自社Web内部キー採番ID','','12','','','','',''); "); // 20161012 革命10アップデートに伴う修正 -add
                        break;
                    }

                case "部屋情報-部屋詳細情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hy_guid','部屋詳細情報','部屋ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hygyomu_setflg','部屋詳細情報','部屋業務期間有無','0','0','','','','','画面上での設定箇所が不明のためとりあえず最大最小を0にしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syozai_kaisu1','部屋詳細情報','所在階１','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syozai_kaisu2','部屋詳細情報','所在階２','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syozai_kaisu3','部屋詳細情報','所在階３','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syozai_tikaflg1','部屋詳細情報','地下フラグ１','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syozai_tikaflg2','部屋詳細情報','地下フラグ２','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syozai_tikaflg3','部屋詳細情報','地下フラグ３','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','mukikbn','部屋詳細情報','向き','-1','8','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','kadoheya','部屋詳細情報','角部屋フラグ','0','1','0','','','',''); ");
                        // 20160829 革命10バージョンアップに伴う修正 -del
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_detail','saikokbn','部屋詳細情報','採光','-1','3','-1','','','',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','balcony_mukikbn','部屋詳細情報','バルコニー方向','-1','8','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_jokyokbn','部屋詳細情報','状況','90','99','','','','','設定値が変更されている。[20151218] 詳細を移行仕様書に追記'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_jokyomemo','部屋詳細情報','入居状況備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_syunflg','部屋詳細情報','入居可能日を時期で指定','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_ym','部屋詳細情報','入居可能時期(年月)','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_syunkbn','部屋詳細情報','入居可能時期(上旬・中旬・下旬)','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_joken','部屋詳細情報','入居条件','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','kakunin_ymd','部屋詳細情報','広告内容確認日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','bosyu_startflg','部屋詳細情報','募集可能種別(しない・する)','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','freerent_flg','部屋詳細情報','フリーレント有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','freerent_month','部屋詳細情報','フリーレントカ月','0','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','freerent_detaill','部屋詳細情報','フリーレント詳細','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hoken_kbn','部屋詳細情報','保険-保険の利用','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hoken_kikan','部屋詳細情報','保険-保険期間','0','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hoken_gak','部屋詳細情報','保険-保険料','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hoken_biko','部屋詳細情報','保険-保険の備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_taiyokbn','部屋詳細情報','取引形態種別','-1','6','-1','','','取引態様','[20151221 ]1：仲介先物が削除されている'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_kyakutuke','部屋詳細情報','客付け状態可否','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_tesumoto','部屋詳細情報','手数料負担割合貸主(%)','0','100','','','','','入力は999まで可能だが割合なので100にする'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_tesukyaku','部屋詳細情報','手数料負担割合借主(%)','0','100','','','','','入力は999まで可能だが割合なので100にする'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_futankasi','部屋詳細情報','手数料配分元付(%)','0','100','','','','','入力は999まで可能だが割合なので100にする'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_futankari','部屋詳細情報','手数料配分先物(%)','0','100','','','','','入力は999まで可能だが割合なので100にする'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_kyakutukecomment','部屋詳細情報','客付会社への物件コメント','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_gykokokukatudokbn','部屋詳細情報','業者間広告広告活動種別','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','moto_gy_fudono','部屋詳細情報','情報元業者No','1','999999','0','','','仲介業者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','koukoku_ryoukbn','部屋詳細情報','広告料有無','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','koukoku_jogengak','部屋詳細情報','広告料上限額','0','999999999999','','','','','画面上から最大値を取得。最大値は単位「円」「ヶ月」共通'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','koukoku_jokennaiyo','部屋詳細情報','広告料条件内容','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','biko','部屋詳細情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','addr_replaceflg','部屋詳細情報','部屋住所の変更フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','addr_replacecyome','部屋詳細情報','部屋住所(丁目)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','addr_replacecyomeptn','部屋詳細情報','部屋住所(丁目区分)','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','addr_replacebanti','部屋詳細情報','部屋住所(町地域)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','addr_replaceetc','部屋詳細情報','部屋住所(その他)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hosyo_gyno','部屋詳細情報','保証会社','1','999999','','','','家賃保証業者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hosyo_naiyo','部屋詳細情報','保証内容','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','kyrui_nokbn','部屋詳細情報','契約の種類','1','9999','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','keiyaku_kikan','部屋詳細情報','定期借家契約(期間)','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','keiyaku_kijitu','部屋詳細情報','定期借家契約(期日)','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','parking_biko','部屋詳細情報','駐車場備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','parking_bikebiko','部屋詳細情報','バイク駐車場備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','parking_cyurinbiko','部屋詳細情報','駐輪場備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','shared_salespoint','部屋詳細情報','共通セールスポイント','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','history','部屋詳細情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','keisai_bkflg','部屋詳細情報','不動産検索サイト掲載設定-物件名','0','1','0','','','','広告情報→インターネット広告掲載指示→送信媒体指示→表示される画面の下部'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','keisai_hyflg','部屋詳細情報','不動産検索サイト掲載設定-部屋NO','0','1','0','','','','広告情報→インターネット広告掲載指示→送信媒体指示→表示される画面の下部'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','keisai_bantiflg','部屋詳細情報','不動産検索サイト掲載設定-丁番地以下','0','1','0','','','','広告情報→インターネット広告掲載指示→送信媒体指示→表示される画面の下部'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','keisai_mapflg','部屋詳細情報','不動産検索サイト掲載設定-地図上','0','1','0','','','','広告情報→インターネット広告掲載指示→送信媒体指示→表示される画面の下部'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','hosyo_kbn','部屋詳細情報','賃貸保証利用区分','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','koukoku_jogengakkbn','部屋詳細情報','広告料上限額区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','koukoku_jogenrit','部屋詳細情報','広告料上限率','0','100','','','','','入力は999まで可能だが割合なので100にする'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','koukoku_jogentaxkbn','部屋詳細情報','広告料上限額税区分','-1','2','-1','','','','画面上のコンボボックスでは表示順が逆になっている'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','commonsalespoint_useflg','部屋詳細情報','共通セールスポイント使用フラグ','0','1','1','','','','画面上での設定箇所が不明のためとりあえず最大最小を0にしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','floors_flg','部屋詳細情報','複数階有りのフラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','jisya_no','部屋詳細情報','支店NO','1','999999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','jisya_tanto','部屋詳細情報','自社担当者No','1','99999','0','','','担当者','最大桁数リストに記載が無かったためマスタ登録画面から取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nextnkin_kosindefault','部屋詳細情報','次回更新時の初期値','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','toki_ymd','部屋詳細情報','登記情報の日付','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syo_kenriflg','部屋詳細情報','所有権にかかる権利有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','syo_kenrikbn','部屋詳細情報','所有権にかかる権利の種類','-1','7','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','other_kenriflg','部屋詳細情報','所有権以外の権利有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','btob_groupkbn','部屋詳細情報','BtoBプラグイングループ設定区分','0','2','0','','','','画面上での設定箇所が不明のためとりあえず最大最小を0にしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','jisyaweb_osusumebk','部屋詳細情報','自社Webオススメ物件表示','0','1','0','','','','広告情報→インターネット広告掲載指示→送信媒体指示→表示される画面の下部'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','kaiyaku_ym','部屋詳細情報','解約日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','taikyo_ym','部屋詳細情報','退去日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','svbunrui_no','部屋詳細情報','サービス分類','0','0','0','','','','カスタマイズ項目の為対象外'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','krbunrui_no','部屋詳細情報','会計グループ分類','0','0','0','','','','使用箇所が不明のため対象外'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','siyo_mokuteki','部屋詳細情報','使用目的','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_sintikukbn','部屋詳細情報','新築区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','nyukyo_jikikbn','部屋詳細情報','入居時期区分','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','torihiki_jisyakbn','部屋詳細情報','取引自社区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_detail','keiyaku_kikankbn','部屋詳細情報','契約期間区分','1','2','1','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋所有者情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','hy_guid','部屋所有者情報','部屋ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','kn_no','部屋所有者情報','管理No','1','10','','2','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','sorule_guid','部屋所有者情報','送金ルールGuid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','kasi1_ow_no','部屋所有者情報','貸主１No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','syo1_ow_no','部屋所有者情報','所有者１No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','kasi2_ow_no','部屋所有者情報','貸主２No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','syo2_ow_no','部屋所有者情報','所有者２No','0','999999999','0','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','syo_startymd','部屋所有者情報','所有期間開始','1980/01/01','2100/12/31','1980/01/01','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','syo_endymd','部屋所有者情報','所有期間終了','1980/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_syo','history','部屋所有者情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋駐車場情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','hy_guid','部屋駐車場情報','部屋ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','parking_kbn','部屋駐車場情報','駐車場区分','1','3','1','2','','','画面上で登録時に3レコード生成される'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','parking_akisu','部屋駐車場情報','駐車場の空き数(手動設定用)','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','parking_status','部屋駐車場情報','駐車場の空き有無(手動設定用)','1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','parking_gakkbn','部屋駐車場情報','駐車場料金区分(手動設定用)','-1','3','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','parking_gak','部屋駐車場情報','駐車場料金(手動設定用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','parking_gakzei','部屋駐車場情報','駐車場料金税区分(手動設定用)','-1','2','-1','','','','画面上のコンボボックスでは表示順が逆になっている'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','history','部屋駐車場情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_parking','parking_tintaisu','部屋駐車場情報','駐車場の賃貸可能数','0','9999','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋特約情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_tokuyaku','hy_guid','部屋特約情報','部屋GUID(Key)','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_tokuyaku','tokuyaku_grpno','部屋特約情報','特約事項区分','1','3','','2','','','画面上で登録時に3レコード生成される'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_tokuyaku','naiyo','部屋特約情報','特約事項内容','','-1','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_tokuyaku','history','部屋特約情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋鍵情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kagi','hy_guid','部屋鍵情報','物件GUID(Key)','','16','','1','','',''); ");
                        // 20161028 物件/部屋鍵取得方法修正 -chg sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','kagi_no','部屋鍵情報','鍵明細No','1','50','','2','','専用鍵タイトル',''); ")
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kagi','kagi_no','部屋鍵情報','鍵明細No','1','50','','2','','',''); ");
                                    break;
                                }
                        }
                        // 20161028 物件/部屋鍵取得方法修正 -chg end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kagi','honsu','部屋鍵情報','鍵本数','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kagi','biko','部屋鍵情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kagi','history','部屋鍵情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kagi','hokan','部屋鍵情報','保管場所','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kagi','gyshare','部屋鍵情報','業者間での情報共有','','100','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋面積情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','hy_guid','部屋面積情報','物件GUID(Key)','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','mensekitype','部屋面積情報','面積タイプ','1','2','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','no','部屋面積情報','面積No','1','10','','3','','','画面上から最大数を取得できないため10に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','mensekikbn','部屋面積情報','面積区分','-1','4','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','name','部屋面積情報','区画名','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','meter','部屋面積情報','面積(m2)','0','999999','','','●','','画面上から最大入力可能桁数を取得しただけなので算出等を考慮した値ではない'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','tubo','部屋面積情報','坪数','0','999999','','','','','画面上から最大入力可能桁数を取得しただけなので算出等を考慮した値ではない'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_othermenseki','history','部屋面積情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋間取内訳情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_madoriutiwake','hy_guid','部屋間取内訳情報','物件GUID(Key)','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_madoriutiwake','no','部屋間取内訳情報','間取り内訳No','1','50','','2','','','画面上から最大値を取得'); ");
                        // 20170524 部屋間取内訳区分の最小値修正対応 -chg sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_madoriutiwake','hykbn','部屋間取内訳情報','間取り内訳区分','1','12','','','','間取内訳',''); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_madoriutiwake','hykbn','部屋間取内訳情報','間取り内訳区分','-1','12','','','','間取内訳',''); ");
                        // 20170524 部屋間取内訳区分の最小値修正対応 -chg end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_madoriutiwake','jo','部屋間取内訳情報','畳数','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_madoriutiwake','history','部屋間取内訳情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_madoriutiwake','syozaikai','部屋間取内訳情報','所在階','0','999','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋修繕維持管理連絡先情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','hy_guid','部屋修繕維持管理連絡先情報','部屋GUID(Key)','','','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_no','部屋修繕維持管理連絡先情報','修繕及び維持管理No','1','10','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_kasyo','部屋修繕維持管理連絡先情報','修繕及び維持管理の箇所','','100','','','','修繕維持箇所分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_taisyokbn','部屋修繕維持管理連絡先情報','対象区分','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_gyno','部屋修繕維持管理連絡先情報','業者no','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_simei','部屋修繕維持管理連絡先情報','氏名(商号または名称)(SJIS)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_address','部屋修繕維持管理連絡先情報','住所(主たる事務所の所在地)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_tel','部屋修繕維持管理連絡先情報','連絡先電話番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_jisyano','部屋修繕維持管理連絡先情報','自社no','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_kasino','部屋修繕維持管理連絡先情報','貸主No','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_syono','部屋修繕維持管理連絡先情報','所有者No','0','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szeniji','syuzenijikanri_simeiu','部屋修繕維持管理連絡先情報','氏名(商号または名称)','','400','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_memo','hy_guid','部屋メモ情報','部屋ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_memo','memo_no','部屋メモ情報','メモNo','1','50','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_memo','memo','部屋メモ情報','メモ','','-1','','','●','','V7からは備考30件 + 重要事項20件の内訳で移行する。中間ファイルはメモ1～50で特に注意点等は記載しない(メモに統合されるため)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_memo','history','部屋メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋設備情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','hy_guid','部屋設備情報','部屋Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','komok_guid','部屋設備情報','賃貸設備・特色項目Guid','','16','','2','','設備','設備の文字列退避用に使用する。紐付け後にguidへ入れ替える'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','override_kbn','部屋設備情報','上書き区分','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','komok_name','部屋設備情報','設備項目名','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','disp1name','部屋設備情報','表示名１','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','disp2name','部屋設備情報','表示名２','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','disp3name','部屋設備情報','表示名３','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','disp1iconguid','部屋設備情報','アイコン１','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','disp2iconguid','部屋設備情報','アイコン２','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','disp3iconguid','部屋設備情報','アイコン３','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_setubilst','history','部屋設備情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋入金項目情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','hy_guid','部屋入金項目情報','物件GUID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','tuki_kbn','部屋入金項目情報','月区分','1','4','','2','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','nkin_no','部屋入金項目情報','入金項目No','1000','3999','','3','●','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','nkin_recno','部屋入金項目情報','入金項目行No','1','50','','','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','nkin_kbn','部屋入金項目情報','入金項目区分','0','11','0','','','','画面上で登録操作を行ったが全て0が設定される'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sq_gak','部屋入金項目情報','請求額','0','99999999999','','','','','0円での登録可'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sq_zeikbn','部屋入金項目情報','税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','calc_kbn','部屋入金項目情報','算出区分','1','2','','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','calc_nkinno','部屋入金項目情報','算出基準入金項目No','1000','1999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','calc_monthcnt','部屋入金項目情報','算出ヶ月','0','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sqsaki_no','部屋入金項目情報','請求先No','0','0','','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sq_mmkbn','部屋入金項目情報','請求月区分','1','5','2','','','','画面上で未設定にすることはできなかった'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sqstart_ymd','部屋入金項目情報','請求開始月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sq_ptn','部屋入金項目情報','請求パターン','1','2','','','','請求パターン','入金項目設定画面の各レコード最右参照ボタン内'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sq_interval','部屋入金項目情報','固定公共料金で使用','1','3','','','','請求間隔',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sq_nen','部屋入金項目情報','請求発生年','','100','','','','請求発生年',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','sq_tuki','部屋入金項目情報','請求発生月','','100','','','','請求発生月',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','biko','部屋入金項目情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_nkin','history','部屋入金項目情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋変動費各戸メーター情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','hy_guid','部屋変動費各戸メーター情報','部屋ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','hyhendo_guid','部屋変動費各戸メーター情報','部屋変動費Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','rec_no','部屋変動費各戸メーター情報','行No','1','50','','2','','','最大数が不明のためとりあえず50を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','hendo_kbn','部屋変動費各戸メーター情報','メーター分類','1','5','','','●','メーター分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','meter_name','部屋変動費各戸メーター情報','メーター名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','nkin_no','部屋変動費各戸メーター情報','変動費入金項目','5000','5999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','biko','部屋変動費各戸メーター情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','hendorule_no','部屋変動費各戸メーター情報','変動費請求ルールNo','1','9999','','','','変動費マスタ','最大桁数が不明なためマスタ作成画面から取得'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','hendorule_biko','部屋変動費各戸メーター情報','変動ルール備考','','100','','','','','使用箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','history','部屋変動費各戸メーター情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','sq_mmkbn','部屋変動費各戸メーター情報','請求月区分','1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_hendo','useflg','部屋変動費各戸メーター情報','使用有無','0','1','1','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋共通セールスポイント情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_commonsalespointparts','hy_guid','部屋共通セールスポイント情報','部屋ユニークID','','','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_commonsalespointparts','parts_no','部屋共通セールスポイント情報','パーツNO','1','6','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_commonsalespointparts','salespoint_parts','部屋共通セールスポイント情報','セールスポイントパーツ','','-1','','','●','',''); ");
                        break;
                    }
                case "部屋情報-部屋契約解約確認事項情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_confirm','hy_guid','部屋契約解約確認事項情報','部屋Guid','','','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_confirm','confirm_kbn','部屋契約解約確認事項情報','確認事項区分','1','6','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_confirm','confirm_no','部屋契約解約確認事項情報','確認事項No','1','20','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_confirm','naiyo','部屋契約解約確認事項情報','内容','','100','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋権利情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kenri','hy_guid','部屋権利情報','部屋ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kenri','kenri_no','部屋権利情報','権利情報No','1','5','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kenri','other_kenrirui','部屋権利情報','所有権以外の権利の種類','1','6','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kenri','other_kenribiko','部屋権利情報','所有権以外の権利備考','','500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kenri','history','部屋権利情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋参照ファイル情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_relfile','hy_guid','部屋参照ファイル情報','部屋Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_relfile','file_no','部屋参照ファイル情報','ファイルNo','0','9','','2','','','行Noは0から始まっている。とりあえず10ファイルにしておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_relfile','fullpath','部屋参照ファイル情報','フルパス','','-1','','','●','パス',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_relfile','addtime','部屋参照ファイル情報','追加時間','1900/1/1','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_relfile','biko','部屋参照ファイル情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_relfile','history','部屋参照ファイル情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "部屋情報-部屋原状回復目安単価情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szen','hy_guid','部屋原状回復目安単価情報','部屋Guid','','','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szen','szen_no','部屋原状回復目安単価情報','修繕No','1','20','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szen','szen_name','部屋原状回復目安単価情報','修繕項目名','','100','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szen','ryo','部屋原状回復目安単価情報','数量','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szen','unit_name','部屋原状回復目安単価情報','単位','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_szen','tanka','部屋原状回復目安単価情報','単価','0','99999999999','','','','',''); ");
                        break;
                    }

                case "送金ルール情報-送金ルール基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sorule_guid','送金ルール基本情報','送金ルールユニークID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sorule_no','送金ルール基本情報','送金ルール管理No','1','10','','2','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','relation_guid','送金ルール基本情報','関連性ユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kubunsyo','送金ルール基本情報','一所有形態区分-棟/区分','1','2','1','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sorule_startymd','送金ルール基本情報','送金ルール適用開始日','1980/01/01','2100/12/31','1980/01/01','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sorule_endymd','送金ルール基本情報','送金ルール適用終了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanri_keitaikbn','送金ルール基本情報','管理形態','1','5','1','','','','6：その他の設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','ikkatu_kanriflg','送金ルール基本情報','一括借上 一部管理','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','cyukai_kyflg','送金ルール基本情報','仲介物件 - 新規契約業務','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','cyukai_koflg','送金ルール基本情報','仲介物件 - 契約更新業務','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','cyukai_kaiflg','送金ルール基本情報','仲介物件 - 解約業務','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd1_gaitoukbn','送金ルール基本情報','送金日決定方法1 - 該当年月','-1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd1_simekbn','送金ルール基本情報','送金日決定方法1 - 締日','-1','31','30','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd1_sokintukikbn','送金ルール基本情報','送金日決定方法1 - 送金月','-1','6','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd1_sokinsimekbn','送金ルール基本情報','送金日決定方法1 - 送金締日','-1','31','10','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd2_gaitoukbn','送金ルール基本情報','送金日決定方法2 - 該当年月','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd2_simekbn','送金ルール基本情報','送金日決定方法2 - 締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd2_sokintukikbn','送金ルール基本情報','送金日決定方法2 - 送金月','-1','6','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd2_sokinsimekbn','送金ルール基本情報','送金日決定方法2 - 送金締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd3_gaitoukbn','送金ルール基本情報','送金日決定方法3 - 該当年月','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd3_simekbn','送金ルール基本情報','送金日決定方法3 - 締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd3_sokintukikbn','送金ルール基本情報','送金日決定方法3 - 送金月','-1','6','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd3_sokinsimekbn','送金ルール基本情報','送金日決定方法3 - 送金締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd4_gaitoukbn','送金ルール基本情報','送金日決定方法4 - 該当年月','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd4_simekbn','送金ルール基本情報','送金日決定方法4 - 締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd4_sokintukikbn','送金ルール基本情報','送金日決定方法4 - 送金月','-1','6','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd4_sokinsimekbn','送金ルール基本情報','送金日決定方法4 - 送金締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd5_gaitoukbn','送金ルール基本情報','送金日決定方法5 - 該当年月','-1','5','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd5_simekbn','送金ルール基本情報','送金日決定方法5 - 締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd5_sokintukikbn','送金ルール基本情報','送金日決定方法5 - 送金月','-1','6','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd5_sokinsimekbn','送金ルール基本情報','送金日決定方法5 - 送金締日','-1','31','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sosaki_multikbn','送金ルール基本情報','送金先単独/複数指定','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sosaki_koteiflg','送金ルール基本情報','送金固定額使用フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sosaki_koteisu','送金ルール基本情報','送金固定数','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sosaki_anbunflg','送金ルール基本情報','均等案分フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sosaki_hasuuketorisaki','送金ルール基本情報','端数受取先','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sosaki_hasuadjustmentflg','送金ルール基本情報','送金額案分端数調整フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','ikkatu_kbn','送金ルール基本情報','一括借上 物件毎/部屋毎','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','ikkatu_bkgak','送金ルール基本情報','一括借上 物件毎設定額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','ikkatu_menseki','送金ルール基本情報','一括借上 免責期間の設定フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','ikkatu_mensekimonth','送金ルール基本情報','一括借上 免責期間(解約翌月から何カ月)','0','99','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanritesu_kbn','送金ルール基本情報','管理手数料区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanritesu_cyosyukbn','送金ルール基本情報','管理手数料徴収区分','1','4','1','','','','5以降の設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanri_reigaiky','送金ルール基本情報','管理手数料 例外：契約金は対象外とする','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanri_reigaikai','送金ルール基本情報','管理手数料 例外：解約金は対象外とする','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyteigaku_kbn','送金ルール基本情報','管理手数料定額区分','0','0','','','','','画面上での設定箇所が不明のためとりあえず0を設定しておく(管理手数料区分と内容が似ているので統一された？)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyteigaku_kanrigak','送金ルール基本情報','部屋毎定額 - 全部屋一律管理手数料','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyteigaku_hiwariflg','送金ルール基本情報','部屋毎定額 - 部屋毎日割りフラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','bkteigaku_kanrigak','送金ルール基本情報','物件毎定額 - 物件管理手数料','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanri_taxflg','送金ルール基本情報','管理手数料 消費税適用フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','biko_basic','送金ルール基本情報','備考 - 基本情報','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sh_daihyohyflg','送金ルール基本情報','支払明細書関連 - 同時契約の場合は代表する部屋を表示','0','0','0','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sh_kozabetuflg','送金ルール基本情報','支払明細書関連 - 口座毎に支払明細書を作成1','0','0','0','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','sh_kozabetuListNokbn','送金ルール基本情報','支払明細書関連 - 口座毎に支払明細書を作成2','-1','-1','-1','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyikkatu_nkin_no1kbn','送金ルール基本情報','部屋毎一括借上額 - 入金項目No1','-1','1999','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyikkatu_nkin_no2kbn','送金ルール基本情報','部屋毎一括借上額 - 入金項目No2','-1','1999','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyikkatu_nkin_no3kbn','送金ルール基本情報','部屋毎一括借上額 - 入金項目No3','-1','1999','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyikkatu_nkin_no4kbn','送金ルール基本情報','部屋毎一括借上額 - 入金項目No4','-1','1999','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','hyikkatu_nkin_no5kbn','送金ルール基本情報','部屋毎一括借上額 - 入金項目No5','-1','1999','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanritesu_zeiumukbn','送金ルール基本情報','管理手数料計算基準区分','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','history','送金ルール基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','rowid','送金ルール基本情報','付箋ID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','no_soruleflg','送金ルール基本情報','満額入金-送金フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','cyukai_zuijisokin','送金ルール基本情報','仲介物件 - 随時送金','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','ikkatu_calckbn','送金ルール基本情報','一括借上 部屋毎一括借上額の計算方法','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','ikkatu_taxflg','送金ルール基本情報','一括借上 消費税適用フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','kanri_utizeiflg','送金ルール基本情報','内税フラグ','0','0','0','','','','タケツーカスタマイズ'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule','soymd_basis','送金ルール基本情報','soymd_basis','0','0','0','','','','画面上での設定箇所が不明'); ");
                        break;
                    }
                case "送金ルール情報-送金ルール送金先情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_sosaki','sorule_guid','送金ルール送金先情報','送金ルールGuid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_sosaki','sorule_no','送金ルール送金先情報','送金ルール管理No','1','10','','2','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_sosaki','sosaki_recno','送金ルール送金先情報','送金先行No','1','3','','3','','','送金ルール作成時に自動で3レコード生成される'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_sosaki','so_ow_no','送金ルール送金先情報','送金先家主NO','1','999999999','','','','家主',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_sosaki','so_ow_kozano','送金ルール送金先情報','送金先口座NO','1','999','','','','家主口座',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_sosaki','so_rit','送金ルール送金先情報','送金率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_sosaki','so_fixsogak','送金ルール送金先情報','固定送金額','0','999999999999','','','','',''); ");
                        break;
                    }
                case "送金ルール情報-送金ルール入金項目情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','sorule_guid','送金ルール入金項目情報','送金ルールユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','sorule_no','送金ルール入金項目情報','送金ルール管理No','1','10','','2','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','taisyokbn','送金ルール入金項目情報','月々/契約時/更新時区分','1','7','','3','','','7:その他入金項目が1レコード生成される'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','nkin_sortorder','送金ルール入金項目情報','行NO','0','49','','4','','','最大値が不明のためとりあえず50で設定しておく (行Noは0から始まっている)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','nkin_no','送金ルール入金項目情報','入金項目No','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','sokin_rit','送金ルール入金項目情報','送金率','0','100','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','kanrigak_rit','送金ルール入金項目情報','管理手数料率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','hosyo_flg','送金ルール入金項目情報','滞納保証有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_nk_cmrule','so_no','送金ルール入金項目情報','送金先No','0','0','','','','','画面上での使用箇所が不明'); ");
                        break;
                    }
                case "送金ルール情報-送金ルール控除項目情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','sorule_guid','送金ルール控除項目情報','送金ルールユニークID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','sorule_no','送金ルール控除項目情報','送金ルール管理No','1','10','','2','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojo_taisyokbn','送金ルール控除項目情報','契約時/更新時区分','1','7','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','nkin_sortorder','送金ルール控除項目情報','行No','0','49','','4','','','最大値が不明のためとりあえず50で設定しておく (行Noは0から始まっている)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','taisyonkin_no','送金ルール控除項目情報','控除対象(入金項目No)','9000','9999','','','','入金項目','同行Noに重複した入金項目を設定できないため中間ファイルのキーからは除外する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','sosai_flg','送金ルール控除項目情報','相殺予定フラグ','0','1','1','','','','0:別途請求　1：相殺'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojocalc_kbn','送金ルール控除項目情報','控除額基準','1','2','2','','','','1:控除率 2:控除額'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojo_gak','送金ルール控除項目情報','控除額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojo_gakzeikbn','送金ルール控除項目情報','控除額税区分','1','3','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojonkin_no','送金ルール控除項目情報','控除項目(入金項目No)','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojo_rit','送金ルール控除項目情報','控除率','0','100','','','','','999まで入力可だが100%に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojo_ritzeikbn','送金ルール控除項目情報','控除率税区分','1','2','2','','','','3:税入力は確認できず'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojo_ritutizei','送金ルール控除項目情報','控除率税有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','zei_rit','送金ルール控除項目情報','適用税率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','tateazu_flg','送金ルール控除項目情報','立替回収または預り金','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','kojo_zeigak','送金ルール控除項目情報','控除税額','0','99999999999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','sotaisyo_flg','送金ルール控除項目情報','送金対象フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sorule_kojo_cmrule','sotaisyonkin_no','送金ルール控除項目情報','送金対象入金項目No','1000','9999','','','','入金項目',''); ");
                        break;
                    }

                case "契約情報-契約基本情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','ky_guid','契約基本情報','契約管理Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','bk_guid','契約基本情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','hy_guid','契約基本情報','部屋No','','16','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','ky_no','契約基本情報','契約No','1','999','','3','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','ky_deleteflg','契約基本情報','削除フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','delete_guid','契約基本情報','削除Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','delete_day','契約基本情報','削除日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','delete_cnt','契約基本情報','削除復旧回数','0','0','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','syokai_kyymd','契約基本情報','初回契約日','1900/01/01','2100/12/31','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','status','契約基本情報','契約状況(ステータス)','1','3','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','cancelriyu','契約基本情報','キャンセル理由','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','status_ymd','契約基本情報','ステータス変更日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','cyukai_gy_fudono','契約基本情報','仲介業者No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','syunin_logonuser_no','契約基本情報','取引主任者No','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','tetuke_gak1','契約基本情報','手付預り額①','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','tetuke_ymd1','契約基本情報','手付預り日①','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','tetuke_biko1','契約基本情報','手付預り備考①','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','kaiyaku_flg','契約基本情報','解約フラグ','0','999','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','history','契約基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','rowid','契約基本情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','movefrom_kyguid','契約基本情報','移動元契約管理Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','moveto_kyguid','契約基本情報','移動先契約管理Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','svbunrui_no','契約基本情報','サービス分類','-1','-1','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','krbunrui_no','契約基本情報','会計グループ分類','0','0','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','kaiyaku_uketukekbn','契約基本情報','解約受付区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','kaiyaku_months','契約基本情報','解約受付月数','0','99','3','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','kaiyaku_days','契約基本情報','解約受付日数','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','kaiyaku_day','契約基本情報','解約受付日にち','0','31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','cyukai_tantoname','契約基本情報','仲介担当者名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata','cyukai_tantonamesjis','契約基本情報','仲介担当者名SJIS','','100','','','','',''); ");
                        break;
                    }
                case "契約情報-契約履歴情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','ky_guid','契約履歴情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','ky_recno','契約履歴情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','ko_no','契約履歴情報','更新No','1','999','','','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','henko_no','契約履歴情報','改定No','1','999','','','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','sqdata_guid','契約履歴情報','請求データGuid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','ky_bango','契約履歴情報','契約番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','ky_ymd','契約履歴情報','契約日','1900/01/01','2100/12/31','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kystart_ymd','契約履歴情報','契約開始日','1900/01/01','2100/12/31','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kyend_ymd','契約履歴情報','契約終了日','1900/01/01','2100/12/31','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','henko_ymd','契約履歴情報','条件変更日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','tuti_ymd','契約履歴情報','契約更新通知日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','print_ymd','契約履歴情報','通知書印刷日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kanri_gy_fudono','契約履歴情報','管理業者No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','gy_hosyono','契約履歴情報','賃貸保証業者No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hosyo_naiyo','契約履歴情報','賃貸保証内容','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kokyaku_bango','契約履歴情報','賃貸保証顧客番号','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kyrui_no','契約履歴情報','契約分類No','1','9999','','','','契約分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','siyo_mokuteki','契約履歴情報','使用目的','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kosin_umu','契約履歴情報','契約更新業務有無','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','yatin_kbn','契約履歴情報','家賃入金区分','1','20','2','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','fkae_startym','契約履歴情報','口座振替開始日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','fkae_willstartflg','契約履歴情報','口座振替開始待ちフラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','yatin_kozakbn','契約履歴情報','家賃入金口座区分','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','yatin_kozano','契約履歴情報','契約一時金入金口座No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','maitukiyatin_kozano','契約履歴情報','毎月分入金口座No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','yokugetu_uketoriflg','契約履歴情報','翌月分受取り有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','biko','契約履歴情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','tougetu_sagakuflg','契約履歴情報','当月分の差額受取り有無','0','0','0','','','','10画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','nextky_startymd','契約履歴情報','次回契約開始日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','nextky_endymd','契約履歴情報','次回契約終了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kosin_hiwariflg','契約履歴情報','更新時日割り有無','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','sokojorule_kbn','契約履歴情報','送金控除ルール適用有無','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','soyotei_ymdflg','契約履歴情報','送金予定日使用フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','soyotei_ymd','契約履歴情報','送金予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kanritesu_flg','契約履歴情報','部屋固定管理手数料フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kanritesu_gak','契約履歴情報','部屋固定管理手数料額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','nextnkinset_kbn','契約履歴情報','次回更新設定区分','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','sqdata_startymd','契約履歴情報','請求データ作成開始日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','sqdata_endymd','契約履歴情報','請求データ作成終了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','history','契約履歴情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','biko2','契約履歴情報','備考2(基本)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hoken_biko','契約履歴情報','備考(保険)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','nkinsime_ymd','契約履歴情報','入金締め日','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','yatin_jisansaki','契約履歴情報','家賃持参先','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hoken_kikan','契約履歴情報','保険期間','0','0','','','','','10画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hoken_gak','契約履歴情報','保険額','0','0','','','','','10画面上での設定箇所が不明'); ");
                        // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','confirmky_sekininsya','契約履歴情報','確認事項(契約)責任者','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','confirmky_ymd','契約履歴情報','確認事項(契約)内容確認日','1900/01/01','2100/12/31','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','confirmky_print','契約履歴情報','確認事項(契約)印刷時','','500','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','confirmkai_sekininsya','契約履歴情報','確認事項(解約)責任者','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','confirmkai_ymd','契約履歴情報','確認事項(解約)内容確認日','1900/01/01','2100/12/31','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','confirmkai_print','契約履歴情報','確認事項(解約)印刷時','','500','','','','',''); ")
                        // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hikiuke_name','契約履歴情報','身元引受人名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hikiuke_addr','契約履歴情報','身元引受人住所','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hikiuke_tel','契約履歴情報','身元引受人連絡先','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kohokennkin_ymd','契約履歴情報','保険料入金日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kokanryo_ymd','契約履歴情報','更新完了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kokanryotuti_ymd','契約履歴情報','更新完了通知日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kotuti_ymd','契約履歴情報','更新通知日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kosaisoku_ymd','契約履歴情報','催促実施日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kosyoruiuke_ymd','契約履歴情報','書類返送受取日','1900/01/01','2100/12/31','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','kokairenraku_umu','契約履歴情報','更新時の解約検討連絡有無','0','0','','','','','10画面上での設定箇所が不明'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','kokairenraku_ymd','契約履歴情報','更新時の解約検討連絡受付日','1900/01/01','2100/12/31','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','kokairenraku_logonuser_no','契約履歴情報','更新時の解約検討連絡受付担当者','0','0','','','','','10画面上での設定箇所が不明'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','kokairenraku_biko','契約履歴情報','更新時の解約検討連絡備考','','100','','','','',''); ")
                        // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','maitukisq_umu','契約履歴情報','毎月分請求有無','1','1','1','','','','10画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','hikiuke_namesjis','契約履歴情報','身元引受人名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','kofkae_ymd','契約履歴情報','口座振替日','1900/01/01','2100/12/31','','','','',''); ");
                        // 20170530 翌月受取フラグおよび月数項目追加対応 -chg sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kihon','yokugetu_uketorimonth','契約履歴情報','翌月分受取り月','1','1','1','','','','10画面上での設定箇所が不明'); ")
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','yokugetu_uketorimonth','契約履歴情報','翌月分受取り月','0','99','','','','',''); ");
                        // 20170530 翌月受取フラグおよび月数項目追加対応 -chg end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','ky_logonuser_no','契約履歴情報','契約担当者No','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','sq_logonuser_no','契約履歴情報','請求担当者No','1','99999','','','','',''); ");
                        // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add sta
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','torihiki_tesumoto','契約履歴情報','配分割合_元付','0','100','','','','','まだ画面に実装されていないためフィールド名から最大最小値を判断'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','torihiki_tesukyaku','契約履歴情報','配分割合_客付','0','100','','','','','まだ画面に実装されていないためフィールド名から最大最小値を判断'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','torihiki_tesugak','契約履歴情報','客付会社の手数料額','0','99999999999','','','','','まだ画面に実装されていないためフィールド名から最大最小値を判断'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','torihiki_tesuzeikbn','契約履歴情報','客付会社の手数料税区分','1','3','','','','','まだ画面に実装されていないためフィールド名から最大最小値を判断'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kihon','torihiki_tesuzeigak','契約履歴情報','客付会社の手数料税額','0','99999999999','','','','','まだ画面に実装されていないためフィールド名から最大最小値を判断'); ");
                        break;
                    }
                // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add end
                case "契約情報-契約契約者情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kys','ky_guid','契約契約者情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kys','ky_recno','契約契約者情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kys','kys_cnt','契約契約者情報','並び順No','1','3','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kys','kys_no','契約契約者情報','契約者No','1','999999999','','','','契約者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kys','nyukyo_flg','契約契約者情報','入居フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kys','history','契約契約者情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "契約情報-契約入居者情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','ky_guid','契約入居者情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','ky_recno','契約入居者情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','nyukyosya_cnt','契約入居者情報','入居者No','1','50','','3','','','10の最大値が不明なのでV7の最大値50に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','kys_no','契約入居者情報','契約者No','1','999999999','','','','契約者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','name','契約入居者情報','氏名','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','nameu','契約入居者情報','Unicode名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','kana','契約入居者情報','カナ','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','gender','契約入居者情報','性別','-1','2','-1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','aidagara','契約入居者情報','続柄','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','birthday','契約入居者情報','生年月日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','kinmusaki','契約入居者情報','勤務先','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','tel','契約入居者情報','連絡先','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','mobiletel','契約入居者情報','携帯電話番号','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','biko','契約入居者情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nyukyo','history','契約入居者情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "契約情報-契約保証人情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hosyonin','ky_guid','契約保証人情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hosyonin','ky_recno','契約保証人情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hosyonin','hosyonin_cnt','契約保証人情報','並び順','1','2','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hosyonin','kys_no','契約保証人情報','契約者No','1','999999999','','','','契約者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hosyonin','hosyonin_no','契約保証人情報','保証人No','1','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hosyonin','history','契約保証人情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "契約情報-契約車情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','ky_guid','契約車情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','ky_recno','契約車情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','car_cnt','契約車情報','車情報No','1','50','','3','','','10の最大値が不明なのでV7の最大値50に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','carmaker','契約車情報','メーカー','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','carname','契約車情報','車名','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','carcolor','契約車情報','車色','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','carnumber','契約車情報','ナンバー','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','biko','契約車情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','history','契約車情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_car','car_kukaku','契約車情報','駐車区画','','100','','','','',''); ");
                        break;
                    }
                case "契約情報-契約保険情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','ky_guid','契約保険情報','契約解約Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','hoken_no','契約保険情報','保険種類No','1','50','','2','','','10の最大値が不明なのでV7の最大値50に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','gy_hokenno','契約保険情報','保険業者No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','ky_ymd','契約保険情報','契約日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','kystart_ymd','契約保険情報','適用開始年月日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','kyend_ymd','契約保険情報','適用終了年月日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','hoken_gak','契約保険情報','保険金額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','mankituti_flg','契約保険情報','満期案内通知有無','0','1','0','','','','画面上での設定箇所が不明'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','biko','契約保険情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','history','契約保険情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hoken','syoken_bango','契約保険情報','証券番号','','200','','','','',''); ");
                        break;
                    }
                case "契約情報-契約特約事項情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_tokuyaku','ky_guid','契約特約事項情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_tokuyaku','ky_recno','契約特約事項情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_tokuyaku','tokuyaku_grpno','契約特約事項情報','特約グループNo','1','3','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_tokuyaku','naiyo','契約特約事項情報','文章内容','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_tokuyaku','history','契約特約事項情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "契約情報-契約メモ情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_memo','ky_guid','契約メモ情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_memo','ky_recno','契約メモ情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_memo','memo_no','契約メモ情報','メモNo','1','50','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_memo','memo','契約メモ情報','メモ内容','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_memo','history','契約メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "契約情報-契約入金項目情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','ky_guid','契約入金項目情報','契約管理GUID','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','ky_recno','契約入金項目情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','tuki_kbn','契約入金項目情報','月区分','1','3','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','nkin_no','契約入金項目情報','入金項目No','1000','9999','','4','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','nkin_recno','契約入金項目情報','入金項目レコードNo','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','nkin_sortorder','契約入金項目情報','並び順No','1','50','','','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','nkin_kbn','契約入金項目情報','入金項目区分','0','11','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_gak','契約入金項目情報','請求額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_zeikbn','契約入金項目情報','税区分','0','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_zeigak','契約入金項目情報','請求税額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','calc_kbn','契約入金項目情報','算出区分','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','calc_nkinno','契約入金項目情報','算出基準入金項目No','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','calc_monthcnt','契約入金項目情報','算出ヶ月','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sqsaki_no','契約入金項目情報','請求先No','1','999999999','','','','契約者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','nkbn_yotei','契約入金項目情報','入金方法','1','20','2','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_mmkbn','契約入金項目情報','請求月区分','1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','frstart_ymd','契約入金項目情報','フリーレント適用開始日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','frend_ymd','契約入金項目情報','フリーレント適用終了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','frsq_gak','契約入金項目情報','フリーレント終了月請求額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sqstart_ymd','契約入金項目情報','請求開始月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_ptn','契約入金項目情報','請求パターン','1','2','','','','請求パターン',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_interval','契約入金項目情報','請求間隔','1','3','','','','請求間隔',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_nen','契約入金項目情報','請求発生年','','100','','','','請求発生年',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','sq_tuki','契約入金項目情報','請求発生月','','100','','','','請求発生月',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','zei_rit','契約入金項目情報','適用税率','0','10','','','','','最大値を10にしておく (20160210 現在 8%)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','biko','契約入金項目情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','nkin_guid','契約入金項目情報','入金項目Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','history','契約入金項目情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin','fr_kbn','契約入金項目情報','フリーレント適用区分','0','1','0','','','',''); ");
                        break;
                    }
                case "契約情報-契約次回入金項目情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','ky_guid','契約次回入金項目情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','ky_recno','契約次回入金項目情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','tuki_kbn','契約次回入金項目情報','月区分','1','3','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','nkin_no','契約次回入金項目情報','入金項目No','1000','9999','','4','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','nkin_recno','契約次回入金項目情報','入金項目レコードNo','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','nkin_sortorder','契約次回入金項目情報','並び順No','1','50','','','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','nkin_kbn','契約次回入金項目情報','入金項目区分','0','11','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_gak','契約次回入金項目情報','請求額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_zeikbn','契約次回入金項目情報','税区分','0','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_zeigak','契約次回入金項目情報','請求税額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','calc_kbn','契約次回入金項目情報','算出区分','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','calc_nkinno','契約次回入金項目情報','算出基準入金項目No','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','calc_monthcnt','契約次回入金項目情報','算出ヶ月','0','99','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sqsaki_no','契約次回入金項目情報','請求先No','1','999999999','','','','契約者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','nkbn_yotei','契約次回入金項目情報','入金方法','1','20','2','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_mmkbn','契約次回入金項目情報','請求月区分','1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sqstart_ymd','契約次回入金項目情報','請求開始月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_ptn','契約次回入金項目情報','請求パターン','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_interval','契約次回入金項目情報','請求発生間隔','1','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_nen','契約次回入金項目情報','請求対象年','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','sq_tuki','契約次回入金項目情報','請求対象月','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','zei_rit','契約次回入金項目情報','適用税率','0','10','','','','','最大値を10にしておく (20160210 現在 8%)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','biko','契約次回入金項目情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_nkin_nx','history','契約次回入金項目情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "契約情報-契約変動費各戸メーター情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','ky_guid','契約変動費各戸メーター情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','ky_recno','契約変動費各戸メーター情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','hyhendo_guid','契約変動費各戸メーター情報','部屋変動費Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','hendo_sortorder','契約変動費各戸メーター情報','変動費No','1','50','','3','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','useflg','契約変動費各戸メーター情報','請求対象フラグ','0','1','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','hendo_kbn','契約変動費各戸メーター情報','メーター分類','1','5','','','','メーター分類 ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','meter_name','契約変動費各戸メーター情報','メーター名','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','nkin_no','契約変動費各戸メーター情報','入金項目No','1','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','hendorule_no','契約変動費各戸メーター情報','変動費請求ルールNo','1','9999','','','','変動費マスタ',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','hendorule_biko','契約変動費各戸メーター情報','変動費請求ルール備考','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','sqsaki_no','契約変動費各戸メーター情報','請求先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','nkbn_yotei','契約変動費各戸メーター情報','入金方法','1','20','2','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','biko','契約変動費各戸メーター情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','history','契約変動費各戸メーター情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_hendo','sq_mmkbn','契約変動費各戸メーター情報','請求月','1','5','2','','','',''); ");
                        break;
                    }
                case "契約情報-契約控除ルール情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','ky_guid','契約控除ルール情報','契約Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','ky_recno','契約控除ルール情報','契約レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','taisyokbn','契約控除ルール情報','対象区分','1','6','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','nkin_no','契約控除ルール情報','控除入金項目No','1000','9999','','4','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','nkin_sortorder','契約控除ルール情報','表示順','1','50','','','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','sosai_flg','契約控除ルール情報','相殺予定フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','calc_kbn','契約控除ルール情報','控除額算出基準','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','kojo_gak','契約控除ルール情報','控除額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','kojo_gakzeikbn','契約控除ルール情報','控除額税区分','1','3','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','kojonkin_no','契約控除ルール情報','控除対象入金項目No','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','kojo_rit','契約控除ルール情報','控除率','0','100','','','','','999まで入力可だが100%に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','kojo_ritzeikbn','契約控除ルール情報','控除率税区分','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','kojo_ritutizei','契約控除ルール情報','控除率内税','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','zei_rit','契約控除ルール情報','適用税率','0','10','','','','','最大値を10にしておく (20160210 現在 8%)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','tateazu_flg','契約控除ルール情報','立替回収または預り金','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','kojo_zeigak','契約控除ルール情報','控除税額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','sotaisyo_flg','契約控除ルール情報','送金対象フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kojorule','sotaisyonkin_no','契約控除ルール情報','送金対象入金項目No','1000','9999','','','','',''); ");
                        break;
                    }
                case "契約情報-契約送金ルール情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','ky_guid','契約送金ルール情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','ky_recno','契約送金ルール情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','taisyokbn','契約送金ルール情報','対象区分','1','6','','3','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','nkin_no','契約送金ルール情報','入金項目No','1000','9999','','4','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','nkin_sortorder','契約送金ルール情報','表示順','1','50','','','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','sokin_rit','契約送金ルール情報','送金率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','kanrigak_rit','契約送金ルール情報','管理手数料率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_sorule','hosyo_flg','契約送金ルール情報','滞納保証有無','0','1','','','','',''); ");
                        break;
                    }
                case "契約情報-契約解約情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','ky_guid','契約解約情報','契約管理Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','ky_recno','契約解約情報','契約管理レコードNo','1','999','','2','','','最大値が不明なのでとりあえず999に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','kai_ymd','契約解約情報','解約日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','seisan_ymd','契約解約情報','解約精算費用決定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','seisan_completeymd','契約解約情報','解約精算業務完了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','uketuke_ymd','契約解約情報','解約受付日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','riyu','契約解約情報','解約理由','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','tatiai_ymd','契約解約情報','退去日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','uketuke_logonuser_no','契約解約情報','受付担当者No','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','tatiai_logonuser_no','契約解約情報','立会担当者No','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','seisan_logonuser_no','契約解約情報','精算担当者No','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_logonuser_no','契約解約情報','修繕担当者No','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_jisyano','契約解約情報','自社・支店No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_endymd','契約解約情報','修繕完了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kojiyoteistartymd','契約解約情報','工事予定期間開始日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kojiyoteiendymd','契約解約情報','工事予定期間終了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kojibasyo','契約解約情報','工事場所','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kojigaiyo','契約解約情報','工事概要','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','next_name','契約解約情報','退去後宛名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','next_namesjis','契約解約情報','退去後宛名Shift-jis','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','next_postcode','契約解約情報','退去後郵便番号','','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','next_addr1','契約解約情報','退去後住所①','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','next_addr2','契約解約情報','退去後住所②','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','next_tel1','契約解約情報','退去後電話番号①','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','yatin_kozano','契約解約情報','不足時入金口座','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','biko','契約解約情報','備考(解約精算)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','biko_tatiai','契約解約情報','備考(立会)','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','seisan_henkinymd','契約解約情報','返金予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','seisan_sqymd','契約解約情報','請求締め日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','kanritesu_flg','契約解約情報','管理手数料フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','kanritesu_gak','契約解約情報','管理手数料額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','rowid','契約解約情報','付箋Guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','history','契約解約情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','so_yoteiymd','契約解約情報','送金予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','tatiai','契約解約情報','立会い','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','tatiai_yoteiymd','契約解約情報','退去予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','bosyu_jokenymd','契約解約情報','募集条件確認日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','ow_logonuser_no','契約解約情報','家主連絡担当者No','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','bosyujoken','契約解約情報','募集条件内容','','500','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 契約解約情報 -del sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','confirmkai_sekininsya','契約解約情報','確認事項入力責任者','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','confirmkai_ymd','契約解約情報','確認事項確認日','1900/01/01','2100/12/31','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','confirmkai_print','契約解約情報','印刷時の確認事項','','500','','','','',''); ")
                        // 20160519 EXEUpdateに伴う修正 契約解約情報 -del end
                        // 20160620 EXEUpdateに伴う修正2 -del sta
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_jisyafutanumuflg','契約解約情報','修繕自社負担フラグ','1','2','2','','','','2がチェックOFF'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_jisyafutangak','契約解約情報','修繕自社負担額','0','99999999999','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_jisyafutanzeikbn','契約解約情報','修繕自社負担税区分','-1','3','-1','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_jisyafutanzeigak','契約解約情報','修繕自社負担税額','0','99999999999','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_bikojisyafutan','契約解約情報','修繕自社負担備考','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_sonotafutanumuflg','契約解約情報','修繕その他負担フラグ','1','2','2','','','','2がチェックOFF'); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_sonotafutangak','契約解約情報','修繕その他負担額','0','99999999999','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_sonotafutanzeikbn','契約解約情報','修繕その他負担税区分','-1','3','-1','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_sonotafutanzeigak','契約解約情報','修繕その他負担税額','0','99999999999','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_sonotafutanname','契約解約情報','修繕その他負担者','','100','','','','',''); ")
                        // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('kydata_kai','szen_bikosonotafutan','契約解約情報','修繕その他負担備考','','100','','','','',''); ")
                        // 20160620 EXEUpdateに伴う修正2 -del end
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','tatiai_yoteitime','契約解約情報','退去予定時刻','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','next_keisyo','契約解約情報','退去後敬称','','10','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 契約解約情報 -add
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','nkin_kaiyakutukinoprintflg','契約解約情報','解約月賃料は扱わないフラグ','1','2','2','','','',''); ");
                        // 20160620 EXEUpdateに伴う修正2 -add sta
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kysno','契約解約情報','契約者修繕請求先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kysnkbn','契約解約情報','契約者修繕入金区分','1','20','2','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kyssorit','契約解約情報','契約者修繕送金率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_kyssogak','契約解約情報','契約者修繕送金額','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owno','契約解約情報','家主修繕控除先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owkaisyukbn','契約解約情報','家主修繕控除請求区分','1','20','2','','','','要確認'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owsosakisoruleguid','契約解約情報','家主修繕控除先送金ルールGuid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owsosakisoruleno','契約解約情報','家主修繕控除先送金ルールNo','1','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owsokozano','契約解約情報','家主修繕控除先口座No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owsqsimeymd','契約解約情報','家主修繕請求締日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_ownkbn','契約解約情報','家主修繕請求入金区分','1','2','2','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owyatinkozano','契約解約情報','家主修繕請求振込先口座No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owsqsakino','契約解約情報','家主修繕請求先No','1','999999999','','','','',''); ");
                        // 20160620 EXEUpdateに伴う修正2 -add end
                        // 20160829 革命10バージョンアップに伴う修正 -add sta
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owkojoymd','契約解約情報','家主修繕控除予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kai','szen_owsqymd','契約解約情報','家主修繕請求書発行予定日','1900/01/01','2100/12/31','','','','',''); ");
                        break;
                    }
                // 20160829 革命10バージョンアップに伴う修正 -add end
                case "契約情報-契約修繕見積情報":    // 20160621 修繕関連移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','ky_guid','契約修繕見積情報','契約No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','ky_recno','契約修繕見積情報','契約レコードNo','1','999','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','szen_mituno','契約修繕見積情報','見積No','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','mitu_title','契約修繕見積情報','見積タイトル','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','mitu_bango','契約修繕見積情報','見積番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','mitu_ymd','契約修繕見積情報','見積日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','seiyaku_flg','契約修繕見積情報','成約フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','seiyaku_ymd','契約修繕見積情報','成約日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','futan_kbn','契約修繕見積情報','負担区分(全体・個別)','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','kys_futanrit','契約修繕見積情報','契約者負担率(全体用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','ow_futanrit','契約修繕見積情報','家主負担率(全体用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','jisya_futanrit','契約修繕見積情報','自社負担率(全体用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','zei_kbn','契約修繕見積情報','税適用区分(全体・個別)','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','gokei_zeikbn','契約修繕見積情報','税区分(全体用)','0','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','gokei_zeirit','契約修繕見積情報','適用税率','3','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','kys_gokeizeigak','契約修繕見積情報','契約者税額(全体用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','ow_gokeizeigak','契約修繕見積情報','家主税額(全体用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','jisya_gokeizeigak','契約修繕見積情報','自社税額(全体用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszen','hasu_futankbn','契約修繕見積情報','端数負担者区分','100','900','','','','',''); ");
                        break;
                    }
                case "契約情報-契約修繕見積詳細情報":    // 20160621 修繕関連移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','ky_guid','契約修繕見積詳細情報','契約No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','ky_recno','契約修繕見積詳細情報','契約レコードNo','1','999','','2','','','最大値が不明のためとりあえず最大値100を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','szen_mituno','契約修繕見積詳細情報','見積No','1','1','','','','','V7では見積が1つのみ作成可能なので1を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','szen_meisaino','契約修繕見積詳細情報','見積明細No','1','99','','','','','10側の最大値が不明の為V7の最大行数を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','szen_name','契約修繕見積詳細情報','修繕項目名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','tekiyo','契約修繕見積詳細情報','摘要','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','mitu_suryo','契約修繕見積詳細情報','見積数量','0','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','mitu_tani','契約修繕見積詳細情報','見積単位','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','mitu_tanka','契約修繕見積詳細情報','見積単価','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','mitu_zeikbn','契約修繕見積詳細情報','見積税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','mitu_zeigak','契約修繕見積詳細情報','見積税額(税入力用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','kys_futanrit','契約修繕見積詳細情報','契約者負担率(個別用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','ow_futanrit','契約修繕見積詳細情報','家主負担率(個別用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','jisya_futanrit','契約修繕見積詳細情報','自社負担率(個別用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','szen_gyno','契約修繕見積詳細情報','発注業者No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','jikko_suryo','契約修繕見積詳細情報','実行数量','0','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','jikko_tani','契約修繕見積詳細情報','実行単位','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','jikko_tanka','契約修繕見積詳細情報','実行単価','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','jikko_zeikbn','契約修繕見積詳細情報','実行税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kydata_kaiszenmeisai','jikko_zeigak','契約修繕見積詳細情報','実行税額(税入力用)','0','99999999999','','','','',''); ");
                        break;
                    }
                // Case "契約情報-契約鍵情報"   '部屋鍵情報へ書き込まれるためhydata_kagiを移行対象テーブルとする
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','hy_guid','契約鍵情報','物件GUID(Key)','','16','','1','','',''); ")
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','kagi_no','契約鍵情報','鍵明細No','1','50','','2','','専用鍵タイトル',''); ")
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','honsu','契約鍵情報','鍵本数','0','9999','','','','',''); ")
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','biko','契約鍵情報','備考','','100','','','','',''); ")
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','history','契約鍵情報','履歴','','-1','','','','',''); ")
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','hokan','契約鍵情報','保管場所','','100','','','','',''); ")
                // rtn_qry.Append(" INSERT INTO " & CVDBINFO_DBNAME & " VALUES ('hydata_kagi','gyshare','契約鍵情報','業者間での情報共有','','100','','','','',''); ")
                case "請求情報-預り金情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','azu_guid','預り金情報','預り金管理GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','azu_ymd','預り金情報','預り金処理日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','nkin_no','預り金情報','入金項目名','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','dispnkin_name','預り金情報','表示用入金項目名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','nkbn_no','預り金情報','入金区分','1','20','','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','azu_gak','預り金情報','預り額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','stakeholder_kbn','預り金情報','利害関係者区分','100','100','','','','','契約者固定'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','stakeholder_no','預り金情報','利害関係者No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','sui_guid','預り金情報','出納管理GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','tanto_no','預り金情報','担当者No','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','bk_guid','預り金情報','物件No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','hy_guid','預り金情報','部屋No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','ky_guid','預り金情報','契約No','','16','','1','','','契約Noへ物件、部屋を統合する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','ky_recno','預り金情報','契約レコードNo','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','tuki_kbn','預り金情報','月区分','1','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','ky_nkin_no','預り金情報','契約入金情報の入金項目No','1000','9999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','nkin_recno','預り金情報','入金項目レコードNo','0','0','','','','','使用されていない？'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','kynkin_guid','預り金情報','契約入金項目GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','yoteiazu_flg','預り金情報','預り予定フラグ','0','1','','','','','使用箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','printazuryo_guid','預り金情報','預り証・領収証GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','siwake_flg','預り金情報','仕訳フラグ','0','1','','','','','使用箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','biko','預り金情報','備考','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','history','預り金情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('azukanri','tujo_siwakegak','預り金情報','通常仕訳額','0','99999999999','','','','','使用箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        break;
                    }
                case "請求情報-未収滞納金情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','unyotaino_guid','未収滞納金情報','運用開始時未納・滞納金GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','bk_guid','未収滞納金情報','物件No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','hy_guid','未収滞納金情報','部屋No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','ky_guid','未収滞納金情報','契約No','','16','','1','','','契約Noへ物件、部屋を統合する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','ky_recno','未収滞納金情報','契約レコードNo','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','gt_ym','未収滞納金情報','該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sq_simeymd','未収滞納金情報','請求締め日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','nkin_no','未収滞納金情報','入金項目No','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','nkbn_yotei','未収滞納金情報','(入金予定)入金区分','1','20','','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','yotei_yatin_kozano','未収滞納金情報','(入金予定)家賃入金口座','1','999999','','','','家賃入金口座',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sqsaki_no','未収滞納金情報','請求先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sq_gak','未収滞納金情報','請求額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sq_zeikbn','未収滞納金情報','請求税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sq_zeigak','未収滞納金情報','請求税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','so_gak','未収滞納金情報','送金額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','so_zeikbn','未収滞納金情報','送金税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','so_zeigak','未収滞納金情報','送金税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sosumi_flg','未収滞納金情報','既に送金済みフラグ','1','2','2','','','','1が ON'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','kanrigak_rit','未収滞納金情報','管理手数料率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sosaki_no','未収滞納金情報','送金先家主No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sokoza_no','未収滞納金情報','送金先家主口座No','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','so_rit','未収滞納金情報','送金率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','soymd_gaitoukbn','未収滞納金情報','送金日決定方法 - 該当年月','1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','soymd_simekbn','未収滞納金情報','送金日決定方法 - 締日','1','31','31','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','soymd_sokintukikbn','未収滞納金情報','送金日決定方法 - 送金月','1','6','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','soymd_sokinsimekbn','未収滞納金情報','送金日決定方法 - 送金締日','1','31','10','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','zei_rit','未収滞納金情報','適用税率','0','8','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','sqbuild_flg','未収滞納金情報','請求データ生成フラグ','0','1','','','','','使用箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','tanto_no','未収滞納金情報','担当者No','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','biko','未収滞納金情報','備考','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','history','未収滞納金情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','kanri_taxflg','未収滞納金情報','管理手数料税フラグ','0','1','','','','','使用箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','kanritesu_zeiumukbn','未収滞納金情報','管理手数料税込フラグ','0','1','','','','','使用箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('unyotainodata','kanri_utizeiflg','未収滞納金情報','管理手数料率内税フラグ','0','1','','','','','使用箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        break;
                    }
                case "請求情報-その他請求情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sqmei_guid','その他請求情報','請求明細管理GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','gt_ym','その他請求情報','該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','nkin_no','その他請求情報','入金項目名','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','dispnkin_name','その他請求情報','入金項目表示用名称','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','nkbn_yotei','その他請求情報','(入金予定)入金区分','1','20','','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sq_gak','その他請求情報','請求額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sq_zeikbn','その他請求情報','請求税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sq_zeigak','その他請求情報','請求税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sq_simeymd','その他請求情報','請求締切日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','multisqsaki_flg','その他請求情報','複数請求先フラグ','0','0','','','','','使用されていない可能性がある'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sqsaki_kbn','その他請求情報','請求先区分','100','900','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sqsaki_no','その他請求情報','請求先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','yotei_yatin_kozano','その他請求情報','(入金予定)家賃振込先口座No','1','999999','','','','家賃入金口座',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','fkae_rirekino','その他請求情報','口座振替履歴No','0','0','','','','','使用箇所が不明のため0を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','fkae_rirekimeino','その他請求情報','口座振替履歴明細No','0','0','','','','','使用箇所が不明のため0を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','fkae_sqtaisyoflg','その他請求情報','振替請求対象フラグ 1:請求中　2:未請求','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','fkae_sqselectflg','その他請求情報','振替請求対象フラグ　1:対象　2:対象外','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','fkae_sqflg','その他請求情報','振替請求中フラグ','0','1','','','','','使用箇所が不明のため0を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','ky_guid','その他請求情報','契約No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','ky_recno','その他請求情報','契約レコードNo','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','tuki_kbn','その他請求情報','月区分','1','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','nkin_recno','その他請求情報','入金項目明細No','1','1','','','','','契約入金項目情報から引用 (設定箇所が不明の為)'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','nkin_kbn','その他請求情報','入金項目区分','0','11','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','kynkin_guid','その他請求情報','契約入金項目GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','nextky_flg','その他請求情報','次回契約情報フラグ','0','1','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','hendosq_guid','その他請求情報','変動費請求GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','hendo_kbn','その他請求情報','変動費区分','0','1','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','kensin_ymd','その他請求情報','変動費検針日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','bk_guid','その他請求情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','hy_guid','その他請求情報','部屋No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','so_gak','その他請求情報','送金額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','so_zeikbn','その他請求情報','送金税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','so_zeigak','その他請求情報','送金税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','etcsq_flg','その他請求情報','その他請求フラグ','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','etcsq_sorit','その他請求情報','その他請求の送金率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','etcsq_soymd','その他請求情報','その他請求の指定送金日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','bunkatu_guid','その他請求情報','分割GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sh_kbn','その他請求情報','支払返金区分','0','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','shsaki_kbn','その他請求情報','支払返金先区分','0','900','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','shsaki_no','その他請求情報','支払返金先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','shsaki_kozano','その他請求情報','支払返金先口座No','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','ikkatu_nkin_bango','その他請求情報','一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','jisya_no','その他請求情報','支店No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','tanto_no','その他請求情報','担当者No','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','kjsqkmk_guid','その他請求情報','控除請求項目GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','koteirule_guid','その他請求情報','固定控除ルールGUID(請求)','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','kojorule_flg','その他請求情報','控除ルールフラグ','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sq_hakkoyoteiymd','その他請求情報','請求書発行予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sqsort_no','その他請求情報','請求ソートNo','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','tateazu_flg','その他請求情報','立替・預りフラグ','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','tateazu_guid','その他請求情報','立替・預り設定時の預り予定GUID（預り金データ）','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sikikinzuiji_guid','その他請求情報','敷金保証金随時処理GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','unyotaino_guid','その他請求情報','運用開始時の未納・滞納金GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sokaisyu_guid','その他請求情報','送金回収GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','szen_kbn','その他請求情報','修繕区分','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','szen_no','その他請求情報','随時修繕No','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','szen_sqno','その他請求情報','修繕請求No','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','zei_rit','その他請求情報','適用税率','0','8','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','edit_flg','その他請求情報','編集フラグ','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','lock_flg','その他請求情報','更新ロックフラグ','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','biko','その他請求情報','備考','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','history','その他請求情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','tatekae_siwakegak','その他請求情報','立替仕訳額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','tatekae_siwakezeigak','その他請求情報','立替仕訳税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','sq_torikomiymd','その他請求情報','取込日（インポート実施日）','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','cvpay_no','その他請求情報','コンビニ収納サービスNo','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','kokyaku_bango','その他請求情報','顧客番号','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','cvpay_sqflg','その他請求情報','コンビニ収納請求フラグ','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','etcsq_kanritesukbn','その他請求情報','その他請求管理手数料区分','1','2','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','misyu_siwake_flg','その他請求情報','未収仕訳フラグ','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','zatusyu_kbn','その他請求情報','雑収入区分','0','0','','','','','設定箇所が不明のためとりあえず最大最小値を設定しておく'); ");
                        // 20160519 EXEUpdateに伴う修正 請求情報 -add sta
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','shdirect_flg','その他請求情報','直接支払返金フラグ','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','koteishrule_guid','その他請求情報','固定支払ルールGUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('sqdata','kojorule_taisyokbn','その他請求情報','控除ルール対象区分','1','6','','','','',''); ");
                        break;
                    }
                // 20160519 EXEUpdateに伴う修正 請求情報 -add end
                case "請求情報-変動費検針情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','bk_guid','変動費検針情報','物件No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','hendo_kbn','変動費検針情報','変動区分','1','5','','','','メーター分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_ymd','変動費検針情報','検針日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','gt_ym','変動費検針情報','該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_ptn','変動費検針情報','検針パターン','1','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','hendo_guid','変動費検針情報','変動guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','hy_guid','変動費検針情報','部屋No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','hendosq_guid','変動費検針情報','変動請求guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','ky_guid','変動費検針情報','契約No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','ky_recno','変動費検針情報','契約レコードNo','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','sqsaki_no','変動費検針情報','請求先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','nkin_no','変動費検針情報','入金項目名','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','nkbn_no','変動費検針情報','入金区分','1','20','','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','sqsime_ymd','変動費検針情報','請求締切日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','zenkai_ymd','変動費検針情報','前回検針日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','zenkai_metervalue','変動費検針情報','前回値','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','zenkai_siyoryo','変動費検針情報','前回使用量','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','zenkai_sqgak','変動費検針情報','前回請求額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','zenkai_sqzeigak','変動費検針情報','前回請求税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_metervalue','変動費検針情報','今回値','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_siyoryo','変動費検針情報','今回使用量','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_ryokin1','変動費検針情報','検針料金1','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_ryokin2','変動費検針情報','検針料金2','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_ryokin3','変動費検針情報','検針料金3','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_ryokin4','変動費検針情報','検針料金4','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_ryokin5','変動費検針情報','検針料金5','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_sqgak','変動費検針情報','請求額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_sqzeigak','変動費検針情報','請求税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_biko','変動費検針情報','変動費検針情報備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_soyoteiymd','変動費検針情報','送金予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_sorit','変動費検針情報','送金率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_sogak','変動費検針情報','送金額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','kensin_sozeigak','変動費検針情報','送金税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hendodata_meisai','zenkaihendosq_guid','変動費検針情報','前回変動請求guid','','16','','','','',''); ");
                        break;
                    }
                case "請求情報-家主固定控除情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','koteirule_guid','家主固定控除情報','固定控除ルールGUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','bk_guid','家主固定控除情報','物件No','','16','','1','','','重複がゆるされているため親マスタのみに使用する'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','nkin_sortorder','家主固定控除情報','入金項目ソートNo','1','50','','','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','nkin_no','家主固定控除情報','入金項目No','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','nkin_name','家主固定控除情報','入金項目名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','hy_guid','家主固定控除情報','部屋No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','sosai_flg','家主固定控除情報','相殺フラグ','1','1','','','','','10画面上で設定箇所が存在しない'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','gak','家主固定控除情報','額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','zei_kbn','家主固定控除情報','税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','zei_gak','家主固定控除情報','税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','nkbn_yotei','家主固定控除情報','入金区分','1','20','','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','yatin_kozano','家主固定控除情報','家賃入金口座No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','gtstart_ym','家主固定控除情報','適用開始該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','gtend_ym','家主固定控除情報','適用終了該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','tateazu_flg','家主固定控除情報','立替・預りフラグ','1','2','2','','','','1:チェックON   2:チェックOFF'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','sq_mmkbn','家主固定控除情報','控除請求月区分','1','5','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','sq_ptn','家主固定控除情報','控除請求発生パターン','1','2','','','','請求パターン',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','sq_interval','家主固定控除情報','控除請求発生間隔','1','3','','','','請求間隔',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','sq_nen','家主固定控除情報','控除請求発生月指定年区分','','100','','','','請求発生年',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','sq_tuki','家主固定控除情報','控除請求発生月指定月区分','','100','','','','請求発生月',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','zei_rit','家主固定控除情報','適用税率','0','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','biko','家主固定控除情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','history','家主固定控除情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','sq_simekbn','家主固定控除情報','請求締区分','0','0','','','','','設定箇所が不明のためとりあえず0を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('koteirule','ryosyu_flg','家主固定控除情報','領収フラグ','1','2','2','','','','1:チェックON   2:チェックOFF'); ");
                        break;
                    }
                case "請求情報-家主請求控除情報":
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kj_guid','家主請求控除情報','控除GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kjkmk_guid','家主請求控除情報','控除項目GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','bk_guid','家主請求控除情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','hy_guid','家主請求控除情報','部屋No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kj_yoteiymd','家主請求控除情報','控除予定日','1900/01/01','2100/12/31','','','●','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kjsort_no','家主請求控除情報','項目ソートNo','1','50','','','','','最大数が不明なためとりあえず50を設定しておく。画面上では入金項目の重複がゆるされていないため中間ファイルではキーとしない。'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','gt_ym','家主請求控除情報','該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','nkin_no','家主請求控除情報','入金項目名','1000','9999','','','','入金項目',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','dispnkin_name','家主請求控除情報','表示用入金項目名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kj_gak','家主請求控除情報','控除額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kj_zeikbn','家主請求控除情報','控除税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kj_zeigak','家主請求控除情報','控除税額','0','99999999999','','','','金額',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','lock_flg','家主請求控除情報','更新ロックフラグ','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','ryosyu_flg','家主請求控除情報','領収扱いフラグ','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','tateazu_flg','家主請求控除情報','立替・預りフラグ','1','2','2','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','tateazu_guid','家主請求控除情報','預り予定GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','bunkatu_guid','家主請求控除情報','分割グループGUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','koteirule_guid','家主請求控除情報','固定控除ルールGUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kojorule_flg','家主請求控除情報','控除ルールフラグ','0','0','','','','','設定値が不明のためとりあえず0を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','so_kakuymd','家主請求控除情報','送金（控除）確定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sokotik_guid','家主請求控除情報','送金データ構築GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sorule_guid','家主請求控除情報','送金ルールGUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sorule_no','家主請求控除情報','送金ルールNo','1','10','','','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sosaki_sorule_guid','家主請求控除情報','(送金先)送金ルールGUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sosaki_sorule_no','家主請求控除情報','(送金先)送金ルールNo','1','10','','','','','最大値が不明のためとりあえず10で設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sosaki_no','家主請求控除情報','(控除先指定時）送金先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sokoza_no','家主請求控除情報','(控除先指定時）送金先口座No','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sikikinzuiji_guid','家主請求控除情報','敷金保証金随時処理GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','ky_guid','家主請求控除情報','契約No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','ky_recno','家主請求控除情報','契約レコードNo','1','999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','sokaisyu_guid','家主請求控除情報','送金回収GUID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kanritesu_kbn','家主請求控除情報','管理手数料区分','0','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kanritesu_cyosyukbn','家主請求控除情報','管理手数料請求条件','0','8','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kanrigak_rit','家主請求控除情報','管理手数料率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','szen_kbn','家主請求控除情報','修繕区分','0','0','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','szen_no','家主請求控除情報','随時修繕No','0','0','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','szen_sqno','家主請求控除情報','修繕請求No','0','0','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','siwake_flg','家主請求控除情報','仕訳フラグ','0','0','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','zei_rit','家主請求控除情報','適用税率','0','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','biko','家主請求控除情報','備考','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kj_torikomiymd','家主請求控除情報','控除データ取込日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','history','家主請求控除情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kubungrp_guid','家主請求控除情報','区分グループguid','','16','','','','',''); ");
                        // 20160519 EXEUpdateに伴う修正 家主請求控除情報 -add
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('kjdata','kojorule_taisyokbn','家主請求控除情報','控除ルール対象区分','1','6','','','','',''); ");
                        break;
                    }

                case "クレーム情報-クレーム基本情報":    // 20160524 クレーム情報移行処理実装 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','claim_no','クレーム基本情報','クレームNo','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','status','クレーム基本情報','対応状況','1','4','3','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','title','クレーム基本情報','タイトル','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_ymd','クレーム基本情報','受付日時','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_logonuser_no','クレーム基本情報','受付担当者','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_name','クレーム基本情報','連絡者名','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_tel1','クレーム基本情報','連絡者TEL1','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_tel2','クレーム基本情報','連絡者TEL2','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','renraku_timestart','クレーム基本情報','連絡可能時間（開始）','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','renraku_timeend','クレーム基本情報','連絡可能時間（終了）','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','kinkyu_kbn','クレーム基本情報','緊急度','1','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taio_limitymd','クレーム基本情報','対応期限','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','kasyo_ruino','クレーム基本情報','箇所分類No','1','50','','','','箇所分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','claim_ruino','クレーム基本情報','クレーム分類No','1','50','','','','クレーム分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_report','クレーム基本情報','内容','','1000','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taio_logonuser_no','クレーム基本情報','対応担当者','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','emailbiko','クレーム基本情報','電子メール備考','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','bk_guid','クレーム基本情報','物件No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hy_guid','クレーム基本情報','部屋No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','ky_guid','クレーム基本情報','契約No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','ky_recno','クレーム基本情報','契約レコードNo','0','999','','','','','契約情報は必須というわけではないので最小値を0に設定'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','reform_guid','クレーム基本情報','リフォームGuid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taiosaki_kbn','クレーム基本情報','対応先区分','1','4','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taiosaki_gy_no','クレーム基本情報','依頼業者No','1','999999','','','','修繕業者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taiosaki_tantoname','クレーム基本情報','業者担当者','','400','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taiosaki_tel','クレーム基本情報','業者担当者TEL','','20','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taio_finishymd','クレーム基本情報','完了日時','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taio_report','クレーム基本情報','結果入力','','1000','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hutan1_kbn','クレーム基本情報','負担者1','0','6','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hutan1_gaketc','クレーム基本情報','負担者1金額等','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hutan2_kbn','クレーム基本情報','負担者2','0','6','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hutan2_gaketc','クレーム基本情報','負担者2金額等','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hutan3_kbn','クレーム基本情報','負担者3','0','6','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hutan3_gaketc','クレーム基本情報','負担者3金額等','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','hutanbiko','クレーム基本情報','備考','','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','jisya_no','クレーム基本情報','自社・支店No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','jisya_name','クレーム基本情報','自社・支店名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','history','クレーム基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','rowid','クレーム基本情報','行ID','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_hm','クレーム基本情報','受付時刻','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taio_finishhm','クレーム基本情報','対応終了時刻','','5','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','uke_namesjis','クレーム基本情報','連絡者名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','taiosaki_tantonamesjis','クレーム基本情報','対応担当者名SJIS','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata','ow_no','クレーム基本情報','家主No','1','999999999','','','','家主',''); ");
                        break;
                    }
                case "クレーム情報-クレーム対応履歴情報":    // 20160524 クレーム情報移行処理実装 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claim_taio','claim_no','クレーム対応履歴情報','クレームNo','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claim_taio','taio_no','クレーム対応履歴情報','対応履歴No','1','20','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claim_taio','taio_ymd','クレーム対応履歴情報','対応日時','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claim_taio','taio_logonuser_no','クレーム対応履歴情報','対応担当者','1','99999','','','','担当者',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claim_taio','taio_report','クレーム対応履歴情報','内容','','500','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claim_taio','history','クレーム対応履歴情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claim_taio','taio_hm','クレーム対応履歴情報','対応時刻','','5','','','','',''); ");
                        break;
                    }
                case "クレーム情報-クレーム関連ファイル情報":    // 20160524 クレーム情報移行処理実装 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata_relfile','claim_no','クレーム関連ファイル情報','クレームNo','1','999999999','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata_relfile','file_no','クレーム関連ファイル情報','ファイルNo','0','4','','2','','','V7の最大値に合わせておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata_relfile','fullpath','クレーム関連ファイル情報','ファイルパス','','-1','','','','画像判別',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata_relfile','addtime','クレーム関連ファイル情報','追加日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata_relfile','biko','クレーム関連ファイル情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('claimdata_relfile','history','クレーム関連ファイル情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "修繕情報-修繕基本情報":          // 20160621 修繕関連移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','bk_guid','修繕基本情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','szen_no','修繕基本情報','修繕No','1','100','','2','','','最大値が不明のためとりあえず最大値100を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','hy_useflg','修繕基本情報','部屋使用フラグ','1','2','','','','','使用されていない可能性あり'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','hy_guid','修繕基本情報','部屋No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ky_useflg','修繕基本情報','契約使用フラグ','1','2','','','','','使用されていない可能性あり'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ky_guid','修繕基本情報','契約No','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','szen_name','修繕基本情報','修繕名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','szen_ukeymd','修繕基本情報','修繕受付日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','szen_endymd','修繕基本情報','修繕終了日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','tanto_no','修繕基本情報','修繕担当者','1','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','jisya_no','修繕基本情報','自社支店No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','koji_yoteistartymd','修繕基本情報','工事開始予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','koji_yoteiendymd','修繕基本情報','工事終了予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','koji_basyo','修繕基本情報','工事場所','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','koji_gaiyo','修繕基本情報','工事概要','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','koji_kagi','修繕基本情報','鍵','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','koji_tatiai','修繕基本情報','立会者','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','biko','修繕基本情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_futanrit','修繕基本情報','契約者負担率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_futanrit','修繕基本情報','家主負担率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','rowid','修繕基本情報','付箋rowid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','history','修繕基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','jisya_futanrit','修繕基本情報','自社負担率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_sqflg','修繕基本情報','契約者請求作成フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_no','修繕基本情報','契約者請求先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_gtym','修繕基本情報','契約者該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_sqsimeymd','修繕基本情報','契約者請求締日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_nkbn','修繕基本情報','契約者入金区分','1','20','1','','','入金区分',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_sorit','修繕基本情報','契約者送金率','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_yatinkozano','修繕基本情報','契約者振込先口座No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','kys_biko','修繕基本情報','契約者請求備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sqflg','修繕基本情報','家主請求作成フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_no','修繕基本情報','家主控除先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_kaisyukbn','修繕基本情報','家主控除請求区分','1','4','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_kojoymd','修繕基本情報','家主控除予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_kojogtym','修繕基本情報','家主控除該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sosaki_soruleguid','修繕基本情報','家主控除先送金ルールGuid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sosaki_sorule_no','修繕基本情報','家主控除先送金ルールNo','1','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sokozano','修繕基本情報','家主控除先口座No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sqymd','修繕基本情報','家主請求書発行予定日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sqgtym','修繕基本情報','家主請求該当年月','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sqsqsimeymd','修繕基本情報','家主請求締日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sqnkbn','修繕基本情報','家主請求入金区分','1','4','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sqyatinkozano','修繕基本情報','家主請求振込先口座No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_sqsakino','修繕基本情報','家主請求先No','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata','ow_biko','修繕基本情報','家主控除請求備考','','100','','','','',''); ");
                        break;
                    }
                case "修繕情報-修繕見積情報":          // 20160621 修繕関連移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','bk_guid','修繕見積情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','szen_no','修繕見積情報','修繕No','1','100','','2','','','最大値が不明のためとりあえず最大値100を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','szen_mituno','修繕見積情報','見積No','1','1','','3','','','V7では見積が1つのみ作成可能なので1を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','mitu_title','修繕見積情報','見積タイトル','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','mitu_bango','修繕見積情報','見積番号','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','mitu_ymd','修繕見積情報','見積日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','seiyaku_flg','修繕見積情報','成約フラグ','0','1','0','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','seiyaku_ymd','修繕見積情報','成約日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','futan_kbn','修繕見積情報','負担区分(全体・個別)','1','2','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','kys_futanrit','修繕見積情報','契約者負担率(全体用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','ow_futanrit','修繕見積情報','家主負担率(全体用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','jisya_futanrit','修繕見積情報','自社負担率(全体用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','zei_kbn','修繕見積情報','税適用区分(全体・個別)','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','gokei_zeikbn','修繕見積情報','税区分(全体用)','0','3','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','gokei_zeirit','修繕見積情報','適用税率','3','10','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','kys_gokeizeigak','修繕見積情報','契約者税額(全体用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','ow_gokeizeigak','修繕見積情報','家主税額(全体用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','jisya_gokeizeigak','修繕見積情報','自社税額(全体用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szen','hasu_futankbn','修繕見積情報','端数負担者区分','100','900','','','','',''); ");
                        break;
                    }
                case "修繕情報-修繕見積詳細情報":          // 20160621 修繕関連移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','bk_guid','修繕見積詳細情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','szen_no','修繕見積詳細情報','修繕No','1','100','','2','','','最大値が不明のためとりあえず最大値100を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','szen_mituno','修繕見積詳細情報','見積No','1','1','','3','','','V7では見積が1つのみ作成可能なので1を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','szen_meisaino','修繕見積詳細情報','見積明細No','1','99','','4','','','10側の最大値が不明の為V7の最大行数を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','szen_name','修繕見積詳細情報','修繕項目名','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','tekiyo','修繕見積詳細情報','摘要','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','mitu_suryo','修繕見積詳細情報','見積数量','0','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','mitu_tani','修繕見積詳細情報','見積単位','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','mitu_tanka','修繕見積詳細情報','見積単価','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','mitu_zeikbn','修繕見積詳細情報','見積税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','mitu_zeigak','修繕見積詳細情報','見積税額(税入力用)','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','kys_futanrit','修繕見積詳細情報','契約者負担率(個別用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','ow_futanrit','修繕見積詳細情報','家主負担率(個別用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','jisya_futanrit','修繕見積詳細情報','自社負担率(個別用)','0','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','szen_gyno','修繕見積詳細情報','発注業者No','1','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','jikko_suryo','修繕見積詳細情報','実行数量','0','99999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','jikko_tani','修繕見積詳細情報','実行単位','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','jikko_tanka','修繕見積詳細情報','実行単価','0','99999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','jikko_zeikbn','修繕見積詳細情報','実行税区分','1','3','1','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_szenmeisai','jikko_zeigak','修繕見積詳細情報','実行税額(税入力用)','0','99999999999','','','','',''); ");
                        break;
                    }
                case "修繕情報-修繕クレーム関連付け情報":          // 20160621 修繕関連移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_claim','bk_guid','修繕クレーム関連付け情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_claim','szen_no','修繕クレーム関連付け情報','修繕No','1','100','','2','','','最大値が不明のためとりあえず最大値100を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_claim','claim_no','修繕クレーム関連付け情報','クレームNo','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_claim','claimdata_sortorder','修繕クレーム関連付け情報','クレーム画面での並び順','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_claim','szendata_sortorder','修繕クレーム関連付け情報','修繕情報登録画面での並び順','1','999999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_claim','history','修繕クレーム関連付け情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "修繕情報-修繕関連ファイル情報":          // 20160627 修繕関連ファイル移行修正 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_relfile','bk_guid','修繕関連ファイル情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_relfile','szen_no','修繕関連ファイル情報','修繕No','1','100','','2','','','最大値が不明のためとりあえず最大値100を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_relfile','file_no','修繕関連ファイル情報','ファイルNo','0','4','','','','','最大値が不明のためとりあえずV7の最大5データを設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_relfile','fullpath','修繕関連ファイル情報','ファイルパス','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_relfile','addtime','修繕関連ファイル情報','追加日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_relfile','biko','修繕関連ファイル情報','備考','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_relfile','history','修繕関連ファイル情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "修繕情報-修繕メモ情報":          // 20160621 修繕関連移行処理追加 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_memo','bk_guid','修繕メモ情報','物件No','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_memo','szen_no','修繕メモ情報','修繕No','1','100','','2','','','最大値が不明のためとりあえず最大値100を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_memo','memo_no','修繕メモ情報','メモNo','1','50','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_memo','memo','修繕メモ情報','内容','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('szendata_memo','history','修繕メモ情報','履歴','','-1','','','','',''); ");
                        break;
                    }

                case "初期設定-初期設定基本情報":    // 20160616 初期設定情報移行処理実装 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','define_no','初期設定','No','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','applicationdefine','初期設定','各初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','aidagaradefine','初期設定','間柄初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','nkinzkseidefine','初期設定','入金項目属性初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','zeidefine','初期設定','税初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','hiwarinkindefine','初期設定','日割り入金初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','sorulemaitukidefine','初期設定','送金ルール毎月初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','sorulehendodefine','初期設定','送金ルール変動費初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','sorulekydefine','初期設定','送金ルール契約初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','sorulekodefine','初期設定','送金ルール更新初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','kojorulekydefine','初期設定','控除ルール契約初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','kojorulekodefine','初期設定','控除ルール更新初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','kojorulemaitukidefine','初期設定','控除ルール毎月初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','kojorulehendodefine','初期設定','控除ルール変動費初期値','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','history','初期設定','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('profile_fk','sorulesonotasqdefine','初期設定','送金ルール送金No初期値','','-1','','','','',''); ");
                        break;
                    }


                case "物件データ連動情報-送信設定基本情報":    // 20160720 連動情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','setting_guid','送信設定基本情報','送信設定guid','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','setting_sortorder','送信設定基本情報','送信設定順','0','9','','1','','','V7の最大値を設定'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','setting_name','送信設定基本情報','送信設定名','','200','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','keisai_siten','送信設定基本情報','掲載支店','0','999999','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','datalink_id','送信設定基本情報','連動ID','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','adconfirm_limitday','送信設定基本情報','広告確認からの確認期間','','4','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','history','送信設定基本情報','履歴','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','basedate','送信設定基本情報','基準日','1900/01/01','2100/12/31','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','basedate_flg','送信設定基本情報','基準日区分','1','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','emailaddr1','送信設定基本情報','Eメールアドレス1','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','emailaddr2','送信設定基本情報','Eメールアドレス2','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','emailaddr3','送信設定基本情報','Eメールアドレス3','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','emailaddr4','送信設定基本情報','Eメールアドレス4','','256','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_sendsetting','emailaddr5','送信設定基本情報','Eメールアドレス5','','256','','','','',''); ");
                        break;
                    }
                case "物件データ連動情報-送信設定自社web情報":
                case "物件データ連動情報-送信設定HOMES情報":
                case "物件データ連動情報-送信設定athome情報":
                case "物件データ連動情報-送信設定SUUMO情報":    // 20160720 連動情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_site_sendsetting','sitesetting_guid','送信設定情報','送信設定サイトユニーク','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_site_sendsetting','setting_guid','送信設定情報','送信設定ユニーク','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_site_sendsetting','site_no','送信設定情報','サイトNo','10','40','','2','','','20160802時点で4ポータル'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_site_sendsetting','site_id','送信設定情報','ポータルサイトID','','100','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_site_sendsetting','site_sendumu','送信設定情報','サイト別送信有無','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_site_sendsetting','site_data','送信設定情報','送信設定内容','','-1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_site_sendsetting','site_password','送信設定情報','ポータルサイトパスワード','','100','','','','',''); ");
                        break;
                    }
                case "物件データ連動情報-広告補足自社web情報":
                case "物件データ連動情報-広告補足HOMES情報":
                case "物件データ連動情報-広告補足athome情報":
                case "物件データ連動情報-広告補足SUUMO情報":    // 20160720 連動情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kokoku','hy_guid','広告補足情報','部屋Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kokoku','site_no','広告補足情報','サイトNo','10','40','','2','','','20160802時点で4ポータル'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_kokoku','item','広告補足情報','項目','','-1','','','','',''); ");
                        break;
                    }
                case "物件データ連動情報-ポータル連動部屋分類情報":    // 20160720 連動情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hy_ruisite','hy_ruino','ポータル連動部屋分類情報','部屋分類No','1','9999','','1','','部屋分類',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hy_ruisite','site_no','ポータル連動部屋分類情報','サイトNo','10','40','','2','','','20160802時点で4ポータル'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('m_hy_ruisite','hy_ruisiteitemdata','ポータル連動部屋分類情報','サイト用部屋分類No','','100','','','','',''); ");
                        break;
                    }
                case "物件データ連動情報-部屋毎送信情報":    // 20160720 連動情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_sosin','hy_guid','部屋毎送信情報','部屋guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_sosin','site_no','部屋毎送信情報','ポータルサイトNo','10','40','','2','','','20160802時点で4ポータル'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_sosin','site_jisyano','部屋毎送信情報','自社サービスNo','0','10','','3','','','20160802時点で6自社webだが最大を10に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_sosin','select_no','部屋毎送信情報','選択No','1','10','','4','','','最大値が不明のため10を設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_sosin','komok_guid','部屋毎送信情報','項目値','','16','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_sosin','history','部屋毎送信情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件データ連動情報-BtoBグループ設定情報":    // 20160720 連動情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_btobgroup','hy_guid','BtoBグループ設定情報','部屋guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_btobgroup','group_no','BtoBグループ設定情報','BtoBプラグイングループNo','1','5','','2','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_btobgroup','display_kbn','BtoBグループ設定情報','グループ公開有無','0','1','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_btobgroup','history','BtoBグループ設定情報','履歴','','-1','','','','',''); ");
                        break;
                    }
                case "物件データ連動情報-地図表示詳細設定情報":    // 20160720 連動情報構築 -add
                    {
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_mapdisp','hy_guid','地図表示詳細設定情報','部屋Guid','','16','','1','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_mapdisp','site_no','地図表示詳細設定情報','外部サイトNo（FK）','10','40','','2','','','20160802時点で4ポータル'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_mapdisp','site_jisyano','地図表示詳細設定情報','自社サイトNo','0','10','','3','','','20160802時点で6自社webだが最大を10に設定しておく'); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_mapdisp','map_dispflg','地図表示詳細設定情報','地図上に表示フラグ','1','2','','','','',''); ");
                        rtn_qry.Append(" INSERT INTO " + CVDBINFO_DBNAME + " VALUES ('hydata_mapdisp','history','地図表示詳細設定情報','履歴','','-1','','','','',''); ");
                        break;
                    }


            }

            return rtn_qry.ToString();

        }

    }
}