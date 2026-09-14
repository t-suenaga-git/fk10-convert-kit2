using System.Collections.Generic;
using System.Linq;
using Converter10.Njc.Frm;

namespace Converter10
{
    static class MidTableModule
    {
        /// <summary>
        /// 中間ファイル紐づけ情報list の「既存中間ファイル」側と列構成が完全一致することを
        /// 確認済みのカテゴリ一覧(item名)。ここに載っている間だけ、CREATE TABLE文を
        /// MakeMidTable_Legacy のハードコードではなく、マッピングリストから自動生成する。
        /// 一致確認はスクリプトで全135カテゴリを突き合わせて検証済み(2026-09時点)。
        /// 新しいカテゴリをここに追加する場合は、既存のMakeMidTable_Legacyの定義と
        /// 列名・順序が完全一致することを必ず確認すること。
        /// </summary>
        private static readonly HashSet<string> AutoGenSupportedItems = new()
        {
            "エリアマスタ", "バス交通マスタ", "バス停マスタ", "ライフライン業者情報",
            "家主メモ情報", "家主基本情報", "家主固定控除情報", "家主口座情報",
            "家賃入金口座情報", "家賃保証業者メモ情報", "家賃保証業者基本情報", "学校区マスタ",
            "契約メモ情報", "契約基本情報", "契約契約者情報", "契約次回入金項目情報",
            "契約者メモ情報", "契約者基本情報", "契約者口座情報", "契約者照合用カナ情報",
            "契約者保証人情報", "契約車情報", "契約特約事項情報", "契約入居者情報",
            "契約入金項目情報", "契約保険情報", "契約保証人情報", "契約履歴情報",
            "施工業者情報", "施設保守業者情報", "自社メモ情報", "自社基本情報",
            "自社担当者情報", "修繕業者メモ情報", "修繕業者基本情報", "修繕業者口座情報",
            "振込手数料情報", "送金ルール基本情報", "送金ルール控除項目情報", "送金ルール送金先情報",
            "送金ルール入金項目情報", "仲介業者メモ情報", "仲介業者基本情報", "仲介業者口座情報",
            "特約マスタ", "部屋メモ情報", "部屋間取内訳情報", "部屋基本情報",
            "部屋共通セールスポイント情報", "部屋鍵情報", "部屋所有者情報", "部屋駐車場情報",
            "部屋特約情報", "部屋入金項目情報", "物件メモ情報", "物件基本情報",
            "物件近隣駐車場情報", "物件鍵情報", "物件交通情報", "物件所有者情報",
            "変動費設定内容", "変動費料金単価表", "保険業者メモ情報", "保険業者基本情報",
            "保険業者口座情報", "保険種類マスタ",
        };

        public static string MakeMidTable(string midsheetname, string replacetblname)
        {
            // midsheetname は "グループ名@#@項目名" 形式。項目名部分だけを取り出す
            string[] tmp_parts = midsheetname.Split(new[] { CommonModule.STR_SPLIT_1 }, System.StringSplitOptions.None);
            string itemName = tmp_parts[tmp_parts.Length - 1];

            if (AutoGenSupportedItems.Contains(itemName))
            {
                string autoGenColumns = MakeMidTable_FromMapping(itemName);
                if (autoGenColumns != null)
                {
                    return " CREATE TABLE " + replacetblname + "(" + autoGenColumns + ")";
                }
            }

            // 許可リスト対象外、またはマッピングリストに該当が無い場合は既存のハードコード定義を使用
            return MakeMidTable_Legacy(midsheetname, replacetblname);
        }

        /// <summary>
        /// 中間ファイル紐づけ情報list の「既存中間ファイル」側から、指定カテゴリの列定義を
        /// 自動生成する。該当エントリが1件も無い場合はnullを返す(呼び出し元でLegacyにフォールバック)。
        /// </summary>
        private static string MakeMidTable_FromMapping(string itemName)
        {
            var columns = MainFrmHelper.中間ファイル紐づけ情報list
                .Where(e => e.既存中間ファイル != null
                            && e.既存中間ファイル.TableName == itemName
                            && !string.IsNullOrEmpty(e.既存中間ファイル.ColumnName))
                .Select(e => e.既存中間ファイル.ColumnName)
                .Distinct()
                .ToList();

            if (columns.Count == 0)
            {
                return null;
            }

            return string.Join(", ", columns.Select(c => $"[{c}] VARCHAR(MAX)"));
        }

        private static string MakeMidTable_Legacy(string midsheetname, string replacetblname)
        {

            string tmp_sql = "";
            string tmp_sql_pre = " CREATE TABLE " + replacetblname + "(";
            string tmp_sql_post = ")";

            switch (midsheetname ?? "")
            {
                case "各マスタ情報@#@バス交通マスタ":
                    {
                        tmp_sql = tmp_sql + " [バス交通No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [バス会社名] VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@バス停マスタ":
                    {
                        tmp_sql = tmp_sql + " [バス交通No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [バス停No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [系統名] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [バス停名] VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@鍵タイトルマスタ":
                    {
                        // 20161028 物件/部屋鍵取得方法修正 -chg sta
                        // tmp_sql = tmp_sql & " [鍵No]  VARCHAR(MAX), "
                        // tmp_sql = tmp_sql & " [鍵名称]  VARCHAR(MAX) "
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    tmp_sql = tmp_sql + " [鍵区分]  VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [鍵No]  VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [鍵名称]  VARCHAR(MAX) ";
                                    break;
                                }
                        }

                        break;
                    }
                // 20161028 物件/部屋鍵取得方法修正 -chg end
                case "各マスタ情報@#@特約マスタ":
                    {
                        tmp_sql = tmp_sql + " [特約区分] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [特約No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [特約タイトル] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [特約詳細] VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@クレーム分類設定内容":
                    {
                        tmp_sql = tmp_sql + " [箇所・クレーム分類区分]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [分類No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [名称]  VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@契約分類マスタ":
                    {
                        tmp_sql = tmp_sql + " [契約分類No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [契約分類名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [更新・解約通知期間(ヶ月)]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [契約分類色]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [定期借地借家権契約扱い有無]  VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@保険種類マスタ":
                    {
                        tmp_sql = tmp_sql + " [保険種類No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [保険種類名] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [保険種類カナ] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [備考] VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@学校区マスタ":
                    {
                        tmp_sql = tmp_sql + " [県No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [市No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [学校区No] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [町] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [丁目] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [番地] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [小学校名] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [中学校名] VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@エリアマスタ":
                    {
                        tmp_sql = tmp_sql + " [エリアNo] VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [エリア名] VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@変動費設定内容":
                    {
                        tmp_sql = tmp_sql + " [ルールNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ルール名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他単位]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他単位SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [料金数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [分類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [端数処理]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [端数処理パターン]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口径]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [計上分類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税区分早見]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税区分単価]  VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@変動費料金単価表":
                    {
                        tmp_sql = tmp_sql + " [変動費ルールNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [行No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [使用量]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [料金１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [料金２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [料金３]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [料金４]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [料金５]  VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@備考タイトルマスタ":
                    {
                        tmp_sql = tmp_sql + " [備考区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考タイトル]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [使用有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [補助アイテム使用有無]  VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@備考入力補助リストマスタ":
                    {
                        tmp_sql = tmp_sql + " [備考区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考入力補助リストNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考入力補助リスト内容]  VARCHAR(MAX) ";
                        break;
                    }
                case "各マスタ情報@#@画像タイトルマスタ":
                    {
                        tmp_sql = tmp_sql + " [画像区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [画像No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [画像名]  VARCHAR(MAX) ";
                        break;
                    }

                case "自社情報@#@自社基本情報":
                    {
                        tmp_sql = tmp_sql + " [自社・支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社・支店名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社・支店名（SJIS）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社・支店カナ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社・支店名２]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社・支店名２（SJIS）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社・支店名３]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社・支店名３（SJIS）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [代表者役職]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [代表者名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [携帯TEL1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [携帯TEL2]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [Webアドレス]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [備考（基本情報）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [免許証番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [免許年月日]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [取引主任者登録番号（代表）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [取引主任者名（代表）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [全国賃貸不動産管理業協会会員番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸不動産経営管理士登録番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸不動産経営管理士名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸住宅管理業者登録番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体１]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体２]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体３]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体４]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体５]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体６]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体７]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体８]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体９]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加入団体１０]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [代表者名（SJIS）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [取引主任者名（代表）（SJIS）]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸不動産経営管理士名（SJIS）]  VARCHAR(MAX) ";
                        break;
                    }
                case "自社情報@#@自社口座情報":
                    {
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    // 汎用の場合は中間ファイルをそのまま仮テーブルにする
                                    tmp_sql = tmp_sql + " [自社・支店No] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [自社・支店口座No] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [金融機関No] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [金融機関店No] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [口座種別] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [口座番号] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [口座名義] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [口座名義カナ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ゆうちょ口座記号１] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ゆうちょ口座記号２] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ゆうちょ口座番号] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [備考(口座情報)] VARCHAR(MAX) ";
                                    break;
                                }
                        }

                        break;
                    }
                case "自社情報@#@自社担当者情報":
                    {
                        tmp_sql = tmp_sql + " [賃貸革命ログオンユーザNo]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸革命ログオンユーザ名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸革命ログオンユーザ名(SJIS)]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸革命ログオンユーザ名カナ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [賃貸革命ログオングループNo]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [携帯電話1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [FAX1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メールアドレス1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [携帯メールアドレス1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [宅建免許番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [担当者ログオン禁止フラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [使用しないフラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ログオンパスワード]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [windowsログオンユーザー名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [意見画面表示フラグ]  VARCHAR(MAX) ";
                        break;
                    }
                case "自社情報@#@自社メモ情報":
                    {
                        tmp_sql = tmp_sql + " [自社・支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "自社情報@#@振込依頼人情報":
                    {
                        tmp_sql = tmp_sql + " [振込依頼人No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振込依頼人名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [金融機関No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [金融機関支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座名義]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振込依頼人コード]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振込依頼人カナ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [送信ファイルパス]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ヲ変換フラグ(0:変換しない 1:変換する)]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [負担調整フラグ(0:調整しない 1:調整する)]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [総合振込データ生成フラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社口座No]  VARCHAR(MAX) ";
                        break;
                    }
                case "自社情報@#@振込手数料情報":
                    {
                        tmp_sql = tmp_sql + " [振込依頼人No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金額From]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金額To]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [同行同支店手数料(電信扱)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [同行同支店手数料(文書扱)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [同行他支店手数料(電信扱)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [同行他支店手数料(文書扱)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [他行手数料(電信扱)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [他行手数料(文書扱)]  VARCHAR(MAX) ";
                        break;
                    }
                case "自社情報@#@口座振替情報":
                    {
                        tmp_sql = tmp_sql + " [振替情報No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振替情報名称]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振替情報カナ名称]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振込依頼人]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振込依頼人カナ名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [加盟店No 依頼人番号と共用]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ジェイリース　入金区分　1:変更なし・・・・]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [金融機関No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [金融機関支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座名義]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [引落日]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [手数料　請求額]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [データ並び替え　0:物件番号と部屋番号　1:契約者番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [契約者Noを契約者番号に出力する  0:出力しない　1:出力する]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [改行コード(CRLF)出力　0:なし　1:あり]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [「ｦ」→「ｵ」変換　0:しない　1:する]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [送信ファイルパス名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [受信ファイルパス名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [請求のまとめ方　0:請求先単位　1:契約単位]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [滞納分請求　0:しない　1:する]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [手数料の請求入金管理をする　0:しない　1:する]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [未入金の手数料は次回請求に加える　0:加えない　1:加える]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行金融機関番号指定　不使用の場合はNULL]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [マルチヘッダー形式フラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行付加コード]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振替でのゆうちょ銀行コードに振込用変換を利用するフラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振替入金処理時に受信ファイルを読み込まないフラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [再振替対応利用フラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [振替先口座No]  VARCHAR(MAX) ";
                        break;
                    }
                case "自社情報@#@入出金取得情報":
                    {
                        break;
                    }
                case "自社情報@#@家賃入金口座情報":
                    {
                        tmp_sql = tmp_sql + " [家賃入金受付口座No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [受取人名]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [金融機関No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [金融機関支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [受取人名Unicode]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [受取人カナ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [請求先区分]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社支店No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [自社口座No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [家主No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [家主口座No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC有無]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC口座指定方式]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC種目付加方法]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC任意番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC加入者番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC接続先個別設定フラグ]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC接続方法]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPCエリアNo]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC地区No]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPCTEL]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPC照会用暗証番号]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ANSER-SPCサービスコード]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [家賃入金口座備考]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [ヘッダーテンプレート]  VARCHAR(MAX), ";
                        tmp_sql = tmp_sql + " [トレーラーテンプレート]  VARCHAR(MAX) ";
                        break;
                    }
                case "自社情報@#@ANSERエリア情報":
                    {
                        break;
                    }
                case "自社情報@#@ANSERアクセスポイント情報":
                    {
                        break;
                    }
                case "自社情報@#@ANSER接続情報":
                    {
                        break;
                    }

                case "家主情報@#@家主基本情報":
                    {
                        tmp_sql = tmp_sql + " [家主No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [個人法人フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家主名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家主名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家主カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [宛名敬称]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [登記住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [登記住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先電話設定]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先メール設定]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [誕生日・設立日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [年収・年商]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [性別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [本籍地]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先業種]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先部署]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先情報記入年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先入社年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [URL]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業種]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者を宛先に含める]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者を宛先に含める]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [資本金]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [従業員数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [記入年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [主要取引先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先宛名敬称]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先携帯１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先携帯２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先間柄]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先宛名敬称]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送付先備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [イベント担当者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [絞り込みキーワード]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [会計情報]  VARCHAR(MAX) ";
                        break;
                    }
                case "家主情報@#@家主口座情報":
                    {
                        tmp_sql = tmp_sql + " [家主No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関支店No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ口座記号１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ口座記号２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ口座番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込情報備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [総合振込区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込依頼人No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料負担区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料計算区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [総合振込情報備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家主向け帳票の表示]  VARCHAR(MAX) ";
                        break;
                    }
                case "家主情報@#@家主イベント情報":
                    {
                        break;
                    }
                case "家主情報@#@家主メモ情報":
                    {
                        tmp_sql = tmp_sql + " [家主No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }

                case "業者情報@#@仲介業者基本情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [Webアドレス(URL)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(メールアドレス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振込通知書発行の許可]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [免許証番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [免許年月日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [取引主任者登録番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [取引主任者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [全国賃貸不動産管理業協会会員番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸不動産経営管理士登録番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸不動産経営管理士名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸住宅管理業者登録番号(仮)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [取引主任者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介業者契約担当者自動入力フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸不動産経営管理士名(SJIS)]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@仲介業者口座情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関支店No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(口座情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(口座振込情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [総合振込区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込依頼人No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料負担区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料計算区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(総合振込情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行記号1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行記号2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行番号]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@仲介業者メモ情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@保険業者基本情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [Webアドレス(URL)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(メールアドレス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@保険業者口座情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関店No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(口座情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(口座振込情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [総合振込区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込依頼人No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料負担区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料計算区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(総合振込情報)]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@保険業者メモ情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@家賃保証業者基本情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名(shift-jis)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [Webアドレス(URL)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(メールアドレス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連動用家賃保証会社区分]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@家賃保証業者メモ情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@修繕業者基本情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [Webアドレス(URL)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(メールアドレス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名(SJIS)]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@修繕業者口座情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関店No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行記号1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行記号2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ銀行番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(口座情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(口座振込情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [総合振込区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込依頼人No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料負担区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料計算区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(総合振込情報)]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@修繕業者メモ情報":
                    {
                        tmp_sql = tmp_sql + " [業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@ライフライン業者情報":
                    {
                        tmp_sql = tmp_sql + " [ライフライン業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン業者名(Shift-jis)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン業者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [営業所]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [営業所(Shift-jis)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [電気]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [上水道]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ガス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [灯油]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [排水]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@施工業者情報":
                    {
                        tmp_sql = tmp_sql + " [施工業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [施工業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [施工業者名(Shift-jis)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [施工業者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [営業所]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [営業所(Shift-jis)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [Webアドレス(URL)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(メールアドレス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX) ";
                        break;
                    }
                case "業者情報@#@施設保守業者情報":
                    {
                        tmp_sql = tmp_sql + " [施設保守業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [施設保守業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [施設保守業者名(Shift-jis)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [施設保守業者名カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [営業所]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [営業所(Shift-jis)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [Webアドレス(URL)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(メールアドレス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX) ";
                        break;
                    }

                case "物件情報@#@物件基本情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [建物識別コード]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [都道府県コード]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [市区町村コード]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [町地域]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [丁番地]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [街区番号地番]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件外部No]  VARCHAR(MAX) ";
                        break;
                    }
                case "物件情報@#@物件詳細情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件分類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [緯度]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [経度]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [緯度(日本測地系)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [経度(日本測地系)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [電鉄物件フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [竣工日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [地上階建て]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [地下階]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [エレベータフラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [エレベータ数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [建物構造(基本)-構造]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [建物構造(基本)-屋根構造]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [情報元業者NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [総戸数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-延床面積]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-延床面積坪数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-敷地面積]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-敷地面積坪数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-駐車面積]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-駐車面積坪数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-延床面積(登記)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-延床面積坪数(登記)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-敷地面積(公簿)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-敷地面積坪数(公簿)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金口座-請求月初期値]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約書用入金締切日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家賃持参先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金口座-家賃入金口座No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金口座-契約金用入金口座の有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金口座-契約金用入金口座No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐火構造区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [施工会社No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保守業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理方式]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理業者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理業者担当者]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理業者電話番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理業者業務形態]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理人電話番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [支店No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他交通１-その他交通]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他交通１-距離]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他交通２-その他交通]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他交通２-距離]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-電気-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-上水-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-ガス-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-排水-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-灯油-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-その他１-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-その他２-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ライフライン-その他３-公共機関No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [検針業務の有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [検針登録の並び順]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [親子メーター変動費の使用有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場-付随駐車場有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場-自動車台数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場-バイク台数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場-自転車台数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場-自転車利用自由フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有者-一棟・所有区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-調査の有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-調査結果の問合せ先-所有者]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-調査結果の問合せ先-管理組合]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-調査結果の問合せ先-管理業者]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-調査結果の問合せ先-施工業者]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-調査年月日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-実施機関]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-調査範囲]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-使用有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-使用箇所]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [石綿使用調査-備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-診断有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-診断記録の問合せ先-所有者]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-診断記録の問合せ先-管理組合]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-診断記録の問合せ先-管理業者]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-耐震基準適合証明書の写し]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-住宅性能評価所の写し]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-耐震診断結果の写し]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [耐震診断-備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-土砂災害警戒地域内外]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-法令分類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-法令内容]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [敷地利用-敷地利用種類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [敷地利用-契約期間開始]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [敷地利用-契約期間終了]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [敷地利用-備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居率一覧対象開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎に業者が異なる場合フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゴミ出しに関する補足情報]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [登記情報の日付]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有権にかかる権利有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有権にかかる権利の種類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有権以外の権利有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-建築面積]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-建築面積坪数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-建築面積(登記)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件面積-建築面積坪数(登記)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [自社担当者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理業者FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [エレベータフラグ(家)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-土砂災害特別警戒区域内外]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-造成宅地防災区域内外]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-津波災害警戒区域内外]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-土砂災害警戒地域備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-土砂災害特別警戒区域備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-造成宅地防災区域備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [法令-津波災害警戒区域備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理人名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [その他建物構造]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [角地フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [都市計画]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [用途地域]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理人名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理業者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態-管理業者担当者名(SJIS)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場空き台数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バイク置き場空き台数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐輪場空き台数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [小学校区]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [小学校距離]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [中学校区]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [中学校距離]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [エリア]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [非常用エレベーターフラグ]  VARCHAR(MAX) ";
                        break;
                    }
                case "物件情報@#@物件所有者情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [貸主１No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有者１No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [貸主２No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有者２No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有期間開始]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有期間終了]  VARCHAR(MAX) ";
                        break;
                    }
                case "物件情報@#@物件交通情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [沿線番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [沿線No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駅No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [始発フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [距離]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [徒歩]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [車]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス会社]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス停]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス停徒歩]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス停までの距離]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [設定する交通のタイプ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス系統・路線名(駅までバスを利用)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス停名(駅までバスを利用)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス停(駅までバスを利用)までの道路距離]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バス停(駅までバスを利用)までの徒歩(分)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [車距離]  VARCHAR(MAX) ";
                        break;
                    }
                case "物件情報@#@物件メモ情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "物件情報@#@物件鍵情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [鍵No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [鍵タイトル名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [鍵本数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保管場所]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者間での情報共有]  VARCHAR(MAX) ";
                        break;
                    }
                case "物件情報@#@物件近隣駐車場情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [距離(m)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [料金(月額)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税区分]  VARCHAR(MAX) ";
                        break;
                    }
                case "物件情報@#@物件ゴミ情報":
                    {
                        break;
                    }
                case "物件情報@#@物件権利情報":
                    {
                        break;
                    }
                case "物件情報@#@物件接道情報":
                    {
                        break;
                    }
                case "物件情報@#@物件周辺情報":
                    {
                        break;
                    }
                case "物件情報@#@物件修繕維持管理連絡先情報":
                    {
                        break;
                    }

                case "部屋情報@#@部屋基本情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋分類名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [間取り]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [間取り備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [専有実面積(m2)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [専有実面積(坪)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [床面積(m2)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [床面積(坪)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [専有登記面積(m2)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [専有登記面積(坪)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [登記延床面積(m2)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [登記延床面積(坪)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バルコニー面積(m2)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バルコニー面積(坪)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [店舗付き住宅店舗部分面積(m2)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [店舗付き住宅店舗部分面積(坪)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [店舗付き住宅住宅部分面積(m2)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [店舗付き住宅住宅部分面積(坪)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付月数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付日数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付日にち]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [自社Web内部キー採番ID]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋詳細情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋業務期間有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所在階１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所在階２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所在階３]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [地下フラグ１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [地下フラグ２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [地下フラグ３]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [向き]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [角部屋フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バルコニー方向]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [状況]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居状況備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居可能日を時期で指定]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居可能時期(年月)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居可能時期(上旬・中旬・下旬)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居条件]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [広告内容確認日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [募集可能種別(しない・する)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [フリーレント有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [フリーレントカ月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [フリーレント詳細]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険-保険の利用]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険-保険期間]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険-保険料]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険-保険の備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [取引形態種別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [客付け状態可否]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [手数料負担割合貸主(%)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [手数料負担割合借主(%)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [手数料配分元付(%)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [手数料配分先物(%)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [客付会社への物件コメント]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者間広告広告活動種別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [情報元業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [広告料有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [広告料上限額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [広告料条件内容]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋住所の変更フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋住所(丁目)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋住所(丁目区分)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋住所(町地域)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋住所(その他)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証会社]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証内容]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約の種類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [定期借家契約(期間)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [定期借家契約(期日)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [バイク駐車場備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐輪場備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [共通セールスポイント]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [不動産検索サイト掲載設定-物件名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [不動産検索サイト掲載設定-部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [不動産検索サイト掲載設定-丁番地以下]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [不動産検索サイト掲載設定-地図上]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸保証利用区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [広告料上限額区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [広告料上限率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [広告料上限額税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [共通セールスポイント使用フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [複数階有りのフラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [支店NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [自社担当者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [次回更新時の初期値]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [登記情報の日付]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有権にかかる権利有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有権にかかる権利の種類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有権以外の権利有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [BtoBプラグイングループ設定区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [自社Webオススメ物件表示]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [退去日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [使用目的]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [新築区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居時期区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [取引自社区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約期間区分]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋所有者情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [貸主１No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有者１No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [貸主２No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有者２No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有期間開始]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所有期間終了]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋駐車場情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場の空き数(手動設定用)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場の空き有無(手動設定用)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場料金区分(手動設定用)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場料金(手動設定用)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場料金税区分(手動設定用)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車場の賃貸可能数]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋特約情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [特約区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容50]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋鍵情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [鍵No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [鍵タイトル名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [鍵本数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保管場所]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業者間での情報共有]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋間取内訳情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [間取り内訳No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [間取り内訳区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [畳数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [所在階]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [既存間取内訳データ]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋設備情報":
                    {
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    tmp_sql = tmp_sql + " [物件NO] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [部屋NO] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-オール電化] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-介護付き] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-フロントサービス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-高層階] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-耐震扉] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-耐震補強工事済] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-日当たり良好] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-温泉引き込み済み] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-オーシャンビュー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-前面棟無] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-眺望良好] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-通風良好] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-花火大会鑑賞] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-リゾート向き] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-タワー型マンション] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-デザイナーズ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-分譲賃貸] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-スマートハウス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-隣接建物距離2m以上] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-3方角住宅] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-4方角住宅] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-店舗付住宅] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-事務所付住宅] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-IT重説対応物件] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-DIY可] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-初期費用] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-学生敷金] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-学生礼金] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-更新料不要] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-仲介手数料不要] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-初期費用クレジット決済可] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [アピールポイント-賃料クレジット決済可] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-II型キッチン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-L字型キッチン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-U字型キッチン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-オーブン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-オープンキッチン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-ガスオーブン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-ガスオーブンレンジ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-ガスキッチン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-ガスレンジ] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [キッチン-ガス種類] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [ライフライン-ガス種類] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [キッチン-ガラストップ(コンロ設備)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-キッチンタイプ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-キッチン出入り口] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-キッチンに窓] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-キッチン床暖房] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-キッチン足元温風器] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-キッチン未使用] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-グリル] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-コンロ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-コンロの口数] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-共同キッチン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-ディスポーザー] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [キッチン-都市ガス引き込み] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [ライフライン-都市ガス引き込み] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [キッチン-ミニキッチン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-浄水器] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-食器乾燥機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-食器洗い乾燥機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-食器棚(カップボード)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-電気クッカー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-電子レンジ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-湯沸器] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [キッチン-備付食器棚] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -del
                                    // tmp_sql = tmp_sql & " [キッチン-冷蔵庫] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [構造・間取り-雨戸] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-玄関ポーチ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-玄関ホール] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-2ドア1ルーム] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [構造・間取り-クッションフロア] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-クッションフロア] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [構造・間取り-じゅうたん張] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-じゅうたん張] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [構造・間取り-スキップフロア] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-南面リビング] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-フラワーボックス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-フローリング] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-南向き部屋数] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-メーターモジュール] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-リビングの隣和室] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-リビング階段] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-リビング吹抜け] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-可動間仕切] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-家事室] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-活性炭敷込み] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-掘りごたつ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-勾配天井] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-腰壁] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-室内らせん階段] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-出窓] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-勝手口] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-床の間] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-昇降機付階段] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-畳コーナー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-振り分け] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-吹抜] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-折上天井] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-続き和室] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-茶室] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-天窓] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-土間] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-内階段] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-内装コンクリート] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-風除室] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-防音設備] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-琉球畳] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-シャッター] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-2階LDK] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-2階リビング] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-自然素材] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-大黒柱] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-無垢材使用] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-四寸柱] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-灯油配管] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-天井高2.5m以上] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-天井高2.7m以上] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-天井高3m以上] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-スケルトンインフィル] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-全居室洋室] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-メゾネット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-1フロア1住戸] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-1フロア2住戸] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-採光] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [構造・間取り-採光面数] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-24時間緊急対応] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-インターホン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-オートライト] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-カードキー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-玄関チャイム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-玄関リモコンキー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-24時間セキュリティーシステム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-ダブルロックキー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-ディンプルキー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-電子キー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-ドアタッチキー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-火災警報器(報知機)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-機械警備] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-携帯電話システム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-指紋認証] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-人感照明センサー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-静脈認証] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-電子ロック] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-非接触型ICカードキー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-防犯カメラ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-防犯ガラス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-防犯シャッター] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-24時間有人管理] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-オートロック] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-セキュリティ会社加入] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-24時間緊急通報システム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [セキュリティ-防犯モデルマンション] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-車庫] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-ハイルーフ可] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-大型バイク可] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-EV車充電設備] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-カーポート] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-自走式タワー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車場地面] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-平置き] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-ビルトガレージ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-車寄せ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-大型車庫] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車場シャッター付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車場リモコンシャッター] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車場屋外] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車場屋根付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車場屋内] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車場地下] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車並列2台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車並列3台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-駐車並列4台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-電動シャッター車庫] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [駐車場-来客用パーキング] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-BS] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-CATV] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-CATV使用料] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-CS] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-LAN配線] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-インターネット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-インターネット使用料] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-パラボラ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [テレビ・通信-マルチメディアコンセント] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [テレビ・通信-有線放送] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [事業用-有線放送] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [入居条件-2世帯] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-SOHO] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-家族] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-障害者] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-ピアノ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-ペット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-大型犬] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-小型犬] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-猫] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-留学生] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-ルームシェア] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-外国人] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-学生] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-楽器] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-高齢者] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-子供] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-事務所] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-性別] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-単身者] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-二人入居] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-保証人] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-法人] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-シェアハウス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [入居条件-喫煙者] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-バランス釜] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-3点給湯] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-エコウィル] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-エコキュート] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-エコジョーズ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-大型給湯機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-オーディオバス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-オートバス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-高温差し湯式バス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-サウナ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-ジェットバス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-シャワー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-シャワールーム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-タンクレストイレ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-トイレ種類] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-トイレ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-トイレ共同専用の別] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-トイレ未使用] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-凍結防止付給湯] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-バストイレ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-バス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-バス共同専用の別] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-パッシブ換気] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-ボイラー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-ミストサウナ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-温水洗浄便座] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-給湯] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-除湿機能付換気扇] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-脱衣所] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-暖房便座] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-追い焚き給湯] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-灯油ボイラー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-浴室1坪以上] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-浴室テレビ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-浴室に窓] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-浴室乾燥機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-浴室床暖房] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-浴室未使用] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-檜バス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バス・トイレ-温泉付き] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バリアフリー-アルコーブ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バリアフリー-バリアフリー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [バリアフリー-フラットフロア] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -del
                                    // tmp_sql = tmp_sql & " [バリアフリー-高齢者専用機能] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [ベランダ-ウッドデッキ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-バルコニー出入り口] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-縁側] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-両面バルコニー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-バルコニーの面数] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-L字型バルコニー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-インナーバルコニー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-サービスバルコニー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-テラス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-ルーフバルコニー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-ベランダ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ベランダ-ワイドバルコニー] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [室内設備・家電・家具-24時間換気システム] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [構造・間取り-24時間換気システム] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [室内設備・家電・家具-FAX付] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [事業用-FAX付] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖房(ガスFF暖房)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖房(灯油FF暖房)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-低放射ガラス] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [室内設備・家電・家具-アプローチライト] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [付帯施設・設備-アプローチライト] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-エアコン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-大型木目調家具] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖房(温水暖房)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-カーテン付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-カーペット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ガスストーブ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖房(ガス暖房)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-シーリングファン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-樹脂サッシ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ダウンライト] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-テレビ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-電話2回線] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-電話配線] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-トップライト付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-二重サッシ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ハイサッシ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ビルトインエアコン] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ファンコンベクタ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-フットライト] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ブラインド付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ベッド] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-ロールスクリーン付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-家具付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-家電付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-家具・家電付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-乾燥機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-間接照明] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-玄関・階段手すり] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-収納ベッド] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-床下換気] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-床暖房] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-照明] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-親子扉] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖房(石油暖房)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-洗濯機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-洗濯機置き場] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖炉] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖房(蓄熱式暖房)] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-電動室内物干機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-電話] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-灯油] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-熱交換集中換気システム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-複層ガラス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-壁面ベッド] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-防音サッシ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-冷房] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-暖房(集中暖房)] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -del
                                    // tmp_sql = tmp_sql & " [室内設備・家電・家具-電気容量(契約アンペア)] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [収納-ウォークインクローゼット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-ウォークスルークローゼット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-大型シューズボックス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-クローゼット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-シューズウォークインクローゼット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-シューズボックス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-シューズインクローゼット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-ストックヤード] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-全居室収納] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-タイヤ置き場] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-トランクルーム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-押入] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-屋根裏収納] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-階段下収納] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-玄関収納] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-収納] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-収納広さ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-床下収納] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-昇降ウォール収納] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-食品庫] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-倉庫] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-耐震ラッチ吊戸棚] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-天井高シューズクローゼット] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-天袋] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-物置] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [収納-物入] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [水道・排水-水道] VARCHAR(MAX), "
                                    // tmp_sql = tmp_sql & " [水道・排水-排水] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [ライフライン-水道] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [ライフライン-排水] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [洗面-2階洗面台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗髪洗面化粧台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-三面鏡付洗面化粧台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-縦型照明付洗面化粧台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面2ボウル] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面化粧台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面所] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面所出入り口] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面所にドア] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面所に窓] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面所独立] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面台の共有] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [洗面-洗面台] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-AED付] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-屋上] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-屋上庭園] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-ガーデニング専用機能] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-クリーニングボックス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-ゲストルーム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-シャトルバス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-スロップシンク] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-ソーラーシステム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-坪庭] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-中庭] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-庭] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-庭坪数] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-フリースペース] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-南庭] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-屋外電源] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-専用庭] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-宅配ボックス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-アスレチックルーム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-敷地内遊び場] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-エントランス] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-キッズルーム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-シアタールーム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-パーティールーム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-コインランドリー] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-大浴場] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-託児所] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-ドライエリア] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-離れ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-フィットネス施設] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-プール] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-ペット専用設備] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-融雪機] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-融雪槽] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-ラウンジ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-ロードヒーティング] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-集合郵便受け] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-地下室] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [付帯施設・設備-集会場] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [事業用-OAフロア] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [事業用-Pタイル] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [室内設備・家電・家具-Pタイル] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    tmp_sql = tmp_sql + " [事業用-専用階段] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [事業用-事務所仕上げ] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [事業用-スケルトン渡し] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [事業用-入退去管理システム] VARCHAR(MAX), ";
                                    tmp_sql = tmp_sql + " [事業用-大型車入庫] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg sta
                                    // tmp_sql = tmp_sql & " [事業用-飲食店・業] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [入居条件-飲食店・業] VARCHAR(MAX), ";
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -chg end
                                    // 20170606 部屋設備情報の仕様変更(グループ名変更)修正対応 -del
                                    // tmp_sql = tmp_sql & " [事業用-24時間利用可] VARCHAR(MAX), "
                                    tmp_sql = tmp_sql + " [事業用-空調方式] VARCHAR(MAX) ";
                                    break;
                                }
                        }

                        break;
                    }
                case "部屋情報@#@部屋入金項目情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目行No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出基準入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出ヶ月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求先No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求開始月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求パターン]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [固定公共料金で使用]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求発生年]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求発生月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋メモ情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋共通セールスポイント情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [パーツNO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [セールスポイントパーツ]  VARCHAR(MAX) ";
                        break;
                    }
                case "部屋情報@#@部屋面積情報":
                    {
                        break;
                    }
                case "部屋情報@#@部屋変動費各戸メーター情報":
                    {
                        break;
                    }
                case "部屋情報@#@部屋修繕維持管理連絡先情報":
                    {
                        break;
                    }

                case "送金ルール情報@#@送金ルール基本情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金ルール管理No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一所有形態区分-棟/区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金ルール適用開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金ルール適用終了日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理形態]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一括借上 一部管理]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介物件 - 新規契約業務]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介物件 - 契約更新業務]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介物件 - 解約業務]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法1 - 該当年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法1 - 締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法1 - 送金月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法1 - 送金締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法2 - 該当年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法2 - 締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法2 - 送金月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法2 - 送金締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法3 - 該当年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法3 - 締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法3 - 送金月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法3 - 送金締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法4 - 該当年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法4 - 締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法4 - 送金月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法4 - 送金締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法5 - 該当年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法5 - 締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法5 - 送金月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金日決定方法5 - 送金締日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金先単独/複数指定]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金固定額使用フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金固定数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [均等案分フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [端数受取先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金額案分端数調整フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一括借上 物件毎/部屋毎]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一括借上 物件毎設定額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一括借上 免責期間の設定フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一括借上 免責期間(解約翌月から何カ月)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料徴収区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料 例外：契約金は対象外とする]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料 例外：解約金は対象外とする]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料定額区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎定額 - 全部屋一律管理手数料]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎定額 - 部屋毎日割りフラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [物件毎定額 - 物件管理手数料]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料 消費税適用フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考 - 基本情報]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [支払明細書関連 - 同時契約の場合は代表する部屋を表示]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [支払明細書関連 - 口座毎に支払明細書を作成1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [支払明細書関連 - 口座毎に支払明細書を作成2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎一括借上額 - 入金項目No1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎一括借上額 - 入金項目No2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎一括借上額 - 入金項目No3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎一括借上額 - 入金項目No4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋毎一括借上額 - 入金項目No5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料計算基準区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [満額入金-送金フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介物件 - 随時送金]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一括借上 部屋毎一括借上額の計算方法]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [一括借上 消費税適用フラグ]  VARCHAR(MAX) ";
                        break;
                    }
                case "送金ルール情報@#@送金ルール送金先情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金ルール管理No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金先行No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金先家主NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金先口座NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [固定送金額]  VARCHAR(MAX) ";
                        break;
                    }
                case "送金ルール情報@#@送金ルール入金項目情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金ルール管理No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [月々/契約時/更新時区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [行NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理手数料率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [滞納保証有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金先No]  VARCHAR(MAX) ";
                        break;
                    }
                case "送金ルール情報@#@送金ルール控除項目情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金ルール管理No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約時/更新時区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [行NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [相殺予定フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除額基準]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除額税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除対象入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除率税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除率税有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用税率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [立替回収または預り金]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除税額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金対象フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金対象入金項目名]  VARCHAR(MAX) ";
                        break;
                    }

                case "契約者情報@#@契約者基本情報":
                    {
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約者名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [個人法人フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [宛名敬称]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯メールアドレス]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(メールアドレス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [生年月日・設立日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [年収・年商]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(基本情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [性別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [本籍地]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先業種]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先部署]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先情報記入年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先入社年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居前連絡先情報記入年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居前連絡先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居前連絡先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居前連絡先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居前連絡先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居前連絡先TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居前連絡先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [Webアドレス(URL)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [業種]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [代表者を宛先に含める]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者部署]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者役職]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [担当者を宛先に含める]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [資本金]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [従業員数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [記入年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [主要取引先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先宛名敬称]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先携帯１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先携帯２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先間柄]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先宛名敬称]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類送付先備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込通知書発行の可否]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振替通知書発行の可否]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [督促状発行の可否]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振込仮想口座使用フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証人複数フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振替の保証有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [コンビニ収納サービス利用有無]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約者情報@#@契約者口座情報":
                    {
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約者口座No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融期間No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関支店No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座種別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座名義カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ記号１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ記号２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ゆうちょ口座番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振替No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振替手数料]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振替契約者番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振込備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [総合振込区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込依頼人No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料負担区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料計算区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [振込手数料固定額2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(総合振込情報)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [印刷口座区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [金融機関区分(ゆうちょ銀行フラグ)]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約者情報@#@契約者メモ情報":
                    {
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約者情報@#@契約者照合用カナ情報":
                    {
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [照合用カナ10]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約者情報@#@契約者保証人情報":
                    {
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証人No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証人名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証人名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証人カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [宛名敬称]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [優先設定(電話番号)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [誕生日・設立日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [年収・年商]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [間柄]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先郵便番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先住所１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先住所２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先TEL１]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先TEL２]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先FAX]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先業種]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先部署]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先情報記入年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先入社年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先退社年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先備考]  VARCHAR(MAX) ";
                        break;
                    }

                case "契約情報@#@契約基本情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [初回契約日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約状況(ステータス)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [キャンセル理由]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ステータス変更日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [取引主任者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [手付預り額①]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [手付預り日①]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [手付預り備考①]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [サービス分類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [会計グループ分類]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付月数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付日数]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [解約受付日にち]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介担当者名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [仲介担当者名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約終了日]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約履歴情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [更新No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [改定No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約終了日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [条件変更日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約更新通知日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [通知書印刷日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [管理業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸保証業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸保証内容]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [賃貸保証顧客番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約分類名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [使用目的]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約更新業務有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家賃入金区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振替開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振替開始待ちフラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家賃入金口座区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約一時金入金口座No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [毎月分入金口座No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [翌月分受取り有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [当月分の差額受取り有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [次回契約開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [次回契約終了日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [更新時日割り有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金控除ルール適用有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金予定日使用フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [送金予定日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋固定管理手数料フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋固定管理手数料額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [次回更新設定区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求データ作成開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求データ作成終了日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考2(基本)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考(保険)]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金締め日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家賃持参先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険期間]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [身元引受人名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [身元引受人住所]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [身元引受人連絡先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険料入金日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [更新完了日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [更新完了通知日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [更新通知日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [催促実施日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [書類返送受取日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [毎月分請求有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [身元引受人名SJIS]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [口座振替日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [翌月分受取り月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約担当者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求担当者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [配分割合_元付]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [配分割合_客付]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [客付会社の手数料額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [客付会社の手数料税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [客付会社の手数料税額]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約車情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [車情報No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メーカー]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [車名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [車色]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [ナンバー]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [駐車区画]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約契約者情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [並び順No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居フラグ]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約保証人情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [並び順]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保証人No]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約入居者情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入居者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [氏名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [氏名Unicode]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [カナ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [性別]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [続柄]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [生年月日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [勤務先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [連絡先]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [携帯電話番号]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約特約事項情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [特約区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容50]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容51]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容52]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容53]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容54]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容55]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容56]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容57]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容58]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容59]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容60]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容61]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容62]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容63]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容64]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容65]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容66]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容67]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容68]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容69]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容70]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容71]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容72]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容73]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容74]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容75]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容76]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容77]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容78]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容79]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容80]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容81]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容82]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容83]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容84]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容85]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容86]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容87]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容88]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容89]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容90]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容91]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容92]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容93]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容94]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容95]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容96]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容97]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容98]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容99]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容100]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [内容101]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約保険情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険種類No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険業者No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用開始年月日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用終了年月日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [保険金額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [満期案内通知有無]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [証券番号]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約メモ情報":
                    {
                        tmp_sql = tmp_sql + " [物件NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約NO]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ1]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ2]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ3]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ4]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ5]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ6]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ7]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ8]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ9]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ10]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ11]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ12]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ13]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ14]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ15]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ16]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ17]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ18]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ19]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ20]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ21]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ22]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ23]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ24]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ25]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ26]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ27]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ28]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ29]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ30]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ31]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ32]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ33]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ34]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ35]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ36]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ37]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ38]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ39]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ40]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ41]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ42]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ43]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ44]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ45]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ46]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ47]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ48]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ49]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [メモ50]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約入金項目情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [並び順No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求税額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出基準入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出ヶ月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求先No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金方法]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [フリーレント適用開始日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [フリーレント適用終了日]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [フリーレント終了月請求額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求開始月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求パターン]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求間隔]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求発生年]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求発生月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用税率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [フリーレント適用区分]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約次回入金項目情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [契約管理レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目レコードNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [並び順No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求税額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出基準入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [算出ヶ月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求先No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金方法]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求開始月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求パターン]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求発生間隔]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求対象年]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求対象月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用税率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX) ";
                        break;
                    }
                case "契約情報@#@契約変動費各戸メーター情報":
                    {
                        break;
                    }
                case "契約情報@#@契約控除ルール情報":
                    {
                        break;
                    }
                case "契約情報@#@契約送金ルール情報":
                    {
                        break;
                    }
                case "契約情報@#@契約解約情報":
                    {
                        break;
                    }
                case "契約情報@#@契約修繕見積情報":
                    {
                        break;
                    }
                case "契約情報@#@契約修繕見積詳細情報":
                    {
                        break;
                    }
                case "契約情報@#@契約変動費親メーター情報":
                    {
                        break;
                    }
                case "契約情報@#@契約解約確認事項情報":
                    {
                        break;
                    }
                case "契約情報@#@契約同時契約情報":
                    {
                        break;
                    }
                case "契約情報@#@契約原状回復目安単価情報":
                    {
                        break;
                    }
                case "契約情報@#@契約敷金保証金随時処理情報":
                    {
                        break;
                    }
                case "契約情報@#@契約関連ファイル情報":
                    {
                        break;
                    }
                case "契約情報@#@契約空室待ち情報":
                    {
                        break;
                    }

                case "請求情報@#@家主固定控除情報":
                    {
                        tmp_sql = tmp_sql + " [物件No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目ソートNo]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金項目名]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [部屋No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [相殺フラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [税額]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [入金区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [家賃入金口座No]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用開始該当年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用終了該当年月]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [立替・預りフラグ]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除請求月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除請求発生パターン]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除請求発生間隔]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除請求発生月指定年区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [控除請求発生月指定月区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [適用税率]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [備考]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [請求締区分]  VARCHAR(MAX),";
                        tmp_sql = tmp_sql + " [領収フラグ]  VARCHAR(MAX) ";
                        break;
                    }
                case "請求情報@#@預り金情報":
                    {
                        break;
                    }
                case "請求情報@#@未収滞納金情報":
                    {
                        break;
                    }
                case "請求情報@#@その他請求情報":
                    {
                        break;
                    }
                case "請求情報@#@変動費検針情報":
                    {
                        break;
                    }
                case "請求情報@#@家主請求控除情報":
                    {
                        break;
                    }

                case "クレーム情報@#@クレーム基本情報":
                    {
                        break;
                    }
                case "クレーム情報@#@クレーム対応履歴情報":
                    {
                        break;
                    }
                case "クレーム情報@#@クレーム関連ファイル情報":
                    {
                        break;
                    }
                case "修繕情報@#@修繕基本情報":
                    {
                        break;
                    }
                case "修繕情報@#@修繕見積情報":
                    {
                        break;
                    }
                case "修繕情報@#@修繕見積詳細情報":
                    {
                        break;
                    }
                case "修繕情報@#@修繕クレーム関連付け情報":
                    {
                        break;
                    }
                case "修繕情報@#@修繕関連ファイル情報":
                    {
                        break;
                    }
                case "修繕情報@#@修繕メモ情報":
                    {
                        break;
                    }
                case "初期設定@#@初期設定基本情報":
                    {
                        break;
                    }
                case "初期設定@#@税編集情報":
                    {
                        break;
                    }
                case "初期設定@#@変換文字情報":
                    {
                        break;
                    }
                case "初期設定@#@入金項目集約情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@送信設定基本情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@送信設定自社web情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@送信設定HOMES情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@送信設定athome情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@送信設定SUUMO情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@地図表示詳細設定情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@BtoBグループ設定情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@部屋毎送信情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@ポータル連動部屋分類情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@広告補足SUUMO情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@広告補足athome情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@広告補足HOMES情報":
                    {
                        break;
                    }
                case "物件データ連動情報@#@広告補足自社web情報":
                    {
                        break;
                    }
            }

            // 結合
            string rtn_sql = tmp_sql_pre + tmp_sql + tmp_sql_post;
            return rtn_sql;

        }


    }
}