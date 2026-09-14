using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace Converter10
{
    static class CommonModule
    {

        #region 共通データ

        public const string CV_FROM_NAME = "賃貸革命V7";
        public const string CV_TO_NAME = "賃貸革命10";
        public const string CV_FROM_SYSTEM = "他社システム";
        public const string CV_KIZON_TITLE = "既存ユーザ用コンバーター";
        public const string CV_HANYO_TITLE = "汎用コンバートキット";
        public const string CV_TASYA_TITLE = "他社システム用コンバーター";

        public enum ConvertTypes                                    // コンバートタイプ
        {
            _汎用 = 0,
        }

        #endregion

        #region メッセージボックス関連

        public const string MSG_END_CV = "コンバート作業を終了します。よろしいですか？";
        public const string MSG_END_CV_TORIREKI = "コンバート実績を保持しますか？" + Constants.vbCrLf + "(再度データコンバートを実行する際、既に実行済みの項目(チェックボックス)が着色されます)";
        // 20160707 コンバート実績保持の処理追加 -add
        public const string MSG_END_MIDCHK = "中間ファイルチェックが完了しました。" + Constants.vbCrLf + "ログを表示します。";
        public const string MSG_STOP_MIDCHK = "中間ファイルチェックを中断しました。";
        public const string MSG_ERR_CHK_MID = "中間ファイルのフォーマットが変更された可能性があります。" + Constants.vbCrLf + "(初期フォーマット状態から、列やシートが削除されているなど)" + Constants.vbCrLf + "中間ファイルチェック結果を参考に、中間ファイルの修正を" + Constants.vbCrLf + "行って下さい。" + Constants.vbCrLf + "※結果ファイルを表示します。" + Constants.vbCrLf + "※対応方法が不明な場合、サポートへお問い合わせ下さい。";




        public const string MSG_STOP_A = "データコンバートを中止してもよろしいですか？";
        public const string MSG_STOP_B = "データコンバートを中止しました。";
        public const string MSG_CANCEL_A = "コンバート処理をキャンセルしてもよろしいですか？";
        public const string MSG_CANCEL_B = "中間ファイルチェック処理をキャンセルしてもよろしいですか？";
        public const string MSG_STA_A = "コンバートを実行します。よろしいですか？";
        public const string MSG_STA_B = "中間ファイルの作成が完了しました。コンバート作業を続行しますか？";
        public const string MSG_STA_C = "中間ファイルのチェックを行います。よろしいですか？";
        public const string MSG_STA_D = "※必ず、中間ファイルのチェックを行ってから実行して下さい。";                                                                                  // 20161014 改善対応：コンバート処理前に中間ファイルチェックを促す -add
        public const string MSG_CNN_SUCCESS_A = "接続が正常に行えることを確認しました。";
        public const string MSG_CNN_FAILURE_A = "DBへの接続に失敗しました。接続情報を確認して下さい。";
        public const string MSG_ERR_DIR_REL = "指定された紐付ファイル格納先フォルダが見つかりません。再度指定して下さい。";
        public const string MSG_ERR_DIR_LOG = "指定されたログファイル格納先フォルダが見つかりません。再度指定して下さい。";
        public const string MSG_ERR_DIR_MID = "指定された中間ファイル格納先フォルダが見つかりません。再度指定して下さい。";
        public const string MSG_ERR_DIR_MIDLOG = "指定された中間ファイルチェックログファイル格納先フォルダが見つかりません。再度指定して下さい。";
        public const string MSG_ERR_DIR_LIST = "指定されたリストの格納先フォルダが見つかりません。再度指定して下さい。";
        public const string MSG_ERR_CHK_CVITEM = "移行対象項目が選択されていません。1つ以上選択して下さい。";
        public const string MSG_ERR_MAKE_CVDB = "作業用DB作成中にエラーが生じました。処理を終了します。";
        public const string MSG_CATION_JIZENCHK = "未チェック項目が存在します。全項目にチェックが入っていない場合、データコンバートを行うことはできません。" + Constants.vbCrLf + "再度、項目チェックを行う場合は「はい」、行わない場合は「いいえ」を選択して下さい。";
        public const string MSG_ERR_UNYOYMD = "運用開始年月が設定されていません。指定して下さい。";                                                                                    // 20160525 運用開始年月の追加 -add
                                                                                                                                                              // 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -add
        public const string MSG_ERR_DUMMYBANKCV = "ダミー口座作成中にエラーが生じました。処理を終了します。";
        public const string MSG_ERR_RELCV = "紐付項目移行中にエラーが生じました。処理を終了します。";                                                                                  // 20160525 紐付項目の挿入処理の追加
        public const string MSG_ERR_FILEOPEN = "以下のファイルが開かれています。保存して終了して下さい。";                                                                             // 20160825 ファイルが開かれているかチェックする機能を追加 -add
        public const string MSG_LIST_OUTSTA = "リストを出力します。よろしいですか？";                                                                                                  // 20160614 事前作業実行処理の追加 -add
        public const string MSG_LIST_OUTEND = "ファイルの出力が正常に完了しました。リストを表示します。";                                                                              // 20160614 事前作業実行処理の追加 -add
        public const string MSG_LIST_OUTEND_ERR = "ファイル出力中にエラーが発生しました。接続情報等を確認して下さい。";                                                                // 20160614 事前作業実行処理の追加 -add
        public const string MSG_DEL_CV_JISSEKI = "チェックボックスの色を元に戻します。よろしいですか？" + Constants.vbCrLf + "(一度クリアすると再度元に戻すことはできません。)";                 // 20160711 コンバート履歴処理の追加 -add
        public const string MSG_DEL_CV_JISSEKIEND = "チェックボックスの色を元に戻しました。";                                                                                          // 20160711 コンバート履歴処理の追加 -add
        public const string MSG_RELFILE_RENAMESTA = "入力された置換前パスを置換後パスに置換します。一度置換されたデータは元に戻すことはできません。よろしいですか？";                  // 20160707 関連ファイルリネーム処理の追加 -add
        public const string MSG_RELFILE_RENAMEEND = "関連ファイルパスの置換が完了しました。";                                                                                          // 20160707 関連ファイルリネーム処理の追加 -add
        public const string MSG_RELFILE_NOTEXISTERR = "紐付設定情報が存在しないためリストを出力することができません。";                                                                // 20160720 紐付設定ファイル出力機能の追加
        public const string MSG_RELLIST_OUTSTA = "紐付設定内容のリストを出力します。よろしいですか？";                                                                                 // 20160720 紐付設定ファイル出力機能の追加
        public const string MSG_ERR_VERBLANK = "賃貸革命V7が既に削除されている可能性があります。" + Constants.vbCrLf + "このままコンバートを進めてもよろしいですか？";                           // 20160829 対象革命のバージョン判定処理を追加 -add
        public const string MSG_ERR_VERBLANKEND = "コンバートを終了します。" + Constants.vbCrLf + "※不明点がある場合、サポートへご確認下さい。";                                                // 20160829 対象革命のバージョン判定処理を追加 -add
        public const string MSG_ERR_VERNOTV7 = "賃貸革命V7からのみコンバートを行うことができます。" + Constants.vbCrLf + "コンバートを行う前に、賃貸革命V7へバージョンアップを" + Constants.vbCrLf + "行って下さい。(コンバーターを終了します)" + Constants.vbCrLf + "※不明点がある場合、サポートへご確認下さい。";

        // 20160829 対象革命のバージョン判定処理を追加 -add
        public const string MSG_ERR_MIDFILENOTEXIST = "指定された中間ファイルが存在しません。" + Constants.vbCrLf + "ファイル格納先を再度指定して下さい。";                                      // 20160829 中間ファイル有無確認処理を追加 -add
        public const string MSG_ERR_BASEMIDFILEREAD = "中間ファイルの読込時にエラーが発生しました。";                                                                                  // 20160913_2 エラー時の処理を追加する修正 -add
        public const string MSG_ERR_EXISTMIDFILEWRITE = "中間ファイルの書込時にエラーが発生しました。";                                                                                // 20160913_2 エラー時の処理を追加する修正 -add
        public const string MSG_ERR_EXISTMIDFILEINIT = "中間ファイルの初期化時にエラーが発生しました。";                                                                               // 20161007 既存中間ファイル初期化処理の追加 -add

        #endregion

        #region ログ関連

        // ログ_ヘッダー
        public const string LOG_HEADER1 = "出力日時";
        public const string LOG_HEADER2 = "ログ分類";
        public const string LOG_HEADER3 = "ログ種別";
        public const string LOG_HEADER4 = "処理項目_大分類";
        public const string LOG_HEADER5 = "処理項目_小分類";
        public const string LOG_HEADER6 = "処理結果内容";
        public const string LOG_HEADER7 = "対象データ";
        public const string LOG_HEADER8 = "不備原因";
        public const string LOG_HEADER9 = "対処方法";
        public const string LOG_HEADER10 = "設定値変更前";
        public const string LOG_HEADER11 = "設定値変更後";
        public const string LOG_HEADER12 = "対象TBL名";
        public const string LOG_HEADER_TOTAL = "出力日時,ログ分類,ログ種別,処理項目_大分類,処理項目_小分類,処理結果内容,対象データ,不備原因,対処方法,設定値変更前,設定値変更後,対象TBL名";

        // ログ_文字列(ログ分類)
        public const string LOG_RUI_RIREKI = "動作履歴";
        public const string LOG_RUI_KEIKOKU = "警告";
        public const string LOG_RUI_TYUUI = "注意";

        // ログ_文字列(ログ種別)
        public const string LOG_SYU_STA = "処理開始";
        public const string LOG_SYU_CON = "移行元DB接続";
        public const string LOG_SYU_READ = "移行元DB読込";
        public const string LOG_SYU_END = "処理終了";
        public const string LOG_SYU_STOP = "処理中断";
        public const string LOG_SYU_MIDMAKE = "中間ファイル作成";
        public const string LOG_SYU_MIDHEADERCHK = "中間ファイルヘッダーチェック";
        public const string LOG_SYU_MIDDATACHK = "中間ファイルデータチェック";
        public const string LOG_SYU_DBWRITE = "DB書込";

        // ログ_文字列(処理項目)
        public const string LOG_SYORIKOMK_STA = "コンバートツール起動";
        public const string LOG_SYORIKOMK_END = "コンバートツール終了";
        public const string LOG_SYORIKOMK_CON_OK = "接続成功";
        public const string LOG_SYORIKOMK_CON_NG = "接続失敗";
        public const string LOG_SYORIKOMK_CV_STA = "コンバート処理開始";
        public const string LOG_SYORIKOMK_CV_END = "コンバート処理終了";
        public const string LOG_SYORIKOMK_CV_STOP = "コンバート処理中断";
        public const string LOG_SYORIKOMK_CV_OK = "正常終了";
        public const string LOG_SYORIKOMK_CV_NG = "異常終了";
        public const string LOG_SYORIKOMK_MAKE_MIDSTA = "中間ファイル作成開始";
        public const string LOG_SYORIKOMK_MAKE_MIDEND = "中間ファイル作成終了";
        public const string LOG_SYORIKOMK_DBCVSTA = "DB移行開始";
        public const string LOG_SYORIKOMK_DBCVEND = "DB移行終了";
        public const string LOG_SYORIKOMK_READSTA = "移行元データ読込開始";
        public const string LOG_SYORIKOMK_READEND = "移行元データ読込終了";
        public const string LOG_SYORIKOMK_MIDWRITESTA = "中間ファイル書込開始";
        public const string LOG_SYORIKOMK_MIDWRITEEND = "中間ファイル書込終了";
        public const string LOG_SYORIKOMK_MIDCHKSTA = "中間ファイルチェック開始";
        public const string LOG_SYORIKOMK_MIDCHKEND = "中間ファイルチェック終了";
        public const string LOG_SYORIKOMK_MIDCHKSTOP = "中間ファイルチェック中断";
        public const string LOG_SYORIKOMK_MIDREADSTA = "中間ファイル読込開始";
        public const string LOG_SYORIKOMK_MIDREADEND = "中間ファイル読込終了";
        public const string LOG_SYORIKOMK_RELSTA = "紐付設定開始";
        public const string LOG_SYORIKOMK_RELEND = "紐付設定終了";
        public const string LOG_SYORIKOMK_RELCANCEL = "紐付設定中断";                                                  // 20160707 本体と紐付ツールの中断を同期させる処理の追加 -add
        public const string LOG_SYORIKOMK_TMPTBLREADSTA = "紐付設定読込開始";
        public const string LOG_SYORIKOMK_TMPTBLREADEND = "紐付設定読込終了";
        public const string LOG_SYORIKOMK_WRITESTA = "DB書込開始";
        public const string LOG_SYORIKOMK_WRITEEND = "DB書込終了";
        public const string LOG_SYORIKOMK_MIDCOPY = "中間ファイルコピー";                                             // 20161009 ログ出力処理追加 -add：汎用から既存U用中間ファイルコピー処理中
        public const string LOG_SYORIKOMK_MIDCHECK = "中間ファイルシート・データチェック";                            // 20161009 ログ出力処理追加 -add：既存U用中間ファイルデータチェック
        public const string LOG_SYORIKOMK_MIDCONDI = "中間ファイルデータ調整";                                        // 20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整
        public const string LOG_SYORIKOMK_MIDFILE = "中間ファイル";                                                   // 20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整
        public const string LOG_SYORIKOMK_MIDONECOPY = "中間ファイルコピー(個別)";                                    // 20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整
        public const string LOG_SYORIKOMK_ERR_MIDONECOPY = "×中間ファイルコピーエラー";                              // 20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整

        // ログ_文字列(処理結果内容)
        public const string LOG_NAIYO_NORMALEND = "正常終了";
        public const string LOG_NAIYO_NOTNORMALEND = "異常終了";

        public const string LOG_NAIYO_ERR_MIDHEADERCHK = "ヘッダー不正";
        public const string LOG_NAIYO_ERR_MIDHEADERRANGECHK = "ヘッダー不正・作成範囲外データ";

        public const string LOG_NAIYO_ERR_OVERLAP = "重複";
        public const string LOG_NAIYO_ERR_MISMATCH_KEY = "不適合(キー)";
        public const string LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED = "必須項目無し";
        public const string LOG_NAIYO_ERR_NOTEXISTDATA_BASE = "親データ無し";

        public const string LOG_NAIYO_ERR_OUTOFRANGE = "有効範囲外";
        public const string LOG_NAIYO_ERR_OUTOFRANGE_STR = "有効範囲外(文字数)";
        public const string LOG_NAIYO_ERR_GAIJI = "移行不可文字列";
        public const string LOG_NAIYO_ERR_NOTEXISTDATA = "データ無し";
        public const string LOG_NAIYO_ERR_MISMATCH = "不適合";

        public const string LOG_NAIYO_ERR_NOTEXISTMSTDATA = "マスタ不一致";
        public const string LOG_NAIYO_ERR_NOTEXISTFILEDATA = "参照ファイル無し";

        public const string LOG_NAIYO_ERR_RELFILEPIC = "関連ファイルに画像ファイル";                                   // 20160525 クレーム関連ファイルの画像判別処理実装 -add
        public const string LOG_NAIYO_ERR_TAIONASI = "対応不要";                                                       // 20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -add
        public const string LOG_NAIYO_ERR_DATACHK = "データチェックエラー";                                            // 20160328 型落ちを考慮した処理へ修正 -add

        public const string LOG_NAIYO_ERR_KYOYOKAGI = "共用鍵";
        public const string LOG_NAIYO_ERR_SENYOKAGI = "専用鍵";

        public const string LOG_NAIYO_GET_POSTCODE = "郵便番号変換";                                                   // 20160704 物件住所から郵便番号を読み込む処理の追加 -add
        public const string LOG_NAIYO_CHG_STR = "文字列変換";                                                          // 20160829 口座名義カナチェック機能の追加 -add

        // ログ_文字列(不備原因)
        public const string LOG_HUBI_MIDHEADERCHK = "ヘッダーの文字列が変更されている可能性があります。";
        public const string LOG_HUBI_MIDHEADERRANGECHK = "ヘッダーの削除、またはヘッダー範囲外にデータが設定されている可能性があります。";
        public const string LOG_HUBI_OVERLAP = "キーが重複しています。最初に処理する1件のみを移行対象とします。";
        public const string LOG_HUBI_MISMATCH_KEY = "キーの値が不正であるため移行できません。";
        public const string LOG_HUBI_NOTEXISTDATA_REQUIRED = "革命でのデータ登録に必須となっている項目が存在しないため移行できません。";
        public const string LOG_HUBI_NOTEXISTDATA_REQUIRED_ADDSTR = "革命でのデータ登録に必須となっている項目が存在しないため任意の文字列を付加して移行します。";       // 20160720 連動情報構築 -add
        public const string LOG_HUBI_NOTEXISTDATA_BASE = "親マスタにデータが存在しないため移行できません。";
        public const string LOG_HUBI_OUTOFRANGE_NODEF = "有効範囲外のデータです。移行できないため空データを設定します。";                                               // 20161014 ログ出力内容修正 -add
        public const string LOG_HUBI_OUTOFRANGE = "有効範囲外のデータです。移行できないためデフォルト値を設定します。";
        public const string LOG_HUBI_OUTOFRANGE_STR = "有効範囲外のデータです。許容範囲の超過分を切り捨てて登録します。";
        public const string LOG_HUBI_OUTOFRANGE_STR_SHORTAGE = "有効範囲外のデータです。許容範囲の不足分に任意の文字列を付加して移行します。";                          // 20160720 連動情報構築 -add
        public const string LOG_HUBI_GAIJI = "移行不可文字列です。「■」に変換して移行します。";
        public const string LOG_HUBI_NOTEXISTDATA = "移行元データが存在しません。デフォルト値を設定します。";
        public const string LOG_HUBI_MISMATCH_NODEF = "移行元データが不正であるため空データを設定します。";                                                             // 20161014 ログ出力内容修正 -add
        public const string LOG_HUBI_MISMATCH = "移行元データが不正です。移行できないためデフォルト値を設定します。";
        public const string LOG_HUBI_NOTEXISTDATA_MSTEXIST = "マスタとデータが一致しないため移行できません。";
        public const string LOG_HUBI_NOTEXISTDATA_RELEXIST = "マスタと紐付けデータが一致しないため移行できません。";
        public const string LOG_HUBI_NOTEXISTDATA_FILEEXIST = "参照ファイルが存在しないため移行できません。";
        public const string LOG_HUBI_MAKEMIDFILE = "中間ファイル出力時にエラーが発生しました。";
        public const string LOG_HUBI_MAKELOGFILE = "ログファイル出力時にエラーが発生しました。";
        public const string LOG_HUBI_MAKEMIDCHKLOGFILE = "中間ファイルチェックログ出力時にエラーが発生しました。";
        public const string LOG_HUBI_MISMATCH_NOTDEFAULT = "革命10では設定できない値であるため移行できません。";
        public const string LOG_HUBI_HYHENDO = "部屋情報から変動費を取得できません。";                                                                                  // 20160603 ユーザーデータ検証による修正 -add

        public const string LOG_HUBI_ERR_DATACHK_A = "データチェック用の最大最小値を取得時にエラーが発生しました。";
        public const string LOG_HUBI_ERR_DATACHK_B = "データチェック用のフィールドサイズを取得時にエラーが発生しました。";
        public const string LOG_HUBI_ERR_DATACHK_C = "データチェック用のキーナンバーを取得時にエラーが発生しました。";

        public const string LOG_HUBI_RELFILEPIC = "関連ファイルに画像ファイルが含まれています。";                                                                       // 20160525 クレーム関連ファイルの画像判別処理実装 -add
        public const string LOG_HUBI_KYOYOKAGI = "共用鍵のデータです。";
        public const string LOG_HUBI_SENYOKAGI = "専用鍵のデータです。";
        public const string LOG_HUBI_GET_POSTCODE = "住所から郵便番号を取得しました。";                                                                                 // 20160704 物件住所から郵便番号を読み込む処理の追加 -add
        public const string LOG_HUBI_TAIONASI = "不要なデータです。";                                                                                                   // 20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -add
        public const string LOG_HUBI_CHG_KANASTR = "口座名義用の文字列へ変更しました。";                                                                                // 20160829 口座名義カナチェック機能の追加 -add
        public const string LOG_HUBI_ERR_KANASTR = "半角カナへ変更できないため移行できません。";                                                                        // 20160829 口座名義カナチェック機能の追加 -add
        public const string LOG_HUBI_CHG_URLSTR = "URL、メールアドレス用の文字列へ変更しました。";                                                                      // 20160829 メールアドレス、URLの正規化処理を追加 -add

        // ログ_文字列(対処方法)
        public const string LOG_TAISYO_MIDHEADERCHK = "ヘッダーを元に戻す、または中間ファイルのバックアップを使用して再度作成して下さい。";
        public const string LOG_TAISYO_MIDHEADERRANGECHK = "ヘッダーを元に戻す、ヘッダー範囲外データの除去、または中間ファイルのバックアップを使用して再度作成して下さい。";

        public const string LOG_TAISYO_OVERLAP = "重複しないように再度キーを設定して下さい。";
        public const string LOG_TAISYO_MISMATCH_KEY = "データタイプ(数値、文字列等)、または最大最小値を確認して下さい。";
        public const string LOG_TAISYO_NOTEXISTDATA_REQUIRED = "必須項目に値を設定する、またはコンバート後に革命上で直接登録して下さい。";
        public const string LOG_TAISYO_NOTEXISTDATA_BASE = "親マスタとの紐付けを確認して下さい。";

        public const string LOG_TAISYO_OUTOOFRANGE_STR = "許容範囲内となるように値を設定する、またはコンバート後に革命上で直接編集して下さい。";
        public const string LOG_TAISYO_GAIJI = "移行可能な文字列へ変更する、またはコンバート後に革命上で直接編集して下さい。";
        public const string LOG_TAISYO_DEFAULT = "値を再度設定する、またはコンバート後に革命上で直接編集して下さい。";

        public const string LOG_TAISYO_NOTEXISTDATA_TODOFUKEN = "マスタと一致するように値を設定して下さい。";
        public const string LOG_TAISYO_DATACHK_A = "作業ファイルに不正があります。確認して下さい。(開発用)";
        public const string LOG_TAISYO_DATACHK_B = "フィールドサイズ取得方法に不正があります。確認して下さい。(開発用)";

        public const string LOG_TAISYO_RELFILEPIC = "画像コンバーターでコンバートを行って下さい。";                                                                     // 20160525 クレーム関連ファイルの画像判別処理実装 -add
        public const string LOG_TAISYO_KYOYOKAGI = "共用鍵として移行します。";
        public const string LOG_TAISYO_SENYOKAGI = "専用鍵として移行します。";

        public const string LOG_TAISYO_GET_POSTCODE = "設定された郵便番号を確認して下さい。";                                                                           // 20160704 物件住所から郵便番号を読み込む処理の追加 -add
        public const string LOG_TAISYO_TAIONASI = "対応はありません。";                                                                                                 // 20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -add

        // ログ_ログ挿入用仮テーブル名
        public const string LOG_TMP_TABLENAME = "cv_log";
        public const string LOG_TMP_MIDCHKTABLENAME = "midchk_log";

        // ログ_その他
        public const string LOG_GET_KEYERRVALUE = "キーエラー";

        #endregion

        #region 状況出力関連

        public const string SITUATION_MID_RUN = "中間ファイル書込中...";                       // 20160825 進捗状況の表示修正 -add
        public const string SITUATION_MID_READ = "中間ファイル読込中...";                      // 20161007 実行ボタン押下→DB書込開始までの進捗表示処理追加 -add
        public const string SITUATION_MID_CHKBEFORE = "中間ファイルチェック準備中...";         // 20161011 コンバート実行時の進捗表示対応 -add
        public const string SITUATION_MID_CHK = "中間ファイルチェック中...";                   // 20161007 実行ボタン押下→DB書込開始までの進捗表示処理追加 -add
        public const string SITUATION_MID_CHKAFTER = "中間ファイルチェック終了処理中...";      // 20161011 コンバート実行時の進捗表示対応 -add
        public const string SITUATION_CV_RUN = "コンバート処理中...";
        public const string SITUATION_REL_RUN = "紐付設定中...";                               // 20160825 進捗状況の表示修正 -add
        public const string SITUATION_CV_BFRUN = "コンバート処理準備中 ...";                               // 20161017 進捗表示ラベル初期表示修正 -add
        public const string SITUATION_CV_STA = "コンバート処理開始";
        public const string SITUATION_CV_END = "コンバート処理終了";
        public const string SITUATION_CV_ENDMSG = "コンバートが正常に完了しました。" + Constants.vbCrLf + "出力されたログを確認して下さい。";
        public const string SITUATION_CV_ERRENDMSG_A = "コンバートを正常に行うことができませんでした。" + Constants.vbCrLf + "出力されたログを確認して下さい。";
        public const string SITUATION_STOP = "コンバート処理中断";
        public const string SITUATION_STOP_END = "コンバート処理を中断しました。";
        public const string SITUATION_EXTRACTION = "件数抽出中...";                            // 20161014 改善対応：移行対象項目件数抽出処理改善 -add
        public const string SITUATION_DATACONDITION = "中間ファイル調整中...";                 // 20161014 改善対応：移行対象項目件数抽出処理改善 -add

        public const string SITUATION_RELSTA = "紐付設定開始";
        public const string SITUATION_RELEND = "紐付設定終了";
        public const string SITUATION_READSTA = "移行元データ読込開始";
        public const string SITUATION_READEND = "移行元データ読込終了";
        public const string SITUATION_MIDSTA = "中間ファイル作成開始";
        public const string SITUATION_MIDEND = "中間ファイル作成終了";
        public const string SITUATION_MIDERREND = "中間ファイルの作成中にエラーが発生しました。" + Constants.vbCrLf + "接続情報に問題がある可能性があります。接続情報を確認し、再度コンバートを実行して下さい。";
        public const string SITUATION_MIDFAIL = "中間ファイル作成失敗";
        public const string SITUATION_MIDREADSTA = "中間ファイル読込開始";
        public const string SITUATION_MIDREADEND = "中間ファイル読込終了";
        public const string SITUATION_WRITESTA = "DB書込開始";
        public const string SITUATION_WRITEEND = "DB書込終了";
        public const string SITUATION_CVITEMSUCCESS = "○";
        public const string SITUATION_CVITEMFAILURE = "×";
        public const string SITUATION_CVITEMTOTALCNT = "全件数      ";
        public const string SITUATION_CVITEMCVCNT = "移行件数   ";
        public const string SITUATION_CVITEMCONDCNT = "調整件数   ";
        public const string SITUATION_CVITEMNOTCVCNT = "未移行件数";

        #endregion

        #region 中間ファイル関連

        // 中間ファイルの最初のセル
        public const string MIDFILE_READWRITE_CELLSTA = "A1";

        // 中間ファイル読込/書込開始　※ヘッダの内容によって読み込み位置を変更
        public const int EXISTMIDFILE_READWRITE_ROW = 2;   // 行                     
        public const string EXISTMIDFILE_READWRITE_COL = "A";  // 列

        // 汎用用中間ファイル読込/書込開始
        public const int BASEMIDFILE_HEADER_ROW_SETUBIGRP = 1;                    // 20160913_2 部屋設備移行処理の追加 -add
        public const int BASEMIDFILE_HEADER_ROW = 2;                              // 20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -add
        public const int BASEMIDFILE_READWRITE_ROW = 12;   // 行                    '20160928 ヘッダー修正に伴う開始行修正 「補足」行の削除に伴い 13→12 へ変更
        public const string BASEMIDFILE_READWRITE_COL = "B";   // 列

        // 紐付情報格納ファイル名
        public const string MIDFILE_RELNAME = "各マスタ紐付け情報";

        #endregion

        #region DBから取得した情報を格納する共通オブジェクト

        // 革命10テーブル/フィールド情報取得用
        public static SafeDictionary<string, string> Hash_TblName_JpToAlpha;                                      // テーブル名(日本語・アルファベット)
        public static SafeDictionary<string, string> Hash_TblName_AlphaToJp;                                      // テーブル名(アルファベット・日本語)
        public static SafeDictionary<string, string> Hash_FldName_JpToAlpha;                                      // フィールド名(日本語・アルファベット)
        public static SafeDictionary<string, string> Hash_FldName_AlphaToJp;                                      // フィールド名(アルファベット・日本語)
        public static SafeDictionary<string, string> Hash_FiledTypeAlpha;                                         // フィールド型(革命フィールド名)
        public static SafeDictionary<string, string> Hash_FiledSizeAlpha;                                         // フィールドサイズ(革命フィールド名)
        public static SafeDictionary<string, string> Hash_FiledTypeJp;                                            // フィールド型(日本語フィールド名)
        public static SafeDictionary<string, string> Hash_FiledSizeJp;                                            // フィールドサイズ(日本語フィールド名)
        public static SafeDictionary<string, string> Hash_DefaultValue;                                           // デフォルト値
        public static SafeDictionary<string, string> Hash_FiledKey;                                               // キーに該当するフィールド
        public static SafeDictionary<string, string> Hash_Min_Code;                                               // 画面上で設定された最小値格納
        public static SafeDictionary<string, string> Hash_Max_Code;                                               // 画面上で設定された最大値格納
        public static SafeDictionary<string, string> Hash_Mst_ReferenceAlpha;                                     // 都道府県等、存在有無を確認する項目を格納
        public static SafeDictionary<string, string> Hash_FldName_JpToAlpha_Mid;
        public static List<string> List_Required_Field = new List<string>();                               // キーではないが革s命上で登録する際に必要な項目のフィールドを格納
        public static List<string> List_FldNameJp = new List<string>();                                    // 日本語フィールド名を格納

        public static List<string> List_Existmidheader = new List<string>();
        public static SafeDictionary<string, string> Hash_ExistMiddatatype = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_ExistMiddatamin = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_ExistMiddatamax = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_ExistMiddatadef = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_ExistMiddatakey = new SafeDictionary<string, string>();
        public static List<string> List_ExistMiddatareq = new List<string>();
        public static SafeDictionary<string, string> Hash_ExistMiddataref = new SafeDictionary<string, string>();
        public static List<string> List_ExistMiddatacv = new List<string>();                               // 20160926 選定した移行項目をプログラムへ反映する修正(開発用) -add
        public static List<string> List_Basemidheader = new List<string>();
        public static SafeDictionary<string, string> Hash_BaseMiddatatype = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_BaseMiddatamin = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_BaseMiddatamax = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_BaseMiddatadef = new SafeDictionary<string, string>();
        public static SafeDictionary<string, string> Hash_BaseMiddatakey = new SafeDictionary<string, string>();
        public static List<string> List_BaseMiddatareq = new List<string>();
        public static SafeDictionary<string, string> Hash_BaseMiddataref = new SafeDictionary<string, string>();
        public static List<string> List_BaseMiddatacv = new List<string>();                                // 20160926 選定した移行項目をプログラムへ反映する修正 -add
        public static SafeDictionary<string, string> Hash_BaseMidToExistMid_FS = new SafeDictionary<string, string>();                               // 20160905 中間ファイルコピー処理改善 -add
        public static SafeDictionary<string, string> Hash_ExistMidToBaseMid_FS = new SafeDictionary<string, string>();                               // 20160905 中間ファイルコピー処理改善 -add
        public static SafeDictionary<string, string> Hash_BaseMidToExistMid_SH = new SafeDictionary<string, string>();                               // 20160905 中間ファイルコピー処理改善 -add
        public static SafeDictionary<string, string> Hash_ExistMidToBaseMid_SH = new SafeDictionary<string, string>();                               // 20160905 中間ファイルコピー処理改善 -add

        public static SafeDictionary<string, string> Hash_MidKey_FieldAlpha;                                      // 中間ファイルキー格納用
        public static SafeDictionary<string, string> Hash_MidKey_FieldJp;                                         // 中間ファイルキー格納用
        public static SafeDictionary<string, string> Hash_Kinyu;                                                  // 金融機関
        public static SafeDictionary<string, string> Hash_KinyuTen;                                               // 金融機関支店
        public static SafeDictionary<string, string> Hash_KagiTitleKyoyo;                                         // 共用鍵格納用
        public static SafeDictionary<string, string> Hash_KagiTitleSenyo;                                         // 専用鍵格納用
        public static SafeDictionary<string, string> Hash_KagiTitle_V7to10;                                       // 鍵タイトル紐付要
        public static SafeDictionary<string, string> Hash_SetubiMid;                                              // 中間ファイルデータと設備Noを紐付 (V7設備データ-10設備No)
        public static SafeDictionary<string, string> Hash_SetubiMst;                                              // 10設備マスタ (設備No-項目guid)
        public static SafeDictionary<string, string> Hash_SetubiMst_UserMake;                                     // 20160829 設備の新規挿入処理を追加 -add
        public static SafeDictionary<string, string> Hash_ClaimBruiKasyo;                                         // クレーム分類(箇所)
        public static SafeDictionary<string, string> Hash_ClaimBruiClaim;                                         // クレーム分類(クレーム)

        #endregion

        #region 紐付情報を格納する共通オブジェクト

        public static SafeDictionary<string, string> Hash_Rel_Bkrui;                                              // 物件分類紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Kozo;                                               // 構造マスタ紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Hyrui;                                              // 部屋分類紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Toritaiyo;                                          // 取引態様紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Kozasyubetu;                                        // 口座種別紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Nkinkomk;                                           // 入金項目紐付情報
        public static SafeDictionary<string, string> Hash_Rel_NkinkomkZksei;                                      // 入金項目属性紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Nkinkbn;                                            // 入金区分紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Hendometer;                                         // 各戸メーター分類紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Setubi;                                             // 設備紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Kyrui;                                              // 契約分類紐付情報
        public static SafeDictionary<string, string> Hash_Rel_Kyrui_Teikisyakuyakbn;                              // 契約分類紐付情報(定期借家区分用)
        public static SafeDictionary<string, string> Hash_Rel_Tosiyotokbn;                                        // 都市計画・用途地域区分情報
        public static SafeDictionary<string, string> Hash_Rel_Tosiyotono;                                         // 都市計画・用途地域No情報
        public static SafeDictionary<string, string> Hash_Rel_JisyaKoza;                                          // 自社口座情報取得用
        public static SafeDictionary<string, string> Hash_Rel_FBInfo_KozaFurikaeFmt;                              // FBフォーマット(口座振替)
        public static SafeDictionary<string, string> Hash_Rel_FBInfo_FuriIraiFmt;                                 // FBフォーマット(振込依頼人)
        public static SafeDictionary<string, string> Hash_Rel_FBInfo_NsSettingFmt;                                // FBフォーマット(入手金設定)
        public static SafeDictionary<string, string> Hash_Rel_FBInfo_NsSettingJisya;                              // FBフォーマット(入手金設定_自社)

        #endregion

        #region その他共通文字列群

        public const string INDENT_0 = "  ";                                          // 20160825 ファイルが開かれているかチェックする機能を追加 -add
        public const string INDENT_1 = "    ";
        public const string INDENT_2 = "        ";

        // VIEW作成クエリ共通構文
        public const string PRE_VIEW_NAME = "CVVW_";
        public const string PRE_VIEW_QRY = " CREATE VIEW ";
        public const string POST_VIEW_QRY = " AS ";

        // 仮テーブル作成クエリ共通構文
        public const string PRE_TBL_NAME = "CVTBL_";

        // フォルダ、ファイル名
        public const string DIR_EXEDIR_NAME = "Exe";                                  // 本体/紐付Exe格納先フォルダ名
        public const string DIR_CONVEXEDIR_NAME = "ConvMain";                         // 本体Exe格納先フォルダ名
        public const string DIR_RELEXEDIR_NAME = "RelationSetting";                    // 紐付Exe格納先フォルダ名
        public const string DIR_RELEXE_NAME = "RelationSetting.exe";                   // 紐付Exe名
        public const string DIR_FILEDIR_NAME = "File";                                // 中間ファイル等、実行時に必要なファイルの格納先フォルダ名
        public const string DIR_MAINLOG_NAME = "convlog";                             // ログファイル格納先フォルダ名
        public const string DIR_MIDLOG_NAME = "midchklog";                            // 中間ファイルログ格納先フォルダ名
        public const string DIR_MID_NAME = "middlefile";                              // 中間ファイル格納先フォルダ名
        public const string DIR_REL_NAME = "relationfile";                            // 紐付ファイル格納先フォルダ名
        public const string DIR_RELLOG_NAME = "rellog";                               // 紐付ツールログ格納先フォルダ名           '20160926 フォルダの自動生成処理を追加 -add
        public const string DIR_TEMP_CSV = "csv";                                     // CSVファイル格納先フォルダ名
        public const string DIR_INI_NAME = "ini";                                     // 次回起動時用文字列保存ファイル格納先 (接続情報等)
        public const string DIR_LIST_NAME = "list";                                   // 事前作業で出力したリストの格納先フォルダ名
        public const string FILE_V7_CONNAME = "coninfo_V7.xml";                       // V7接続情報保管ファイル
        public const string FILE_10_CONNAME = "coninfo_10.xml";                       // 10接続情報保管ファイル
        public const string FILE_CVJISSEKINAME = "cvjisseki.xml";                     // コンバート実績保持用ファイル             '20160707 コンバート実績保持の処理追加 -add
        public const string DIR_GAZOEXEDIR_NAME = "GazoConverter";                    // 画像コンバートツール格納先フォルダ名     '20160711 画像CV呼出処理の追加 -add
        public const string DIR_GAZOEXE_NAME = "GazoConverter.exe";                      // 画像コンバートツール名                   '20160711 画像CV呼出処理の追加 -add
        public const string MAIN_VIEW_NAME_BKGUID = "tmp_bk_guid";
        public const string FILE_HMIDD_NAME = "中間ファイル.xlsx";                    // 汎用中間ファイル   

        public const string COMPANY_NAME = "n-create.co.jp";
        public const string PRODUCT_NAME = "FK8Host";
        public const string DEF_REC_USER = "CONVUSER";                                // 作業者名のデフォルト値
        public const string NJC_DEV = "_njc_dev";                                     // ツール起動時の開発用フラグ

        // 革命10の改行コード
        public const string LINE_BREAK = "$0D$0A";

        // 日付の最大最小デフォルト値
        public const string DEF_MIN_YMD = "1900/01/01";
        public const string DEF_MAX_YMD = "2100/12/31";

        // 区切文字文字列
        public const string STR_SPLIT_1 = "@#@";                                      // 20160913_2 「@#@」を共通変数に変更 -add
        public const string STR_SPLIT_2 = "@$@";                                      // 20160927 ログの内容が不正になっているため修正 -add

        #endregion

        #region その他共通オブジェクト

        public static List<string> List_ImportableImageFileAttributes = new List<string>();                // 20160525 クレーム関連ファイルの画像判別処理実装 -add

        #endregion

        #region 変数

        public static string DefHistory;
        public static bool CancelFlg = false;
        public static bool MidChkCancelFlg = false;                                       // 20160222 中間ファイルチェック時の中断処理の追加 -add
        public static bool Dev_CVFlg = false;                                             // 開発用Converterフラグ
        public static int Log_OutputCnt;                                                 // ログ出力件数
        public static string LogFilePath;
        public static string RelationDirPath;
        public static string MiddleDirPath;
        public static string MiddleLogFilePath;
        public static string BaseMidDirPath;                                                 // 汎用用中間ファイル格納先
        public static string CV_FROM_MIDDLE;                                                 // 汎用用の中間ファイル名
        public static int LimitTimeOut;
        public static DialogResult MsgResult;
        public static string RecUser;
        public static bool InitDBFlg;                                                     // DB初期化フラグ
        public static bool OverWriteDBFlg;                                                // 上書きフラグ(開発用)
        public static int taihitab = 0;                                                  // 前回表示タブINDEX格納(開発タブONOFF用)
        public static string UnyoYMD;                                                        // 20160525 運用開始年月の追加 -add
                                                                                             // ----- 要対応 ----- sta
                                                                                             // Public CNVNO As Integer = ConvertTypes._既存ユーザ用
        public static int CNVNO = (int)ConvertTypes._汎用;                                    // 20160921 初期起動変更 
                                                                                            // ----- 要対応 ----- end

        #endregion

    }
}