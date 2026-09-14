Module CommonModule

#Region "共通データ"

    Public Const CV_FROM_NAME As String = "賃貸革命V7"
    Public Const CV_TO_NAME As String = "賃貸革命10"
    Public Const CV_FROM_SYSTEM As String = "他社システム"
    Public Const CV_KIZON_TITLE As String = "既存ユーザ用コンバーター"
    Public Const CV_HANYO_TITLE As String = "汎用コンバートキット"
    Public Const CV_TASYA_TITLE As String = "他社システム用コンバーター"

    Public Enum ConvertTypes                                    'コンバートタイプ
        _汎用 = 0
        _既存ユーザ用 = 1
        _他社システム用
        _不明 = 99
    End Enum
    
#End Region

#Region "メッセージボックス関連"

    Public Const MSG_END_CV As String = "コンバート作業を終了します。よろしいですか？"
    Public Const MSG_END_CV_TORIREKI As String = "コンバート実績を保持しますか？" & vbCrLf & _
                                                 "(再度データコンバートを実行する際、既に実行済みの項目(チェックボックス)が着色されます)"                                            '20160707 コンバート実績保持の処理追加 -add
    Public Const MSG_END_MIDCHK As String = "中間ファイルチェックが完了しました。" & vbCrLf & "ログを表示します。"
    Public Const MSG_STOP_MIDCHK As String = "中間ファイルチェックを中断しました。"
    Public Const MSG_ERR_CHK_MID As String = "中間ファイルのフォーマットが変更された可能性があります。" & vbCrLf & _
                                             "(初期フォーマット状態から、列やシートが削除されているなど)" & vbCrLf & _
                                             "中間ファイルチェック結果を参考に、中間ファイルの修正を" & vbCrLf & _
                                             "行って下さい。" & vbCrLf & _
                                             "※結果ファイルを表示します。" & vbCrLf & _
                                             "※対応方法が不明な場合、サポートへお問い合わせ下さい。"
    Public Const MSG_STOP_A As String = "データコンバートを中止してもよろしいですか？"
    Public Const MSG_STOP_B As String = "データコンバートを中止しました。"
    Public Const MSG_CANCEL_A As String = "コンバート処理をキャンセルしてもよろしいですか？"
    Public Const MSG_CANCEL_B As String = "中間ファイルチェック処理をキャンセルしてもよろしいですか？"
    Public Const MSG_STA_A As String = "コンバートを実行します。よろしいですか？"
    Public Const MSG_STA_B As String = "中間ファイルの作成が完了しました。コンバート作業を続行しますか？"
    Public Const MSG_STA_C As String = "中間ファイルのチェックを行います。よろしいですか？"
    Public Const MSG_STA_D As String = "※必ず、中間ファイルのチェックを行ってから実行して下さい。"                                                                                  '20161014 改善対応：コンバート処理前に中間ファイルチェックを促す -add
    Public Const MSG_CNN_SUCCESS_A As String = "接続が正常に行えることを確認しました。"
    Public Const MSG_CNN_FAILURE_A As String = "DBへの接続に失敗しました。接続情報を確認して下さい。"
    Public Const MSG_ERR_DIR_REL As String = "指定された紐付ファイル格納先フォルダが見つかりません。再度指定して下さい。"
    Public Const MSG_ERR_DIR_LOG As String = "指定されたログファイル格納先フォルダが見つかりません。再度指定して下さい。"
    Public Const MSG_ERR_DIR_MID As String = "指定された中間ファイル格納先フォルダが見つかりません。再度指定して下さい。"
    Public Const MSG_ERR_DIR_MIDLOG As String = "指定された中間ファイルチェックログファイル格納先フォルダが見つかりません。再度指定して下さい。"
    Public Const MSG_ERR_DIR_LIST As String = "指定されたリストの格納先フォルダが見つかりません。再度指定して下さい。"
    Public Const MSG_ERR_CHK_CVITEM As String = "移行対象項目が選択されていません。1つ以上選択して下さい。"
    Public Const MSG_ERR_MAKE_CVDB As String = "作業用DB作成中にエラーが生じました。処理を終了します。"
    Public Const MSG_CATION_JIZENCHK As String = "未チェック項目が存在します。全項目にチェックが入っていない場合、データコンバートを行うことはできません。" & vbCrLf & _
                                                 "再度、項目チェックを行う場合は「はい」、行わない場合は「いいえ」を選択して下さい。"
    Public Const MSG_ERR_UNYOYMD As String = "運用開始年月が設定されていません。指定して下さい。"                                                                                    '20160525 運用開始年月の追加 -add
    '20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -add
    Public Const MSG_ERR_DUMMYBANKCV As String = "ダミー口座作成中にエラーが生じました。処理を終了します。"
    Public Const MSG_ERR_RELCV As String = "紐付項目移行中にエラーが生じました。処理を終了します。"                                                                                  '20160525 紐付項目の挿入処理の追加
    Public Const MSG_ERR_FILEOPEN As String = "以下のファイルが開かれています。保存して終了して下さい。"                                                                             '20160825 ファイルが開かれているかチェックする機能を追加 -add
    Public Const MSG_LIST_OUTSTA As String = "リストを出力します。よろしいですか？"                                                                                                  '20160614 事前作業実行処理の追加 -add
    Public Const MSG_LIST_OUTEND As String = "ファイルの出力が正常に完了しました。リストを表示します。"                                                                              '20160614 事前作業実行処理の追加 -add
    Public Const MSG_LIST_OUTEND_ERR As String = "ファイル出力中にエラーが発生しました。接続情報等を確認して下さい。"                                                                '20160614 事前作業実行処理の追加 -add
    Public Const MSG_DEL_CV_JISSEKI As String = "チェックボックスの色を元に戻します。よろしいですか？" & vbCrLf & "(一度クリアすると再度元に戻すことはできません。)"                 '20160711 コンバート履歴処理の追加 -add
    Public Const MSG_DEL_CV_JISSEKIEND As String = "チェックボックスの色を元に戻しました。"                                                                                          '20160711 コンバート履歴処理の追加 -add
    Public Const MSG_RELFILE_RENAMESTA As String = "入力された置換前パスを置換後パスに置換します。一度置換されたデータは元に戻すことはできません。よろしいですか？"                  '20160707 関連ファイルリネーム処理の追加 -add
    Public Const MSG_RELFILE_RENAMEEND As String = "関連ファイルパスの置換が完了しました。"                                                                                          '20160707 関連ファイルリネーム処理の追加 -add
    Public Const MSG_RELFILE_NOTEXISTERR As String = "紐付設定情報が存在しないためリストを出力することができません。"                                                                '20160720 紐付設定ファイル出力機能の追加
    Public Const MSG_RELLIST_OUTSTA As String = "紐付設定内容のリストを出力します。よろしいですか？"                                                                                 '20160720 紐付設定ファイル出力機能の追加
    Public Const MSG_ERR_VERBLANK As String = "賃貸革命V7が既に削除されている可能性があります。" & vbCrLf & "このままコンバートを進めてもよろしいですか？"                           '20160829 対象革命のバージョン判定処理を追加 -add
    Public Const MSG_ERR_VERBLANKEND As String = "コンバートを終了します。" & vbCrLf & "※不明点がある場合、サポートへご確認下さい。"                                                '20160829 対象革命のバージョン判定処理を追加 -add
    Public Const MSG_ERR_VERNOTV7 As String = "賃貸革命V7からのみコンバートを行うことができます。" & vbCrLf & _
                                              "コンバートを行う前に、賃貸革命V7へバージョンアップを" & vbCrLf & _
                                              "行って下さい。(コンバーターを終了します)" & vbCrLf & "※不明点がある場合、サポートへご確認下さい。"                                   '20160829 対象革命のバージョン判定処理を追加 -add
    Public Const MSG_ERR_MIDFILENOTEXIST As String = "指定された中間ファイルが存在しません。" & vbCrLf & "ファイル格納先を再度指定して下さい。"                                      '20160829 中間ファイル有無確認処理を追加 -add
    Public Const MSG_ERR_BASEMIDFILEREAD As String = "中間ファイルの読込時にエラーが発生しました。"                                                                                  '20160913_2 エラー時の処理を追加する修正 -add
    Public Const MSG_ERR_EXISTMIDFILEWRITE As String = "中間ファイルの書込時にエラーが発生しました。"                                                                                '20160913_2 エラー時の処理を追加する修正 -add
    Public Const MSG_ERR_EXISTMIDFILEINIT As String = "中間ファイルの初期化時にエラーが発生しました。"                                                                               '20161007 既存中間ファイル初期化処理の追加 -add

#End Region

#Region "ログ関連"

    'ログ_ヘッダー
    Public Const LOG_HEADER1 As String = "出力日時"
    Public Const LOG_HEADER2 As String = "ログ分類"
    Public Const LOG_HEADER3 As String = "ログ種別"
    Public Const LOG_HEADER4 As String = "処理項目_大分類"
    Public Const LOG_HEADER5 As String = "処理項目_小分類"
    Public Const LOG_HEADER6 As String = "処理結果内容"
    Public Const LOG_HEADER7 As String = "対象データ"
    Public Const LOG_HEADER8 As String = "不備原因"
    Public Const LOG_HEADER9 As String = "対処方法"
    Public Const LOG_HEADER10 As String = "設定値変更前"
    Public Const LOG_HEADER11 As String = "設定値変更後"
    Public Const LOG_HEADER12 As String = "対象TBL名"
    Public Const LOG_HEADER_TOTAL As String = "出力日時,ログ分類,ログ種別,処理項目_大分類,処理項目_小分類,処理結果内容,対象データ,不備原因,対処方法,設定値変更前,設定値変更後,対象TBL名"

    'ログ_文字列(ログ分類)
    Public Const LOG_RUI_RIREKI As String = "動作履歴"
    Public Const LOG_RUI_KEIKOKU As String = "警告"
    Public Const LOG_RUI_TYUUI As String = "注意"

    'ログ_文字列(ログ種別)
    Public Const LOG_SYU_STA As String = "処理開始"
    Public Const LOG_SYU_CON As String = "移行元DB接続"
    Public Const LOG_SYU_READ As String = "移行元DB読込"
    Public Const LOG_SYU_END As String = "処理終了"
    Public Const LOG_SYU_STOP As String = "処理中断"
    Public Const LOG_SYU_MIDMAKE As String = "中間ファイル作成"
    Public Const LOG_SYU_MIDHEADERCHK As String = "中間ファイルヘッダーチェック"
    Public Const LOG_SYU_MIDDATACHK As String = "中間ファイルデータチェック"
    Public Const LOG_SYU_DBWRITE As String = "DB書込"

    'ログ_文字列(処理項目)
    Public Const LOG_SYORIKOMK_STA As String = "コンバートツール起動"
    Public Const LOG_SYORIKOMK_END As String = "コンバートツール終了"
    Public Const LOG_SYORIKOMK_CON_OK As String = "接続成功"
    Public Const LOG_SYORIKOMK_CON_NG As String = "接続失敗"
    Public Const LOG_SYORIKOMK_CV_STA As String = "コンバート処理開始"
    Public Const LOG_SYORIKOMK_CV_END As String = "コンバート処理終了"
    Public Const LOG_SYORIKOMK_CV_STOP As String = "コンバート処理中断"
    Public Const LOG_SYORIKOMK_CV_OK As String = "正常終了"
    Public Const LOG_SYORIKOMK_CV_NG As String = "異常終了"
    Public Const LOG_SYORIKOMK_MAKE_MIDSTA As String = "中間ファイル作成開始"
    Public Const LOG_SYORIKOMK_MAKE_MIDEND As String = "中間ファイル作成終了"
    Public Const LOG_SYORIKOMK_DBCVSTA As String = "DB移行開始"
    Public Const LOG_SYORIKOMK_DBCVEND As String = "DB移行終了"
    Public Const LOG_SYORIKOMK_READSTA As String = "移行元データ読込開始"
    Public Const LOG_SYORIKOMK_READEND As String = "移行元データ読込終了"
    Public Const LOG_SYORIKOMK_MIDWRITESTA As String = "中間ファイル書込開始"
    Public Const LOG_SYORIKOMK_MIDWRITEEND As String = "中間ファイル書込終了"
    Public Const LOG_SYORIKOMK_MIDCHKSTA As String = "中間ファイルチェック開始"
    Public Const LOG_SYORIKOMK_MIDCHKEND As String = "中間ファイルチェック終了"
    Public Const LOG_SYORIKOMK_MIDCHKSTOP As String = "中間ファイルチェック中断"
    Public Const LOG_SYORIKOMK_MIDREADSTA As String = "中間ファイル読込開始"
    Public Const LOG_SYORIKOMK_MIDREADEND As String = "中間ファイル読込終了"
    Public Const LOG_SYORIKOMK_RELSTA As String = "紐付設定開始"
    Public Const LOG_SYORIKOMK_RELEND As String = "紐付設定終了"
    Public Const LOG_SYORIKOMK_RELCANCEL As String = "紐付設定中断"                                                  '20160707 本体と紐付ツールの中断を同期させる処理の追加 -add
    Public Const LOG_SYORIKOMK_TMPTBLREADSTA As String = "紐付設定読込開始"
    Public Const LOG_SYORIKOMK_TMPTBLREADEND As String = "紐付設定読込終了"
    Public Const LOG_SYORIKOMK_WRITESTA As String = "DB書込開始"
    Public Const LOG_SYORIKOMK_WRITEEND As String = "DB書込終了"
    Public Const LOG_SYORIKOMK_MIDCOPY As String = "中間ファイルコピー"                                             '20161009 ログ出力処理追加 -add：汎用から既存U用中間ファイルコピー処理中
    Public Const LOG_SYORIKOMK_MIDCHECK As String = "中間ファイルシート・データチェック"                            '20161009 ログ出力処理追加 -add：既存U用中間ファイルデータチェック
    Public Const LOG_SYORIKOMK_MIDCONDI As String = "中間ファイルデータ調整"                                        '20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整
    Public Const LOG_SYORIKOMK_MIDFILE As String = "中間ファイル"                                                   '20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整
    Public Const LOG_SYORIKOMK_MIDONECOPY As String = "中間ファイルコピー(個別)"                                    '20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整
    Public Const LOG_SYORIKOMK_ERR_MIDONECOPY As String = "×中間ファイルコピーエラー"                              '20161009 ログ出力処理追加 -add：既存U用中間ファイルデータ調整

    'ログ_文字列(処理結果内容)
    Public Const LOG_NAIYO_NORMALEND As String = "正常終了"
    Public Const LOG_NAIYO_NOTNORMALEND As String = "異常終了"

    Public Const LOG_NAIYO_ERR_MIDHEADERCHK As String = "ヘッダー不正"
    Public Const LOG_NAIYO_ERR_MIDHEADERRANGECHK As String = "ヘッダー不正・作成範囲外データ"

    Public Const LOG_NAIYO_ERR_OVERLAP As String = "重複"
    Public Const LOG_NAIYO_ERR_MISMATCH_KEY As String = "不適合(キー)"
    Public Const LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED As String = "必須項目無し"
    Public Const LOG_NAIYO_ERR_NOTEXISTDATA_BASE As String = "親データ無し"

    Public Const LOG_NAIYO_ERR_OUTOFRANGE As String = "有効範囲外"
    Public Const LOG_NAIYO_ERR_OUTOFRANGE_STR As String = "有効範囲外(文字数)"
    Public Const LOG_NAIYO_ERR_GAIJI As String = "移行不可文字列"
    Public Const LOG_NAIYO_ERR_NOTEXISTDATA As String = "データ無し"
    Public Const LOG_NAIYO_ERR_MISMATCH As String = "不適合"

    Public Const LOG_NAIYO_ERR_NOTEXISTMSTDATA As String = "マスタ不一致"
    Public Const LOG_NAIYO_ERR_NOTEXISTFILEDATA As String = "参照ファイル無し"

    Public Const LOG_NAIYO_ERR_RELFILEPIC As String = "関連ファイルに画像ファイル"                                   '20160525 クレーム関連ファイルの画像判別処理実装 -add
    Public Const LOG_NAIYO_ERR_TAIONASI As String = "対応不要"                                                       '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -add
    Public Const LOG_NAIYO_ERR_DATACHK As String = "データチェックエラー"                                            '20160328 型落ちを考慮した処理へ修正 -add

    Public Const LOG_NAIYO_ERR_KYOYOKAGI As String = "共用鍵"
    Public Const LOG_NAIYO_ERR_SENYOKAGI As String = "専用鍵"
   
    Public Const LOG_NAIYO_GET_POSTCODE As String = "郵便番号変換"                                                   '20160704 物件住所から郵便番号を読み込む処理の追加 -add
    Public Const LOG_NAIYO_CHG_STR As String = "文字列変換"                                                          '20160829 口座名義カナチェック機能の追加 -add

    'ログ_文字列(不備原因)
    Public Const LOG_HUBI_MIDHEADERCHK As String = "ヘッダーの文字列が変更されている可能性があります。"
    Public Const LOG_HUBI_MIDHEADERRANGECHK As String = "ヘッダーの削除、またはヘッダー範囲外にデータが設定されている可能性があります。"
    Public Const LOG_HUBI_OVERLAP As String = "キーが重複しています。最初に処理する1件のみを移行対象とします。"
    Public Const LOG_HUBI_MISMATCH_KEY As String = "キーの値が不正であるため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_REQUIRED As String = "革命でのデータ登録に必須となっている項目が存在しないため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_REQUIRED_ADDSTR As String = "革命でのデータ登録に必須となっている項目が存在しないため任意の文字列を付加して移行します。"       '20160720 連動情報構築 -add
    Public Const LOG_HUBI_NOTEXISTDATA_BASE As String = "親マスタにデータが存在しないため移行できません。"
    Public Const LOG_HUBI_OUTOFRANGE_NODEF As String = "有効範囲外のデータです。移行できないため空データを設定します。"                                               '20161014 ログ出力内容修正 -add
    Public Const LOG_HUBI_OUTOFRANGE As String = "有効範囲外のデータです。移行できないためデフォルト値を設定します。"
    Public Const LOG_HUBI_OUTOFRANGE_STR As String = "有効範囲外のデータです。許容範囲の超過分を切り捨てて登録します。"
    Public Const LOG_HUBI_OUTOFRANGE_STR_SHORTAGE As String = "有効範囲外のデータです。許容範囲の不足分に任意の文字列を付加して移行します。"                          '20160720 連動情報構築 -add
    Public Const LOG_HUBI_GAIJI As String = "移行不可文字列です。「■」に変換して移行します。"
    Public Const LOG_HUBI_NOTEXISTDATA As String = "移行元データが存在しません。デフォルト値を設定します。"
    Public Const LOG_HUBI_MISMATCH_NODEF As String = "移行元データが不正であるため空データを設定します。"                                                             '20161014 ログ出力内容修正 -add
    Public Const LOG_HUBI_MISMATCH As String = "移行元データが不正です。移行できないためデフォルト値を設定します。"
    Public Const LOG_HUBI_NOTEXISTDATA_MSTEXIST As String = "マスタとデータが一致しないため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_RELEXIST As String = "マスタと紐付けデータが一致しないため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_FILEEXIST As String = "参照ファイルが存在しないため移行できません。"
    Public Const LOG_HUBI_MAKEMIDFILE As String = "中間ファイル出力時にエラーが発生しました。"
    Public Const LOG_HUBI_MAKELOGFILE As String = "ログファイル出力時にエラーが発生しました。"
    Public Const LOG_HUBI_MAKEMIDCHKLOGFILE As String = "中間ファイルチェックログ出力時にエラーが発生しました。"
    Public Const LOG_HUBI_MISMATCH_NOTDEFAULT As String = "革命10では設定できない値であるため移行できません。"
    Public Const LOG_HUBI_HYHENDO As String = "部屋情報から変動費を取得できません。"                                                                                  '20160603 ユーザーデータ検証による修正 -add

    Public Const LOG_HUBI_ERR_DATACHK_A As String = "データチェック用の最大最小値を取得時にエラーが発生しました。"
    Public Const LOG_HUBI_ERR_DATACHK_B As String = "データチェック用のフィールドサイズを取得時にエラーが発生しました。"
    Public Const LOG_HUBI_ERR_DATACHK_C As String = "データチェック用のキーナンバーを取得時にエラーが発生しました。"
  
    Public Const LOG_HUBI_RELFILEPIC As String = "関連ファイルに画像ファイルが含まれています。"                                                                       '20160525 クレーム関連ファイルの画像判別処理実装 -add
    Public Const LOG_HUBI_KYOYOKAGI As String = "共用鍵のデータです。"
    Public Const LOG_HUBI_SENYOKAGI As String = "専用鍵のデータです。"
    Public Const LOG_HUBI_GET_POSTCODE As String = "住所から郵便番号を取得しました。"                                                                                 '20160704 物件住所から郵便番号を読み込む処理の追加 -add
    Public Const LOG_HUBI_TAIONASI As String = "不要なデータです。"                                                                                                   '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -add
    Public Const LOG_HUBI_CHG_KANASTR As String = "口座名義用の文字列へ変更しました。"                                                                                '20160829 口座名義カナチェック機能の追加 -add
    Public Const LOG_HUBI_ERR_KANASTR As String = "半角カナへ変更できないため移行できません。"                                                                        '20160829 口座名義カナチェック機能の追加 -add
    Public Const LOG_HUBI_CHG_URLSTR As String = "URL、メールアドレス用の文字列へ変更しました。"                                                                      '20160829 メールアドレス、URLの正規化処理を追加 -add

    'ログ_文字列(対処方法)
    Public Const LOG_TAISYO_MIDHEADERCHK As String = "ヘッダーを元に戻す、または中間ファイルのバックアップを使用して再度作成して下さい。"
    Public Const LOG_TAISYO_MIDHEADERRANGECHK As String = "ヘッダーを元に戻す、ヘッダー範囲外データの除去、または中間ファイルのバックアップを使用して再度作成して下さい。"

    Public Const LOG_TAISYO_OVERLAP As String = "重複しないように再度キーを設定して下さい。"
    Public Const LOG_TAISYO_MISMATCH_KEY As String = "データタイプ(数値、文字列等)、または最大最小値を確認して下さい。"
    Public Const LOG_TAISYO_NOTEXISTDATA_REQUIRED As String = "必須項目に値を設定する、またはコンバート後に革命上で直接登録して下さい。"
    Public Const LOG_TAISYO_NOTEXISTDATA_BASE As String = "親マスタとの紐付けを確認して下さい。"

    Public Const LOG_TAISYO_OUTOOFRANGE_STR As String = "許容範囲内となるように値を設定する、またはコンバート後に革命上で直接編集して下さい。"
    Public Const LOG_TAISYO_GAIJI As String = "移行可能な文字列へ変更する、またはコンバート後に革命上で直接編集して下さい。"
    Public Const LOG_TAISYO_DEFAULT As String = "値を再度設定する、またはコンバート後に革命上で直接編集して下さい。"

    Public Const LOG_TAISYO_NOTEXISTDATA_TODOFUKEN As String = "マスタと一致するように値を設定して下さい。"
    Public Const LOG_TAISYO_DATACHK_A As String = "作業ファイルに不正があります。確認して下さい。(開発用)"
    Public Const LOG_TAISYO_DATACHK_B As String = "フィールドサイズ取得方法に不正があります。確認して下さい。(開発用)"

    Public Const LOG_TAISYO_RELFILEPIC As String = "画像コンバーターでコンバートを行って下さい。"                                                                     '20160525 クレーム関連ファイルの画像判別処理実装 -add
    Public Const LOG_TAISYO_KYOYOKAGI As String = "共用鍵として移行します。"
    Public Const LOG_TAISYO_SENYOKAGI As String = "専用鍵として移行します。"

    Public Const LOG_TAISYO_GET_POSTCODE As String = "設定された郵便番号を確認して下さい。"                                                                           '20160704 物件住所から郵便番号を読み込む処理の追加 -add
    Public Const LOG_TAISYO_TAIONASI As String = "対応はありません。"                                                                                                 '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -add

    'ログ_ログ挿入用仮テーブル名
    Public Const LOG_TMP_TABLENAME As String = "cv_log"
    Public Const LOG_TMP_MIDCHKTABLENAME As String = "midchk_log"

    'ログ_その他
    Public Const LOG_GET_KEYERRVALUE As String = "キーエラー"

#End Region

#Region "状況出力関連"

    Public Const SITUATION_MID_RUN As String = "中間ファイル書込中..."                       '20160825 進捗状況の表示修正 -add
    Public Const SITUATION_MID_READ As String = "中間ファイル読込中..."                      '20161007 実行ボタン押下→DB書込開始までの進捗表示処理追加 -add
    Public Const SITUATION_MID_CHKBEFORE As String = "中間ファイルチェック準備中..."         '20161011 コンバート実行時の進捗表示対応 -add
    Public Const SITUATION_MID_CHK As String = "中間ファイルチェック中..."                   '20161007 実行ボタン押下→DB書込開始までの進捗表示処理追加 -add
    Public Const SITUATION_MID_CHKAFTER As String = "中間ファイルチェック終了処理中..."      '20161011 コンバート実行時の進捗表示対応 -add
    Public Const SITUATION_CV_RUN As String = "コンバート処理中..."
    Public Const SITUATION_REL_RUN As String = "紐付設定中..."                               '20160825 進捗状況の表示修正 -add
    Public Const SITUATION_CV_BFRUN As String = "コンバート処理準備中 ..."                               '20161017 進捗表示ラベル初期表示修正 -add
    Public Const SITUATION_CV_STA As String = "コンバート処理開始"
    Public Const SITUATION_CV_END As String = "コンバート処理終了"
    Public Const SITUATION_CV_ENDMSG As String = "コンバートが正常に完了しました。" & vbCrLf & "出力されたログを確認して下さい。"
    Public Const SITUATION_CV_ERRENDMSG_A As String = "コンバートを正常に行うことができませんでした。" & vbCrLf & "出力されたログを確認して下さい。"
    Public Const SITUATION_STOP As String = "コンバート処理中断"
    Public Const SITUATION_STOP_END As String = "コンバート処理を中断しました。"
    Public Const SITUATION_EXTRACTION As String = "件数抽出中..."                            '20161014 改善対応：移行対象項目件数抽出処理改善 -add
    Public Const SITUATION_DATACONDITION As String = "中間ファイル調整中..."                 '20161014 改善対応：移行対象項目件数抽出処理改善 -add

    Public Const SITUATION_RELSTA As String = "紐付設定開始"
    Public Const SITUATION_RELEND As String = "紐付設定終了"
    Public Const SITUATION_READSTA As String = "移行元データ読込開始"
    Public Const SITUATION_READEND As String = "移行元データ読込終了"
    Public Const SITUATION_MIDSTA As String = "中間ファイル作成開始"
    Public Const SITUATION_MIDEND As String = "中間ファイル作成終了"
    Public Const SITUATION_MIDERREND As String = "中間ファイルの作成中にエラーが発生しました。" & vbCrLf & "接続情報に問題がある可能性があります。接続情報を確認し、再度コンバートを実行して下さい。"
    Public Const SITUATION_MIDFAIL As String = "中間ファイル作成失敗"
    Public Const SITUATION_MIDREADSTA As String = "中間ファイル読込開始"
    Public Const SITUATION_MIDREADEND As String = "中間ファイル読込終了"
    Public Const SITUATION_WRITESTA As String = "DB書込開始"
    Public Const SITUATION_WRITEEND As String = "DB書込終了"
    Public Const SITUATION_CVITEMSUCCESS As String = "○"
    Public Const SITUATION_CVITEMFAILURE As String = "×"
    Public Const SITUATION_CVITEMTOTALCNT As String = "全件数      "
    Public Const SITUATION_CVITEMCVCNT As String = "移行件数   "
    Public Const SITUATION_CVITEMCONDCNT As String = "調整件数   "
    Public Const SITUATION_CVITEMNOTCVCNT As String = "未移行件数"

#End Region

#Region "中間ファイル関連"

    '中間ファイルの最初のセル
    Public Const MIDFILE_READWRITE_CELLSTA As String = "A1"

    '中間ファイル読込/書込開始　※ヘッダの内容によって読み込み位置を変更
    Public Const EXISTMIDFILE_READWRITE_ROW As Integer = 2   '行                     
    Public Const EXISTMIDFILE_READWRITE_COL As String = "A"  '列

    '汎用用中間ファイル読込/書込開始
    Public Const BASEMIDFILE_HEADER_ROW_SETUBIGRP As Integer = 1                    '20160913_2 部屋設備移行処理の追加 -add
    Public Const BASEMIDFILE_HEADER_ROW As Integer = 2                              '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -add
    Public Const BASEMIDFILE_READWRITE_ROW As Integer = 12   '行                    '20160928 ヘッダー修正に伴う開始行修正 「補足」行の削除に伴い 13→12 へ変更
    Public Const BASEMIDFILE_READWRITE_COL As String = "B"   '列

    '紐付情報格納ファイル名
    Public Const MIDFILE_RELNAME As String = "各マスタ紐付け情報"

#End Region

#Region "DBから取得した情報を格納する共通オブジェクト"

    '革命10テーブル/フィールド情報取得用
    Public Hash_TblName_JpToAlpha As Hashtable                                      'テーブル名(日本語・アルファベット)
    Public Hash_TblName_AlphaToJp As Hashtable                                      'テーブル名(アルファベット・日本語)
    Public Hash_FldName_JpToAlpha As Hashtable                                      'フィールド名(日本語・アルファベット)
    Public Hash_FldName_AlphaToJp As Hashtable                                      'フィールド名(アルファベット・日本語)
    Public Hash_FiledTypeAlpha As Hashtable                                         'フィールド型(革命フィールド名)
    Public Hash_FiledSizeAlpha As Hashtable                                         'フィールドサイズ(革命フィールド名)
    Public Hash_FiledTypeJp As Hashtable                                            'フィールド型(日本語フィールド名)
    Public Hash_FiledSizeJp As Hashtable                                            'フィールドサイズ(日本語フィールド名)
    Public Hash_DefaultValue As Hashtable                                           'デフォルト値
    Public Hash_FiledKey As Hashtable                                               'キーに該当するフィールド
    Public Hash_Min_Code As Hashtable                                               '画面上で設定された最小値格納
    Public Hash_Max_Code As Hashtable                                               '画面上で設定された最大値格納
    Public Hash_Mst_ReferenceAlpha As Hashtable                                     '都道府県等、存在有無を確認する項目を格納
    Public Hash_FldName_JpToAlpha_Mid As Hashtable
    Public List_Required_Field As New List(Of String)                               'キーではないが革s命上で登録する際に必要な項目のフィールドを格納
    Public List_FldNameJp As New List(Of String)                                    '日本語フィールド名を格納

    Public List_Existmidheader As New List(Of String)
    Public Hash_ExistMiddatatype As New Hashtable
    Public Hash_ExistMiddatamin As New Hashtable
    Public Hash_ExistMiddatamax As New Hashtable
    Public Hash_ExistMiddatadef As New Hashtable
    Public Hash_ExistMiddatakey As New Hashtable
    Public List_ExistMiddatareq As New List(Of String)
    Public Hash_ExistMiddataref As New Hashtable
    Public List_ExistMiddatacv As New List(Of String)                               '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -add
    Public List_Basemidheader As New List(Of String)
    Public Hash_BaseMiddatatype As New Hashtable
    Public Hash_BaseMiddatamin As New Hashtable
    Public Hash_BaseMiddatamax As New Hashtable
    Public Hash_BaseMiddatadef As New Hashtable
    Public Hash_BaseMiddatakey As New Hashtable
    Public List_BaseMiddatareq As New List(Of String)
    Public Hash_BaseMiddataref As New Hashtable
    Public List_BaseMiddatacv As New List(Of String)                                '20160926 選定した移行項目をプログラムへ反映する修正 -add
    Public Hash_BaseMidToExistMid_FS As New Hashtable                               '20160905 中間ファイルコピー処理改善 -add
    Public Hash_ExistMidToBaseMid_FS As New Hashtable                               '20160905 中間ファイルコピー処理改善 -add
    Public Hash_BaseMidToExistMid_SH As New Hashtable                               '20160905 中間ファイルコピー処理改善 -add
    Public Hash_ExistMidToBaseMid_SH As New Hashtable                               '20160905 中間ファイルコピー処理改善 -add

    Public Hash_MidKey_FieldAlpha As Hashtable                                      '中間ファイルキー格納用
    Public Hash_MidKey_FieldJp As Hashtable                                         '中間ファイルキー格納用
    Public Hash_Kinyu As Hashtable                                                  '金融機関
    Public Hash_KinyuTen As Hashtable                                               '金融機関支店
    Public Hash_KagiTitleKyoyo As Hashtable                                         '共用鍵格納用
    Public Hash_KagiTitleSenyo As Hashtable                                         '専用鍵格納用
    Public Hash_KagiTitle_V7to10 As Hashtable                                       '鍵タイトル紐付要
    Public Hash_SetubiMid As Hashtable                                              '中間ファイルデータと設備Noを紐付 (V7設備データ-10設備No)
    Public Hash_SetubiMst As Hashtable                                              '10設備マスタ (設備No-項目guid)
    Public Hash_SetubiMst_UserMake As Hashtable                                     '20160829 設備の新規挿入処理を追加 -add
    Public Hash_ClaimBruiKasyo As Hashtable                                         'クレーム分類(箇所)
    Public Hash_ClaimBruiClaim As Hashtable                                         'クレーム分類(クレーム)

#End Region

#Region "紐付情報を格納する共通オブジェクト"

    Public Hash_Rel_Bkrui As Hashtable                                              '物件分類紐付情報
    Public Hash_Rel_Kozo As Hashtable                                               '構造マスタ紐付情報
    Public Hash_Rel_Hyrui As Hashtable                                              '部屋分類紐付情報
    Public Hash_Rel_Toritaiyo As Hashtable                                          '取引態様紐付情報
    Public Hash_Rel_Kozasyubetu As Hashtable                                        '口座種別紐付情報
    Public Hash_Rel_Nkinkomk As Hashtable                                           '入金項目紐付情報
    Public Hash_Rel_NkinkomkZksei As Hashtable                                      '入金項目属性紐付情報
    Public Hash_Rel_Nkinkbn As Hashtable                                            '入金区分紐付情報
    Public Hash_Rel_Hendometer As Hashtable                                         '各戸メーター分類紐付情報
    Public Hash_Rel_Setubi As Hashtable                                             '設備紐付情報
    Public Hash_Rel_Kyrui As Hashtable                                              '契約分類紐付情報
    Public Hash_Rel_Kyrui_Teikisyakuyakbn As Hashtable                              '契約分類紐付情報(定期借家区分用)
    Public Hash_Rel_Tosiyotokbn As Hashtable                                        '都市計画・用途地域区分情報
    Public Hash_Rel_Tosiyotono As Hashtable                                         '都市計画・用途地域No情報
    Public Hash_Rel_JisyaKoza As Hashtable                                          '自社口座情報取得用
    Public Hash_Rel_FBInfo_KozaFurikaeFmt As Hashtable                              'FBフォーマット(口座振替)
    Public Hash_Rel_FBInfo_FuriIraiFmt As Hashtable                                 'FBフォーマット(振込依頼人)
    Public Hash_Rel_FBInfo_NsSettingFmt As Hashtable                                'FBフォーマット(入手金設定)
    Public Hash_Rel_FBInfo_NsSettingJisya As Hashtable                              'FBフォーマット(入手金設定_自社)
   
#End Region

#Region "その他共通文字列群"

    Public Const INDENT_0 As String = "  "                                          '20160825 ファイルが開かれているかチェックする機能を追加 -add
    Public Const INDENT_1 As String = "    "
    Public Const INDENT_2 As String = "        "

    'VIEW作成クエリ共通構文
    Public Const PRE_VIEW_NAME As String = "CVVW_"
    Public Const PRE_VIEW_QRY As String = " CREATE VIEW "
    Public Const POST_VIEW_QRY As String = " AS "

    '仮テーブル作成クエリ共通構文
    Public Const PRE_TBL_NAME As String = "CVTBL_"

    'フォルダ、ファイル名
    Public Const DIR_EXEDIR_NAME As String = "Exe"                                  '本体/紐付Exe格納先フォルダ名
    Public Const DIR_CONVEXEDIR_NAME As String = "ConvMain"                         '本体Exe格納先フォルダ名
    Public Const DIR_RELEXEDIR_NAME As String = "RelationSetting"                    '紐付Exe格納先フォルダ名
    Public Const DIR_RELEXE_NAME As String = "RelationSetting.exe"                   '紐付Exe名
    Public Const DIR_FILEDIR_NAME As String = "File"                                '中間ファイル等、実行時に必要なファイルの格納先フォルダ名
    Public Const DIR_MAINLOG_NAME As String = "convlog"                             'ログファイル格納先フォルダ名
    Public Const DIR_MIDLOG_NAME As String = "midchklog"                            '中間ファイルログ格納先フォルダ名
    Public Const DIR_MID_NAME As String = "middlefile"                              '中間ファイル格納先フォルダ名
    Public Const DIR_REL_NAME As String = "relationfile"                            '紐付ファイル格納先フォルダ名
    Public Const DIR_RELLOG_NAME As String = "rellog"                               '紐付ツールログ格納先フォルダ名           '20160926 フォルダの自動生成処理を追加 -add
    Public Const DIR_TEMP_CSV As String = "csv"                                     'CSVファイル格納先フォルダ名
    Public Const DIR_INI_NAME As String = "ini"                                     '次回起動時用文字列保存ファイル格納先 (接続情報等)
    Public Const DIR_LIST_NAME As String = "list"                                   '事前作業で出力したリストの格納先フォルダ名
    Public Const FILE_V7_CONNAME As String = "coninfo_V7.xml"                       'V7接続情報保管ファイル
    Public Const FILE_10_CONNAME As String = "coninfo_10.xml"                       '10接続情報保管ファイル
    Public Const FILE_CVJISSEKINAME As String = "cvjisseki.xml"                     'コンバート実績保持用ファイル             '20160707 コンバート実績保持の処理追加 -add
    Public Const DIR_GAZOEXEDIR_NAME As String = "GazoConverter"                    '画像コンバートツール格納先フォルダ名     '20160711 画像CV呼出処理の追加 -add
    Public Const DIR_GAZOEXE_NAME As String = "GazoConverter.exe"                      '画像コンバートツール名                   '20160711 画像CV呼出処理の追加 -add
    Public Const MAIN_VIEW_NAME_BKGUID As String = "tmp_bk_guid"
    Public Const FILE_HMIDD_NAME As String = "中間ファイル.xlsx"                    '汎用中間ファイル   

    Public Const COMPANY_NAME As String = "n-create.co.jp"
    Public Const PRODUCT_NAME As String = "FK8Host"
    Public Const DEF_REC_USER As String = "CONVUSER"                                '作業者名のデフォルト値
    Public Const NJC_DEV As String = "_njc_dev"                                     'ツール起動時の開発用フラグ

    '革命10の改行コード
    Public Const LINE_BREAK As String = "$0D$0A"

    '日付の最大最小デフォルト値
    Public Const DEF_MIN_YMD As String = "1900/01/01"
    Public Const DEF_MAX_YMD As String = "2100/12/31"

    '区切文字文字列
    Public Const STR_SPLIT_1 As String = "@#@"                                      '20160913_2 「@#@」を共通変数に変更 -add
    Public Const STR_SPLIT_2 As String = "@$@"                                      '20160927 ログの内容が不正になっているため修正 -add

#End Region

#Region "その他共通オブジェクト"

    Public List_ImportableImageFileAttributes As New List(Of String)                '20160525 クレーム関連ファイルの画像判別処理実装 -add

#End Region

#Region "変数"

    Public DefHistory As String
    Public CancelFlg As Boolean = False
    Public MidChkCancelFlg As Boolean = False                                       '20160222 中間ファイルチェック時の中断処理の追加 -add
    Public Dev_CVFlg As Boolean = False                                             '開発用Converterフラグ
    Public Log_OutputCnt As Integer                                                 'ログ出力件数
    Public LogFilePath As String
    Public RelationDirPath As String
    Public MiddleDirPath As String
    Public MiddleLogFilePath As String
    Public BaseMidDirPath As String                                                 '汎用用中間ファイル格納先
    Public CV_FROM_MIDDLE As String                                                 '汎用用の中間ファイル名
    Public LimitTimeOut As Integer
    Public MsgResult As DialogResult
    Public RecUser As String
    Public InitDBFlg As Boolean                                                     'DB初期化フラグ
    Public OverWriteDBFlg As Boolean                                                '上書きフラグ(開発用)
    Public taihitab As Integer = 0                                                  '前回表示タブINDEX格納(開発タブONOFF用)
    Public UnyoYMD As String                                                        '20160525 運用開始年月の追加 -add
    '----- 要対応 ----- sta
    'Public CNVNO As Integer = ConvertTypes._既存ユーザ用
    Public CNVNO As Integer = ConvertTypes._汎用                                    '20160921 初期起動変更 
    '----- 要対応 ----- end

#End Region

End Module
