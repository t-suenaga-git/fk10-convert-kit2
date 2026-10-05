Module CommonModule

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
    Public Const LOG_SYU_BEGIN As String = "処理開始"
    Public Const LOG_SYU_CON As String = "移行元DB接続"
    Public Const LOG_SYU_READ As String = "移行元DB読込"
    Public Const LOG_SYU_FIN As String = "処理終了"
    Public Const LOG_SYU_MIDREAD As String = "CSV読込"
    Public Const LOG_SYU_STATYPE_A As String = "単体起動"
    Public Const LOG_SYU_STATYPE_B As String = "呼出起動"

    'ログ_文字列(処理結果内容)
    Public Const LOG_NAIYO_REL_STA As String = "紐付け処理開始"
    Public Const LOG_NAIYO_REL_END As String = "紐付け処理終了"
    Public Const LOG_NAIYO_NORMALEND As String = "正常終了"
    Public Const LOG_NAIYO_NOTNORMALEND As String = "異常終了"
    Public Const LOG_NAIYO_READSTA As String = "データ読込開始"
    Public Const LOG_NAIYO_READEND As String = "データ読込終了"
    Public Const LOG_NAIYO_WRITE_STA As String = "紐付データ出力開始"
    Public Const LOG_NAIYO_WRITE_END As String = "紐付データ出力終了"
    'Public Const LOG_NAIYO_OVERLAP As String = "重複データ"
    'Public Const LOG_NAIYO_STROVER As String = "文字数オーバー"

    Public Const LOG_NAIYO_ERR_OVERLAP As String = "重複"
    Public Const LOG_NAIYO_ERR_MISMATCH_KEY As String = "不適合(キー)"
    Public Const LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED As String = "必須項目無し"
    Public Const LOG_NAIYO_ERR_NOTEXISTDATA_BASE As String = "親データ無し"

    Public Const LOG_NAIYO_ERR_OUTOFRANGE As String = "有効範囲外"
    Public Const LOG_NAIYO_ERR_OUTOFRANGE_STR As String = "有効範囲外(文字数)"
    Public Const LOG_NAIYO_ERR_NOTEXISTDATA As String = "データ無し"
    Public Const LOG_NAIYO_ERR_MISMATCH As String = "不適合"

    Public Const LOG_NAIYO_ERR_NOTEXISTMSTDATA As String = "マスタ不一致"
    Public Const LOG_NAIYO_ERR_NOTEXISTFILEDATA As String = "参照ファイル無し"

    'ログ_文字列(処理項目)
    Public Const LOG_SYORI_BEGIN As String = ""

    'ログ_文字列(対象データ詳細)
    Public Const LOG_SYOSAI_BEGIN As String = ""

    'ログ_文字列(不備原因)
    Public Const LOG_HUBI_CON As String = "接続情報が間違っている可能性があります。"
    'Public Const LOG_HUBI_OVERLAP As String = "が重複しています。"
    Public Const LOG_HUBI_STROVER As String = "が許容文字数をオーバーしています。オーバーしている分は切り捨てます。"
    Public Const LOG_HUBI_MIDHEADERCHK As String = "ヘッダーの文字列が変更されている可能性があります。"
    Public Const LOG_HUBI_MIDHEADERRANGECHK As String = "ヘッダーの削除、またはヘッダー範囲外にデータが設定されている可能性があります。"
    Public Const LOG_HUBI_OVERLAP As String = "キーが重複しています。最初に処理する1件のみを紐付対象とします。"
    Public Const LOG_HUBI_MISMATCH_KEY As String = "キーの値が不正であるため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_REQUIRED As String = "革命でのデータ登録に必須となっている項目が存在しないため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_BASE As String = "親マスタにデータが存在しないため移行できません。"
    Public Const LOG_HUBI_OUTOFRANGE As String = "有効範囲外のデータです。"
    Public Const LOG_HUBI_OUTOFRANGE_STR As String = "有効範囲外のデータです。許容範囲の超過分を切り捨てて登録します。"
    Public Const LOG_HUBI_NOTEXISTDATA As String = "移行元データが存在しません。デフォルト値を設定します。"
    Public Const LOG_HUBI_MISMATCH As String = "移行元データが不正です。"
    Public Const LOG_HUBI_NOTEXISTDATA_MSTEXIST As String = "マスタとデータが一致しないため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_RELEXIST As String = "マスタと紐付けデータが一致しないため移行できません。"
    Public Const LOG_HUBI_NOTEXISTDATA_FILEEXIST As String = "参照ファイルが存在しないため移行できません。"
    Public Const LOG_HUBI_MAKEMIDFILE As String = "中間ファイル出力時にエラーが発生しました。"
    Public Const LOG_HUBI_MAKELOGFILE As String = "ログファイル出力時にエラーが発生しました。"
    Public Const LOG_HUBI_MAKEMIDCHKLOGFILE As String = "中間ファイルチェックログ出力時にエラーが発生しました。"


    'ログ_文字列(対処方法)
    Public Const LOG_TAISYO_STROVER As String = "許容文字数オーバー分を切り捨てます。"

    Public Const LOG_TAISYO_MIDHEADERCHK As String = "ヘッダーを元に戻す、または中間ファイルのバックアップを使用して再度作成して下さい。"
    Public Const LOG_TAISYO_MIDHEADERRANGECHK As String = "ヘッダーを元に戻す、ヘッダー範囲外データの除去、または中間ファイルのバックアップを使用して再度作成して下さい。"

    Public Const LOG_TAISYO_OVERLAP As String = "重複しないように再度キーを設定して下さい。"
    Public Const LOG_TAISYO_MISMATCH_KEY As String = "データタイプ(数値、文字列等)、または最大最小値を確認して下さい。"
    Public Const LOG_TAISYO_NOTEXISTDATA_REQUIRED As String = "必須項目に値を設定する、またはコンバート後に革命上で直接登録して下さい。"
    Public Const LOG_TAISYO_NOTEXISTDATA_BASE As String = "親マスタとの紐付けを確認して下さい。"

    Public Const LOG_TAISYO_OUTOOFRANGE_STR As String = "許容範囲内となるように値を設定する、またはコンバート後に革命上で直接編集して下さい。"
    Public Const LOG_TAISYO_DEFAULT As String = "値を再度設定して下さい。"

    Public Const LOG_TAISYO_NOTEXISTDATA_TODOFUKEN As String = "マスタと一致するように値を設定して下さい。"

    'ログ_文字列(設定値変更前)
    Public Const LOG_BEVALUE_BEGIN As String = ""

    'ログ_文字列(設定値変更後)
    Public Const LOG_AFVALUE_BEGIN As String = ""

    'ログ_文字列(対象TBL名)
    Public Const LOG_TABLE_BEGIN As String = ""

    'ログ_文字列(共通)
    Public Const LOG_NULL As String = "-"
    Public Const LOG_CHG As String = "を変更しました。"

    'ログ_文字列(出力日時)
    Public Log_Date As String = String.Format(Now)
    Public Log_Executor As String = "CONVUSER"

    'ログ_紐付ログ挿入用仮テーブル名
    Public Const LOG_MAIN_TABLENAME As String = "cv_log"        '本体コンバートログ挿入用仮テーブル      '2016.03.23 変数追加
    Public Const LOG_REL_TABLENAME As String = "cvrel_log"      '紐付けツールログ挿入用仮テーブル        '2016.03.23 変数名変更

#End Region

#Region "ファイル関連"

    Public Const REL_FILENAME As String = "各マスタ紐付け情報"
    Public Const CV_FROM_MIDDLE As String = "中間ファイル"        '汎用用中間ファイル名 '20161004 自社口座の口座種別取得処理の修正 -add
    Public MidDirPath As String = ""
    Public RelDirPath As String = ""
    Public LogDirPath As String = ""
    Public BaseMidDirPath As String = ""                        '20161004 自社口座の口座種別取得処理の修正 -add

    '中間ファイルの最初のセル
    Public Const MIDFILE_READWRITE_CELLSTA As String = "A1"

    '中間ファイル読込/書込開始
    Public Const MIDFILE_READWRITE_ROW As Integer = 2   '行
    Public Const MIDFILE_READWRITE_COL As String = "A"  '列
    Public Const RELFILESHEET_ADDNAME = "紐付情報"

    Public Const BASEMIDFILE_HEADER_ROW As Integer = 2       '汎用用中間ファイルのヘッダー行   '20161004 自社口座の口座種別取得処理の修正 -add
    Public Const BASEMIDFILE_READWRITE_ROW As Integer = 12   '汎用用中間ファイルの読込開始行   '20161004 自社口座の口座種別取得処理の修正 -add


    '2016.04.11 呼出起動の修正 -add sta
    Public Const DIR_RELEXEDIR_NAME As String = "RelationSetting"    '紐付Exe格納先フォルダ名
    Public Const DIR_MID_NAME As String = "middlefile"              '中間ファイル格納先フォルダ名
    Public Const DIR_REL_NAME As String = "relationfile"            '紐付ファイル格納先フォルダ名
    Public Const DIR_RELLOG_NAME As String = "rellog"             'ログファイル格納先フォルダ名
    '2016.04.11 呼出起動の修正 -add end

#End Region

#Region "変数"

    Public LimitTimeOut As Integer                                          'タイムアウト(秒)
    Public MsgResult As DialogResult                                        '返却メッセージ
    'Public RelationDirPath As String
    Public LogFilePath As String
    Public LogFileName As String = "\cvrel_log_" & (Replace(Replace((String.Format(Now)), ":", ""), "/", "")).Trim & ".csv"
    '2016.04.11 呼出起動の修正 -add
    Public CancelFlg As Boolean = False

#End Region

#Region "紐付項目関連"

    Public Const REL_ITEMNAME_BKRUI As String = "物件分類マスタ"
    Public Const REL_FLDNAME_BKRUI As String = "物件分類No"





    Public Hash_Bkrui_DataType As New Hashtable

#End Region

#Region "紐付作業用共通オブジェクト"

    Public RelItem_M_brui As New Object
    Public RelItem_M_crui As New Object
    Public RelItem_M_kozasyu_Rev7 As New Object
    Public RelItem_M_kozo_Rev7 As New Object
    Public RelItem_M_nkbn_Rev7 As New Object
    Public RelItem_M_nkin_Rev7 As New Object
    Public RelItem_M_setubi_Rev7 As New Object
    Public RelItem_M_toritaiyo_Rev7 As New Object
    Public RelItem_M_kagi_Rev7 As New Object
    Public RelItem_M_jisyakoza_Rev7 As New Object
    Public RelItem_M_gazo_title_Rev7 As New Object  '20160523 画像紐付設定処理の追加 -add
    Public RelItem_M_FBInfo_Furiirai_Rev7 As New Object       '20160525 FBフォーマット紐付設定処理の追加 -add
    Public RelItem_M_FBInfo_Kozafurikae_Rev7 As New Object       '20160525 FBフォーマット紐付設定処理の追加 -add
    Public RelItem_M_FBInfo_Nssetting_Rev7 As New Object       '20160525 FBフォーマット紐付設定処理の追加 -add
    Public RelItem_M_Keirui As New Object       '20160526 契約分類マスタの追加 -add
    Public RelItem_M_TosiYoto_Rev7 As New Object       '20160531 都市計画用途地域の追加 -add

    Public RelItem_M_bk_rui As New Object
    Public RelItem_M_hy_rui As New Object
    Public RelItem_M_kozasyu As New Object
    Public RelItem_M_kozo As New Object
    Public RelItem_M_nkbn As New Object
    Public RelItem_M_nkbn_z As New Object
    Public RelItem_M_nkin As New Object
    Public RelItem_M_nkin_z As New Object
    Public RelItem_M_nkin_hendometer As New Object
    Public RelItem_M_setubi_Grp As New Object
    Public RelItem_M_setubi_Ms As New Object
    Public RelItem_M_setubi_Komk As New Object
    Public RelItem_M_setubi As New Object
    Public RelItem_M_toritaiyo As New Object
    Public RelItem_M_kagi As New Object
    Public RelItem_M_jisyakoza As New Object
    '20160523 画像紐付設定処理の追加 -add sta
    Public RelItem_M_gazo_title_kbn As New Object
    Public RelItem_M_gazo_title_name As New Object
    '20160523 画像紐付設定処理の追加 -add end
    Public RelItem_M_FBInfo As New Object       '20160525 FBフォーマット紐付設定処理の追加 -add
    Public RelItem_M_ky_rui As New Object               '20160526 契約分類マスタの追加 -add
    Public RelItem_M_Tosi As New Object       '20160531 都市計画用途地域の追加 -add
    Public RelItem_M_Yoto As New Object       '20160531 都市計画用途地域の追加 -add

#End Region
    
    Public Const SITUATION_READ As String = "紐付項目を読込中です…"
    Public Const SITUATION_WRITE As String = "紐付設定値を出力中です…"

End Module



