''' <summary>
''' 【共通化設定 Module】
''' </summary>
''' <remarks>
''' ・全体で共通使用する設定値
''' </remarks>
Module RelationModule

#Region "定数"

    '設定値
    Public Const DefDate As String = "0:00:00"                              '日付デフォルト値
    Public Const DefFstDate As String = "1900/01/01"                        '許容範囲日付：開始
    Public Const DefLstDate As String = "2100/12/31"                        '許容範囲日付：末
    Public Const REL_FROM_NAME As String = "移行元"
    Public Const REL_TO_NAME As String = "賃貸革命10"
    'メッセージ
    Public Const MSG_SYURYO_B As String = "紐付作業が完了しました。"
    Public Const MSG_STOP_A As String = "紐付作業を中止してもよろしいですか？"
    Public Const MSG_STOP_B As String = "紐付作業を中止しました。"
    Public Const MSG_CANCEL_A As String = "紐付処理をキャンセルしてもよろしいですか？"
    Public Const MSG_CANCEL_B As String = "紐付処理をキャンセルしました。"
    Public Const MSG_RUN_A As String = "紐付処理を実行します。よろしいですか？"
    Public Const MSG_CNN_SUCCESS_B As String = "接続が正常に行えることを確認しました。"
    Public Const MSG_CNN_FAILURE_A As String = "DBへの接続に失敗しました。正しく設定されているか確認して下さい。"
    Public Const MSG_FILE_NOTHING_B As String = "ファイルパスが設定されていません。"
    Public Const MSG_DIR_CHK_A As String = "フォルダが見つかりませんでした。"
    Public Const MSG_CNN_NOTRUN_A As String = "接続テストを行って下さい。"
    Public Const MSG_CV_FAILURE_B As String = "紐付処理をが一部失敗しました。"
    Public Const MSG_RELRUN_B As String = "設定した条件で紐付を行います。よろしいですか？"
    Public Const MSG_DUPLICATE_A As String = "重複した値が設定されています。再度設定を行って下さい。"
    Public Const MSG_VALUE_FAILURE_A As String = "設定値が正しくありません。"
    Public Const MSG_CHK_BLANK As String = "紐付けが未設定の項目が存在します。全ての項目に対して紐付けを行って下さい。"
    Public Const MSG_CHK_BLANK_A As String = "以下のマスターにて、紐付けされていない項目はデフォルト値が設定されますがよろしいですか？"
    Public Const MSG_CHK_BLANK_B As String = "以下のマスターにて、紐付けされていない項目は新規に作成されますがよろしいですか？"
    Public Const MSG_CHK_BLANK_C As String = "以下のマスターにて、紐付けされていない項目は何も設定されませんがよろしいですか？"
    Public Const MSG_CHK_DUPLICATE As String = "重複不可の設定にします。よろしいですか？" & vbCrLf & "※重複不可にした場合、現在設定されている値で重複しているデータは削除されます"
    Public Const MSG_ERR_OUTOFRANGE As String = "有効範囲外です。値を再度設定して下さい。"  '2016.03.23 最大最小値チェック処理追加
    Public Const MSG_END_CALL As String = "紐付処理が完了しました。引き続きコンバートを行います。"

    Public Const MSG_READ_ITEM As String = "指定項目の読込を行います。よろしいですか？"
    Public Const MSG_DIRERR As String = "指定されたフォルダが見つかりません。"
    Public Const MSG_RELITEMERR As String = "紐付項目が選択されていません。"

    Public Const MSG_EXISTDATA_READ As String = "前回設定された紐付データを復元しますか？"    '20160527 前回設定値復元判別処理の追加 -add sta
    Public Const MSG_EXISTDATA_READ_DEV As String = vbCrLf & "※再コンバートを行う前に、賃貸革命V7のデータを" & vbCrLf & "　変更している場合は「いいえ」を選択して下さい。"             '20160527 前回設定値復元判別処理の追加 -add sta
    Public Const MSG_H_EXISTDATA_READ_DEV As String = vbCrLf & "※再コンバートを行う前に、中間ファイルのデータを" & vbCrLf & "　変更している場合は「いいえ」を選択して下さい。"         '20160725 汎用CKV対応 add

    Public Const MSG_ERR_SET_NKINNO As String = "入金項目開始Noが設定されていません。"                                  '20160622 入金項目No任意コード一括設定 -add
    Public Const MSG_STA_SET_NKINNO As String = "入金項目Noが未設定の項目へ一括設定を行います。よろしいですか？"    '20160622 入金項目No任意コード一括設定 -add
    Public Const MSG_END_SET_NKINNO As String = "一括設定が完了しました。"                                              '20160622 入金項目No任意コード一括設定 -add

    'Public Const MSG_ERR_SET_EXISTRELDATA As String = "が変更された可能性があるため前回設定値を復元できません。" & vbCrLf & "再度紐付設定を行って下さい。"        '20160622 前回設定値復元処理修正
    'Public Const MSG_ADD_SET_EXISTRELDATA As String = "に新たに追加された項目が存在します。" & vbCrLf & "新規に追加された項目の紐付設定を行って下さい。" & vbCrLf & "※現在構築中のため初期状態を読込みます。"          '20160622 前回設定値復元処理修正
    Public Const MSG_ERR_EXISTRELDATAREAD As String = "以下の紐付項目については、賃貸革命V7と前回の紐付設定の内容が異なる為、復元できません。" & vbCrLf & "(前回の紐付設定時から、賃貸革命V7の設定値が変更されている可能性があります)"        '20160622 前回設定値復元処理修正

    Public Const MSG_CHK_BLANK_REQ As String = "以下の項目で紐付け未設定の箇所が存在します。" & vbCrLf & "全項目について紐付設定を行って下さい。"      '20160705 未設定項目に関するメッセージ表示機能の追加 -chg '20160905 未設定項目がある場合のダイアログ修正 -chg
    Public Const MSG_CHK_BLANK_NOTREQ As String = "以下の項目で紐付未設定の箇所が存在します。" & vbCrLf & "紐付けされていない項目は新規に登録されますがよろしいですか？"      '20160705 未設定項目に関するメッセージ表示機能の追加 -chg '20160905 未設定項目がある場合のダイアログ修正 -chg

    Public Const MSG_ERR_READMID As String = "読込時にエラーが発生しました。処理を終了します。" '20161012 紐付ツール速度改善対応

    '20160725 紐付設定改善 add -sta
    '固定カラー
    Public escFColor As Color = Color.LightGray
    Public escBColor As Color = Color.Pink
    '20160825 紐付画面件数表示処理対応 -chg sta
    'Public escRelLabel As String = "lblTitleRelSelect"
    Public List_escRelLabel As New List(Of String) From {"lblTitleRelSelect", "lblTitleRelKomk", "lblTitleRelCnt", "lblTitleRelTotalCnt"}
    '20160825 紐付画面件数表示処理対応 -chg end
    Public CNVNO As Integer = ConvertTypes._既存ユーザ用            'コンバートタイプ (0.汎用, 1.既存ユーザ用, 2.他社システム用)
    '20160725 紐付設定改善 add -end

    '紐付項目
    Public Const REL_BKRUI As String = "物件分類マスタ"
    Public Const REL_HYRUI As String = "部屋分類マスタ"

#End Region


#Region "共通データ"
    '20160725 汎用CVK対応 add -sta
    Public Enum ConvertTypes
        _汎用 = 0
        _既存ユーザ用 = 1
        _他社システム用
        _不明 = 99
    End Enum
    '20160725 汎用CVK対応 add -end
#End Region


#Region "紐付項目関連"

    Public Hash_Bkrui As New Hashtable
    Public Hash_RelFile_Bkrui As New Hashtable
    Public Hash_RelDB_Bkrui As New Hashtable

#End Region


End Module
