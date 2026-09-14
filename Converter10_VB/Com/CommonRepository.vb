Imports System.Data.SqlClient
Imports Njc.Model
Imports Converter10.Njc.N3Lib.Utys
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Data.OleDb

Namespace Njc.Common

#Region "共通処理クラス"

    ''' <summary>
    ''' 共通処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CommonRepository

        ''' <summary>
        ''' 進捗率をラベルに表示
        ''' </summary>
        ''' <param name="pgbcnt"></param>
        ''' <param name="pgbtotalcnt"></param>
        ''' <remarks></remarks>
        Public Sub ProgressOutPut(ByVal pgbcnt As Integer, ByVal pgbtotalcnt As Integer, Optional ByVal totalflg As Boolean = False)

            With Njc.Frm.MainFrm

                If totalflg Then
                    .lblPgbTotal.Text = Int(((pgbcnt / pgbtotalcnt) * 100)).ToString & " %"
                Else
                    .lblPgbPartial.Text = Int(((pgbcnt / pgbtotalcnt) * 100)).ToString & " %"
                End If

            End With

        End Sub

        '20161009 改善対応：中間ファイルチェック時の進捗表示 -add sta
        '※とりあえずメイン進捗とは別に用意し、既存U用コンバーター対応時に修正する
        ''' <summary>
        ''' 進捗率をラベルに表示
        ''' </summary>
        ''' <param name="pgbcnt"></param>
        ''' <param name="pgbtotalcnt"></param>
        ''' <remarks></remarks>
        Public Sub ProgressChkOutPut(ByVal pgbcnt As Integer, ByVal pgbtotalcnt As Integer, Optional ByVal totalflg As Boolean = False)

            With Njc.Frm.MainFrm

                If totalflg Then
                    .lblPgbCheck.Text = Int(((pgbcnt / pgbtotalcnt) * 100)).ToString & " %"
                Else
                    '.lblPgbChkPartial.Text = Int(((pgbcnt / pgbtotalcnt) * 100)).ToString & " %"
                End If

            End With

        End Sub
        '20161009 改善対応：中間ファイルチェック時の進捗表示 -add end

    End Class

#End Region

#Region "タブページ表示設定クラス"

    ''' <summary>
    ''' タブページ表示処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TabPageManager

        '--------------------------------------------------
        ' 変数
        '--------------------------------------------------
        Private _tabPageInfos As TabPageInfo() = Nothing
        Private _tabControl As TabControl = Nothing


        ''' <summary>
        ''' TabPage情報セット
        ''' </summary>
        ''' <remarks></remarks>
        Private Class TabPageInfo
            Public objTabPage As TabPage
            Public blnVisible As Boolean

            ''' <summary>
            ''' インスタンス化
            ''' </summary>
            ''' <param name="page"></param>
            ''' <param name="visi"></param>
            ''' <remarks></remarks>
            Public Sub New(ByVal page As TabPage, ByVal visi As Boolean)
                objTabPage = page
                blnVisible = visi
            End Sub
        End Class

        ''' <summary>
        ''' TabPageManagerクラスのインスタンス化
        ''' </summary>
        ''' <param name="crl">基となるTabControlオブジェクト</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal crl As TabControl)
            _tabControl = crl
            _tabPageInfos = New TabPageInfo(_tabControl.TabPages.Count - 1) {}

            Dim ii As Integer
            For ii = 0 To _tabControl.TabPages.Count - 1
                _tabPageInfos(ii) = _
                    New TabPageInfo(_tabControl.TabPages(ii), True)
            Next ii
        End Sub

        ''' <summary>
        ''' TabPageの表示・非表示設定
        ''' </summary>
        ''' <param name="index">変更するTabPageのIndex番号</param>
        ''' <param name="visi">TABページ表示…TRUE.表示 FALSE.非表示</param>
        ''' <remarks></remarks>
        Public Sub ChangeTabPageVisible(ByVal index As Integer, ByVal visi As Boolean)
            If _tabPageInfos(index).blnVisible = visi Then
                Return
            End If
            _tabPageInfos(index).blnVisible = visi
            _tabControl.SuspendLayout()
            _tabControl.TabPages.Clear()

            Dim ii As Integer
            For ii = 0 To _tabPageInfos.Length - 1
                If _tabPageInfos(ii).blnVisible Then
                    _tabControl.TabPages.Add(_tabPageInfos(ii).objTabPage)
                End If
            Next ii
            _tabControl.ResumeLayout()
        End Sub

    End Class

#End Region

#Region "接続処理関連クラス"

    Public Class DBConnection

        Private _rtn As Boolean

        ''' <summary>
        ''' 接続処理 (OPEN)
        ''' </summary>
        ''' <param name="cnninfo">接続情報</param>
        ''' <param name="sqlcnn">接続 [out]</param>
        ''' <param name="flg_authent">認証方法flg</param>
        ''' <returns>接続状態<br/> True.成功<br/> False.失敗<br/></returns>
        ''' <remarks>・接続情報</remarks>
        Public Function CnnSession( _
                                   ByVal cnninfo As Njc.Model.DefSQLConnection, _
                                   ByRef sqlcnn As System.Data.SqlClient.SqlConnection, _
                                   ByVal flg_authent As Boolean
                                   ) As Boolean

            Dim cnnset As New System.Data.SqlClient.SqlConnection

            _rtn = True

            Try

                If flg_authent Then

                    'SQLServer認証
                    cnnset.ConnectionString = _
                         "Persist Security Info=True" & _
                         ";Data Source = " & cnninfo.ServerName & _
                         ";Initial Catalog = " & cnninfo.InitialCatalog & _
                         ";User ID = " & cnninfo.User & _
                         ";Password = " & cnninfo.Pass & _
                         ";Connection Timeout = " & cnninfo.TimeOut
                Else

                    'Windows認証
                    cnnset.ConnectionString = _
                         "Persist Security Info=True" & _
                         ";Data Source = " & cnninfo.ServerName & _
                         ";Initial Catalog = " & cnninfo.InitialCatalog & _
                         ";Integrated Security = SSPI" & _
                         ";Connection Timeout = " & cnninfo.TimeOut

                End If


                cnnset.Open()
                Console.WriteLine("{0}の{1}に接続しました", cnninfo.ServerName, cnninfo.InitialCatalog)

            Catch ex As Exception
                Console.WriteLine("Error! {0}", ex.Message)
                _rtn = False

            Finally
                sqlcnn = cnnset
                LimitTimeOut = cnninfo.TimeOut

            End Try
            Return _rtn

        End Function

        ''' <summary>
        ''' 接続処理(CLOSE:リソース解放)
        ''' </summary>
        ''' <param name="sqlcnn">接続情報</param>
        Public Sub CnnClose(ByVal sqlcnn As System.Data.SqlClient.SqlConnection)

            If Not sqlcnn Is Nothing Then
                If Me.CnnCheck(sqlcnn) = False Then
                    sqlcnn.Close()
                End If
                sqlcnn.Dispose()
            End If

        End Sub

        ''' <summary>
        ''' 接続状態をチェック
        ''' </summary>
        ''' <param name="sqlcnn">接続情報</param>
        ''' <returns>接続状態：<br/> True.接続<br/> False.非接続<br/></returns>
        Public Function CnnCheck(ByVal sqlcnn As System.Data.SqlClient.SqlConnection) As Boolean

            _rtn = True
            If sqlcnn.State <> ConnectionState.Closed Then
                _rtn = False
            End If
            Return _rtn

        End Function

    End Class

#End Region

#Region "インターフェース"

    Public Interface IConv

        '抽出クエリ(VIEW作成用)
        Function Get_UseQry(ByRef sortstr As String) As String

        '中間ファイル読込→変数→DB書込
        Function Set_Vari( _
                         ByVal sqlcnnv10 As SqlConnection, _
                         ByVal filename As String, ByVal sheetname As String, _
                         ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
                         ) As Boolean

        '親マスタ有無確認用オブジェクト作成
        Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String))

        '追加コンバート時の既存データのキーを取得してオブジェクトへ格納
        Sub Get_ExistData(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_existdata As List(Of String))

        'データチェック(特別にチェックが必要な場合)
        Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean

    End Interface

#End Region

#Region "プログレスバー表示設定クラス"

    Public Class ProgressBarManager

        ''' <summary>
        ''' 個別進捗プログレスバー初期設定
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbInitPart(ByVal int As Integer)

            With Njc.Frm.MainFrm.pgbpartial
                .Minimum = 0
                .Value = 0
                .Maximum = int
            End With

        End Sub

        ''' <summary>
        ''' 全体進捗プログレスバー初期設定
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbInitTotal(ByVal int As Integer)

            With Njc.Frm.MainFrm.pgbTotal
                .Minimum = 0
                .Value = 0
                .Maximum = int
            End With

        End Sub

        ''' <summary>
        ''' 個別進捗プログレスバー更新
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbsettingPart(ByVal int As Integer)

            With Njc.Frm.MainFrm.pgbpartial
                If int < .Maximum Then
                    .Value = int + 1
                    .Value = int
                Else
                    .Maximum = .Maximum + 1
                    .Value = int + 1
                    .Value = int
                    .Maximum = .Maximum - 1
                End If
            End With

        End Sub

        ''' <summary>
        ''' 全体進捗プログレスバー更新
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbsettingTotal(ByVal int As Integer)

            With Njc.Frm.MainFrm.pgbTotal
                If int < .Maximum Then
                    .Value = int + 1
                    .Value = int
                Else
                    .Maximum = .Maximum + 1
                    .Value = int + 1
                    .Value = int
                    .Maximum = .Maximum - 1
                End If
            End With

        End Sub

        '20161009 改善対応：中間ファイルチェック時の進捗表示 -add sta
        '※とりあえずメイン進捗とは別に用意し、既存ユーザ用コンバーター対応時に修正する
        ''' <summary>
        ''' 全体進捗プログレスバー初期設定
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbInitChkTotal(ByVal int As Integer)

            With Njc.Frm.MainFrm.pgbCheck
                .Minimum = 0
                .Value = 0
                .Maximum = int
            End With

        End Sub

        ''' <summary>
        ''' 全体進捗プログレスバー更新
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbsettingChkTotal(ByVal int As Integer)
            Debug.Print(Njc.Frm.MainFrm.pgbCheck.Maximum)
            With Njc.Frm.MainFrm.pgbCheck
                If int < .Maximum Then
                    .Value = int + 1
                    .Value = int
                Else
                    .Maximum = .Maximum + 1
                    .Value = int + 1
                    .Value = int
                    .Maximum = .Maximum - 1
                End If
            End With

        End Sub
        '20161009 改善対応：中間ファイルチェック時の進捗表示 -add end
    End Class

#End Region

#Region "データチェッククラス(ログ出力文字列を設定)"

    Public Class DataChk

        ''' <summary>
        ''' 日付チェック
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataDate(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim chgdate As DateTime

            'テスト用                                   'kakaka 日付チェックの内容(min_str～max_strの範囲)を仕様書に記載しておく
            '2016.03.28 日付の最大最小値の移動 -del sta
            'min_str = "1900/01/01"                      'kakaka4 CommonModule.vb へ移動(ここでもいいが、後で編集が予想されるものは1か所に纏めたい為)
            'max_str = "2100/12/31"
            '2016.03.28 日付の最大最小値の移動 -del end

            '2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            'Dim tmp_date_min As DateTime = DateTime.Parse(min_str)          'kakaka4 0322_1300 念の為、変換時の型落ちを考慮。
            'Dim tmp_date_max As DateTime = DateTime.Parse(max_str)          'kakaka4 0322_1300 念の為、変換時の型落ちを考慮。
            Dim tmp_date_min As DateTime
            Dim tmp_date_max As DateTime
            If DateTime.TryParse(min_str, tmp_date_min) = False Then
                tmp_date_min = DEF_MIN_YMD
            End If
            If DateTime.TryParse(max_str, tmp_date_max) = False Then
                tmp_date_max = DEF_MAX_YMD
            End If
            '2016.03.28 型落ちを考慮した処理へ修正 -chg end

            Dim def As String = defvalue

            '空文字チェック
            If chkstr = "" Then
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                'chkstrafter = def                       'kakaka 所有期間、契約期間など、チェックした期間に関連する期間がある場合はデフォルト設定は不可
                'If def <> "" Then
                '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                '    rtn = False
                'End If
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                '20160929 デフォルト値設定処理の追加 -add sta
                If def <> "" Then
                    chkstrafter = def
                Else
                    chkstrafter = chkstr
                End If
                '20160929 デフォルト値設定処理の追加 -add end
                Return rtn
            End If

            '日付変換チェック
            If Not (DateTime.TryParse(chkstr, chgdate)) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (tmp_date_min <= chgdate AndAlso chgdate <= tmp_date_max) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '20170524 竣工日がyyyy/mm/dd形式以外の日付の場合の移行処理対応 -chg sta
            ''正常値はそのまま格納
            'chkstrafter = chkstr
            'チェックで問題ない場合は変換した日付を返却する
            chkstrafter = chgdate.ToString
            '20170524 竣工日がyyyy/mm/dd形式以外の日付の場合の移行処理対応 -chg end

            '判定結果を返却
            Return rtn

        End Function

        ''' <summary>
        ''' 時間チェック
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataTime(ByVal chkstr As String, ByRef chkstrafter As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            '構築中

            Return rtn

        End Function

        ''' <summary>
        ''' 数値チェック(整数)
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <param name="max"></param>
        ''' <param name="min"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataNumeric(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim lng As Long = 0
            Dim rtn As Boolean = True
            '2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            'Dim min As Long = Int64.Parse(min_str)                      'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            'Dim max As Long = Int64.Parse(max_str)                      'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            Dim min As Long = 0
            Dim max As Long = 0
            If Int64.TryParse(min_str, min) = False OrElse Int64.TryParse(max_str, max) = False Then
                errstr = LOG_NAIYO_ERR_DATACHK & "-" & LOG_HUBI_ERR_DATACHK_A & "-" & LOG_TAISYO_DATACHK_A
                rtn = False
                Return rtn
            End If
            '2016.03.28 型落ちを考慮した処理へ修正 -chg end

            Dim def As String = defvalue

            '20161130 契約者入居フラグの移行処理修正_再 -chg sta
            '空文字チェックで「0」が含まれていたため空文字と「0」でチェックを分ける
            ''空文字チェック
            'If chkstr = "" Or chkstr = "0" Then     '20160603 ユーザーデータ検証による修正 条件に「0」を追加
            '    '20160829 送金予定日デフォルト値設定 -chg sta
            '    'chkstrafter = chkstr                '20160603 ユーザーデータ検証による修正 条件に「0」の追加に伴いチェック対象をそのまま返却する処理を追加
            '    If def <> "" Then
            '        chkstrafter = def
            '    Else
            '        chkstrafter = chkstr
            '    End If
            '    '20160829 送金予定日デフォルト値設定 -chg end
            '    '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
            '    'chkstrafter = def
            '    'If def <> "" Then
            '    '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT        'kakaka 中間ファイルチェックの場合でもデフォルト値に設定と表示される(以下チェック、他メソッドも同様)
            '    '    rtn = False
            '    'End If
            '    '2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
            '    Return rtn
            'End If

            '空文字チェック
            If chkstr = "" Then
                If def <> "" Then
                    'デフォルト値がある場合は設定する
                    chkstrafter = def
                End If
                Return rtn
            End If

            '「0」チェック
            '--------------------------------- ↓「0」チェックの内容↓ ---------------------------------
            '最小値が「0」以下の場合は「0」を値として移行する
            '最小値が「0」より大きい場合は値として意味をなさないので空データで移行する
            '※物件情報に登録されている業者情報No等は画面上で未設定でもDB上で「0」として登録されている
            '　業者情報No等は1から始まる値なので値は空データとして移行する必要がある
            '--------------------------------- ↑「0」チェックの内容 ↑ ---------------------------------
            If chkstr = "0" Then
                If min <= 0 Then
                    '最小値が「0」以下の場合はそのまま移行
                    chkstrafter = chkstr
                    Return rtn
                ElseIf min > 0 Then
                    '最小値が「0」より大きい場合は空データを移行
                    chkstrafter = ""
                    Return rtn
                End If
            End If
            '20161130 契約者入居フラグの移行処理修正_再 -chg end

            '数値変換チェック
            If Not (Int64.TryParse(chkstr, lng)) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= lng AndAlso lng <= max) Then             'kakaka Or→OrElse、And→AndAlsoに変更してください。(他のチェックのところも)
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '正常値はそのまま格納
            chkstrafter = chkstr

            '判定結果を返却
            Return rtn

        End Function

        ''' <summary>
        ''' 数値チェック(Decimal)
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <param name="max"></param>
        ''' <param name="min"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataDecimal(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim dec As Decimal = 0
            Dim rtn As Boolean = True
            '2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            'Dim min As Decimal = Decimal.Parse(min_str)             'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            'Dim max As Decimal = Decimal.Parse(max_str)             'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            Dim min As Decimal = 0
            Dim max As Decimal = 0
            If Decimal.TryParse(min_str, min) = False OrElse Decimal.TryParse(max_str, max) = False Then
                errstr = LOG_NAIYO_ERR_DATACHK & "-" & LOG_HUBI_ERR_DATACHK_A & "-" & LOG_TAISYO_DATACHK_A
                rtn = False
                Return rtn
            End If
            '2016.03.28 型落ちを考慮した処理へ修正 -chg end

            Dim def As String = defvalue

            '空文字チェック
            If chkstr = "" Then
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                'chkstrafter = def
                'If def <> "" Then
                '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                '    rtn = False
                'End If
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                '20160929 デフォルト値設定処理の追加 -add sta
                If def <> "" Then
                    chkstrafter = def
                Else
                    chkstrafter = chkstr
                End If
                '20160929 デフォルト値設定処理の追加 -add end
                Return rtn
            End If

            '数値変換チェック
            If Not (Decimal.TryParse(chkstr, dec)) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= dec AndAlso dec <= max) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '正常値はそのまま格納
            chkstrafter = chkstr

            '判定結果を返却
            Return rtn

        End Function

        ''' <summary>
        ''' 数値チェック(小数)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataDouble(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim dbl As Double = 0
            Dim rtn As Boolean = True
            '2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            'Dim min As Double = Double.Parse(min_str)                   'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            'Dim max As Double = Double.Parse(max_str)                   'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            Dim min As Double = 0
            Dim max As Double = 0
            If Double.TryParse(min_str, min) = False OrElse Double.TryParse(max_str, max) = False Then
                errstr = LOG_NAIYO_ERR_DATACHK & "-" & LOG_HUBI_ERR_DATACHK_A & "-" & LOG_TAISYO_DATACHK_A
                rtn = False
                Return rtn
            End If
            '2016.03.28 型落ちを考慮した処理へ修正 -chg end

            Dim def As String = defvalue

            '空文字チェック
            If chkstr = "" Then
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                'chkstrafter = def
                'If def <> "" Then
                '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                '    rtn = False
                'End If
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                '20160929 デフォルト値設定処理の追加 -add sta
                If def <> "" Then
                    chkstrafter = def
                Else
                    chkstrafter = chkstr
                End If
                '20160929 デフォルト値設定処理の追加 -add end
                Return rtn
            End If

            '数値変換チェック
            If Not (Double.TryParse(chkstr, dbl)) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= dbl AndAlso dbl <= max) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                'errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '正常値はそのまま格納
            chkstrafter = chkstr

            '判定結果を返却
            Return rtn

        End Function

        ''' <summary>
        ''' 数値チェック(通貨)
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <param name="max"></param>
        ''' <param name="min"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMoney(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim dec As Decimal = 0
            Dim rtn As Boolean = True
            '2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            'Dim min As Decimal = Decimal.Parse(min_str)                 'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            'Dim max As Decimal = Decimal.Parse(max_str)                 'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            Dim min As Decimal = 0
            Dim max As Decimal = 0
            If Decimal.TryParse(min_str, min) = False OrElse Decimal.TryParse(max_str, max) = False Then
                errstr = LOG_NAIYO_ERR_DATACHK & "-" & LOG_HUBI_ERR_DATACHK_A & "-" & LOG_TAISYO_DATACHK_A
                rtn = False
                Return rtn
            End If
            '2016.03.28 型落ちを考慮した処理へ修正 -chg end

            Dim def As String = defvalue
            Dim tmp_str As String = ""

            '空文字チェック
            If chkstr = "" Then
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                'chkstrafter = def
                'If def <> "" Then
                '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                '    rtn = False
                'End If
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                '20160929 デフォルト値設定処理の追加 -add sta
                If def <> "" Then
                    chkstrafter = def
                Else
                    chkstrafter = chkstr
                End If
                '20160929 デフォルト値設定処理の追加 -add end
                Return rtn
            End If

            '半角変換
            '20160531 不正金額の移行制御修正 -chg sta
            'tmp_str = EtcMethod.Get_LenB(EtcMethod.Chg_NumNarrow(chkstr))
            tmp_str = EtcMethod.Chg_NumNarrow(chkstr)
            '20160531 不正金額の移行制御修正 -chg end

            'カンマチェック
            If InStr(tmp_str, ",") Then
                tmp_str = Replace(tmp_str, ",", "")
            End If

            '数値変換チェック
            If Not (Decimal.TryParse(tmp_str, dec)) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def                                               'kakaka moneyのデフォルト値は0円？
                ''20160531 不正金額の移行制御修正 -add
                'errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def                                               'kakaka moneyのデフォルト値は0円？
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= dec AndAlso dec <= max) Then
                '20161014 ログ出力内容修正 -chg sta
                'chkstrafter = def
                ''20160531 不正金額の移行制御修正 -add
                'errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                If def <> "" Then
                    chkstrafter = def
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                Else
                    errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE_NODEF & "-" & LOG_TAISYO_DEFAULT
                End If
                '20161014 ログ出力内容修正 -chg end
                rtn = False
                Return rtn
            End If

            '正常値はそのまま格納
            chkstrafter = chkstr

            '判定結果を返却
            Return rtn

        End Function

        ''' <summary>
        ''' 論理値チェック
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataBit(ByVal chkstr As String, ByRef chkstrafter As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim int As Integer = 0
            Dim rtn As Boolean = True
            Dim def As String = defvalue

            '空文字チェック
            If chkstr = "" Then
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                'chkstrafter = def
                'If def <> "" Then
                '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                '    rtn = False
                'End If
                '2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                '20160929 デフォルト値設定処理の追加 -add sta
                If def <> "" Then
                    chkstrafter = def
                Else
                    chkstrafter = chkstr
                End If
                '20160929 デフォルト値設定処理の追加 -add end
                Return rtn
            End If

            '数値変換チェック
            If Not (Int32.TryParse(chkstr, int)) Then
                chkstrafter = def
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If int >= 2 Then
                chkstrafter = def
                rtn = False
                Return rtn
            End If

            '正常値はそのまま格納
            chkstrafter = chkstr

            '判定結果を返却
            Return rtn

        End Function

        ''' <summary>
        ''' 文字列チェック
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <param name="max"></param>
        ''' <param name="strtype"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataString(ByVal chkstr As String, ByRef chkstrafter As String, ByVal max As Integer, ByVal strtype As Integer, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            'サイズチェック
            If max = -1 Then
                chkstrafter = chkstr
                Return rtn
            End If

            '空文字チェック
            If chkstr = "" Then
                chkstrafter = chkstr
                Return rtn
            End If

            '英数文字を半角変換してバイト数取得                           'kakaka この辺の条件を表など一覧にして仕様書に記載しておく
            Dim cnt_byte As Integer = EtcMethod.Get_LenB(EtcMethod.Chg_NumNarrow(chkstr))

            '許容範囲チェック
            If cnt_byte > max Then
                If strtype = 0 Then
                    'SJIS用
                    chkstrafter = N3Lib.Utys.Typ.ToStrSafe(chkstr, max)                     'kakaka メモ：文字列調整
                Else
                    'Unicode用
                    chkstrafter = N3Lib.Utys.Typ.ToStrSafeUnicode(chkstr, max)
                End If
            Else
                chkstrafter = chkstr
            End If

            If chkstr <> chkstrafter Then
                errstr = LOG_NAIYO_ERR_OUTOFRANGE_STR & "-" & LOG_HUBI_OUTOFRANGE_STR & "-" & LOG_TAISYO_OUTOOFRANGE_STR
                rtn = False
            End If

            '外字チェック
            '外字チェック処理
            '変換処理をして作業用変数へ格納(未実装)                       'kakaka 10の導入先がV7と同じはずなので、外字はクリアされると思うが、汎用を考えて入れておく
            '2016.04.04 外字チェック処理追加対応 -add sta
            '許容範囲内にした状態でチェックを行う(切断された部分に外字が含まれている場合は変換する必要がないため)

            Dim tmp_gaijichkstr As String = chkstrafter
            Dim cnt_str As Integer = chkstrafter.Length

            '1文字ずつ抽出して照合
            For cntii = 1 To cnt_str

                Dim tmp_chkstr As String = Mid(tmp_gaijichkstr, cntii, 1)

                'アプリで処理できない文字
                If (Char.GetUnicodeCategory(tmp_chkstr) = Globalization.UnicodeCategory.OtherNotAssigned) Then
                    tmp_gaijichkstr = tmp_gaijichkstr.Replace(tmp_chkstr, "■")     '外字を置換
                End If

                'サロゲートペア
                If (Char.IsHighSurrogate(tmp_chkstr) = True Or Char.IsLowSurrogate(tmp_chkstr) = True) Then
                    tmp_gaijichkstr = Left(tmp_gaijichkstr, cntii - 1) & "■" & Mid(tmp_gaijichkstr, cntii + 2, cnt_str - cntii)     '外字を置換
                End If

                'Unicode文字
                If (Char.IsHighSurrogate(tmp_chkstr) = False And Char.IsLowSurrogate(tmp_chkstr) = False) Then
                    Dim bytes() As Byte = Nothing
                    Dim encBytes() As Byte = Nothing
                    Dim chgAfterStr As String = ""
                    Dim _sjisEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

                    '文字列をBytes配列に変換 (unicode -> bytes()unicode)
                    bytes = _sjisEncoding.GetBytes(tmp_chkstr)
                    'Bytes配列の中身をShift_JISに変換 (bytes()unicode -> bytes()Shift_JIS)
                    encBytes = System.Text.Encoding.Convert(_sjisEncoding, System.Text.Encoding.Unicode, bytes)
                    'Bytes配列を文字列に変換 (bytes()Shift_JIS -> unicode)
                    chgAfterStr = System.Text.Encoding.Unicode.GetString(encBytes)

                    If (tmp_chkstr.ToString <> "?" And chgAfterStr = "?") Then
                        tmp_gaijichkstr = tmp_gaijichkstr.Replace(tmp_chkstr, "■")     '外字を置換
                    End If

                    bytes.Initialize()
                    encBytes.Initialize()
                    chgAfterStr = ""
                End If

            Next

            If chkstrafter <> tmp_gaijichkstr Then
                errstr = LOG_NAIYO_ERR_GAIJI & "-" & LOG_HUBI_GAIJI & "-" & LOG_TAISYO_GAIJI
                rtn = False
                chkstrafter = tmp_gaijichkstr
            End If
            '2016.04.04 外字チェック処理追加対応 -add end

            Return rtn

        End Function

        ''' <summary>
        ''' 口座名義カナチェック 20160706 半角カナの移行処理修正
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_Kana(ByVal chkstr As String, ByRef chkstrafter As String, ByRef errstr As String) As Boolean

            '20160829 口座名義カナチェック機能の追加 -chg sta
            'Dim rtn As Boolean = True
            'Dim tmp_str As String = chkstr

            ''構築中

            ''空文字チェック
            'If chkstr = "" Then
            '    Return rtn
            'End If

            ''カタカナ変換
            'tmp_str = StrConv(tmp_str, VbStrConv.Katakana)

            ''半角変換
            'tmp_str = StrConv(tmp_str, VbStrConv.Narrow)

            ''半角カナチェック

            '↓↓↓ここから↓↓↓
            Dim rtn As Boolean = True
            Dim tmp_str As String = ""

            '空文字チェック
            If chkstr = "" Then
                Return rtn
            End If

            '口座名義用に変換
            Dim kana_obj As New DataChk
            tmp_str = kana_obj.ToStrKozaKana(chkstr)

            '口座名義用に変換できなかった場合
            'If normalflg = False Then
            '    errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_ERR_KANASTR & "-" & LOG_TAISYO_DEFAULT
            '    rtn = normalflg
            '    Return rtn
            'End If

            '口座名義用に変換可、かつ変換されている場合
            If chkstr <> tmp_str Then
                errstr = LOG_NAIYO_CHG_STR & "-" & LOG_HUBI_CHG_KANASTR & "-" & ""
                rtn = False
            End If

            '戻り値を格納
            '20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -chg sta
            '大文字変換処理を行う
            'chkstrafter = tmp_str
            chkstrafter = kana_obj.Get_KozaKanaUpper(tmp_str)
            '20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -chg end
            '20160829 口座名義カナチェック機能の追加 -chg end

            Return rtn

        End Function

        ''' <summary>
        ''' 口座名義カナ用に変換する(革命10から引用) '20160829 口座名義カナチェック機能の追加 -add
        ''' </summary>
        ''' <param name="value">対象文字列</param>
        ''' <returns>変換後の文字列</returns>
        ''' <remarks>
        ''' 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正
        ''' 　大文字変換処理の削除に伴い引数「Optional wohenkan As Boolean」を削除
        ''' </remarks>
        Public Function ToStrKozaKana(ByVal value As Object) As String
            Dim result As String = ToStr(value)

            'ひらがな->カタカナ
            '全角->半角
            '小文字 -> 大文字
            result = StrConv(result, VbStrConv.Katakana Or VbStrConv.Narrow Or VbStrConv.Uppercase)

            '20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -del sta
            'ログ出力制御のためここでは変換しない
            ''カナのチェック
            'Dim replacefrom As String = _kozakanareplacefrom
            'Dim replaceto As String = _kozakanareplaceto
            'If wohenkan = False Then
            '    replacefrom = _kozakanareplacefrom2
            '    replaceto = _kozakanareplaceto2
            'End If
            'For thisstep = 0 To replacefrom.Length - 1
            '    Dim checkfrom = replacefrom(thisstep)
            '    If (0 <= result.IndexOf(checkfrom)) Then
            '        Dim checkto = replaceto(thisstep)
            '        result = result.Replace(checkfrom, checkto)
            '    End If
            'Next
            '20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -del end

            '使用できない文字のチェック
            For thisstep = 0 To _kozakanarareplaceclear.Length - 1
                Dim checkclear = _kozakanarareplaceclear(thisstep)
                If (0 <= result.IndexOf(checkclear)) Then
                    result = result.Replace(checkclear, String.Empty)
                End If
            Next

            '2bytes文字が残っていたら除去
            Dim bytes() = _sjisEncoding.GetBytes(result)
            If (bytes.Length <> result.Length) Then
                Dim target = result
                Dim resultsb = New System.Text.StringBuilder
                For Each this In target
                    If (_sjisEncoding.GetBytes(this).Length = 1) Then
                        resultsb.Append(this)
                    End If
                Next
                result = resultsb.ToString
            End If

            Return result
        End Function

        ''' <summary>
        ''' 指定文字列を大文字へ変換する処理
        ''' </summary>
        ''' <param name="value">対象文字列</param>
        ''' <param name="wohenkan">「ｦ」→「ｵ」変換フラグ True:変換する False:変換しない</param>
        ''' <returns>大文字変換後の文字列</returns>
        ''' <remarks>
        ''' 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 新規追加
        ''' 　大文字へ変換した場合の変換ログは不要なためこの処理だけ別メソッドにして個別で処理する
        ''' 　処理内容は革命10本体から引用
        ''' </remarks>
        Public Function Get_KozaKanaUpper(ByVal value As String, Optional wohenkan As Boolean = True) As String

            Dim result As String = ToStr(value)

            Dim replacefrom As String = _kozakanareplacefrom
            Dim replaceto As String = _kozakanareplaceto
            If wohenkan = False Then
                replacefrom = _kozakanareplacefrom2
                replaceto = _kozakanareplaceto2
            End If
            For thisstep = 0 To replacefrom.Length - 1
                Dim checkfrom = replacefrom(thisstep)
                If (0 <= result.IndexOf(checkfrom)) Then
                    Dim checkto = replaceto(thisstep)
                    result = result.Replace(checkfrom, checkto)
                End If
            Next

            Return result

        End Function

        '20160829 口座名義カナチェック機能の追加 革命10から引用 -add sta
        Private ReadOnly _sjisEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        Private _kozakanareplacefrom As String = "ｧｨｩｪｫｬｭｮｯｰ･ヵヶｦ"
        Private _kozakanareplaceto As String = "ｱｲｳｴｵﾔﾕﾖﾂ-.ｶｹｵ"
        Private _kozakanarareplaceclear As String = "@`!""#$%&'*+:;[{<|=]}>^~?_｡､"
        Private _kozakanareplacefrom2 As String = "ｧｨｩｪｫｬｭｮｯｰ･ヵヶ"
        Private _kozakanareplaceto2 As String = "ｱｲｳｴｵﾔﾕﾖﾂ-.ｶｹ"
        '20160829 口座名義カナチェック機能の追加 革命10から引用 -add end

        '20161005 自社、家主口座デフォルト値設定処理追加 -del sta
        '家主口座チェック処理と統合するためコメントアウト
        ' ''' <summary>
        ' ''' 自社口座有無の確認 '20160829 自社口座にデフォルト値を設定する処理を追加 -add
        ' ''' </summary>
        ' ''' <param name="chkstr"></param>
        ' ''' <param name="chkstrafter"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_Jisyakoza(ByVal chkstr As String, ByRef chkstrafter As String, ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True

        '    '空の場合は「1：みずほ銀行」をセット
        '    If chkstr = "" Or Chk_DataNumeric(chkstr, chkstrafter, "1", "9999", "", errstr) = False Then     '20160928 自社口座取得方法の修正 条件追加
        '        errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
        '        chkstrafter = "1"
        '        rtn = False
        '        Return rtn
        '    End If

        '    chkstrafter = chkstr

        '    Return rtn

        'End Function
        '20161005 自社、家主口座デフォルト値設定処理追加 -del end

        ''' <summary>
        ''' 自社、家主口座有無の確認 '20161005 自社、家主口座デフォルト値設定処理追加 -add
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <param name="chkstrafter"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_JisyaOwkoza(ByVal chkstr As String, ByRef chkstrafter As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            '20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正(コメント追記) -add sta
            '空の場合は「9999:ダミー口座」をセット(DB書込み前に予め作成したデータ)
            'この処理は家賃入金口座情報作成時に、元となる自社口座情報に金融機関情報が存在しないと保存できない
            '状態になるのを回避するために仮で口座情報を設定する
            '金融機関Noを設定することで保存可能となるため金融機関Noが空の場合は仮口座を設定する
            '20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正(コメント追記) -add end

            If chkstr = "" Or Chk_DataNumeric(chkstr, chkstrafter, "1", "9999", "", errstr) = False Then
                '20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -chg sta
                'Dim tmp_str As String = erritem & "の金融機関情報が存在しないためデフォルト値「1」を設定します。"
                'errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_str & "-" & LOG_TAISYO_DEFAULT
                'chkstrafter = "1"
                Dim tmp_str As String = erritem & "の金融機関情報が存在しないためデフォルト値「9999」を設定します。"
                errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_str & "-" & LOG_TAISYO_DEFAULT
                chkstrafter = "9999"
                '20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -chg end
                rtn = False
                Return rtn
            End If

            chkstrafter = chkstr

            Return rtn

        End Function

        ''' <summary>
        ''' 箇所分類チェック 20160829 箇所未登録データにデフォルト値を設定 -add
        ''' 「箇所分類」+「行No」の形でデータをチェックする
        ''' データが数値の場合は箇所分類が存在しない場合はデフォルト値をセットする
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <param name="chkstrafter"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ChK_KasyoBrui(ByVal chkstr As String, ByRef chkstrafter As String, ByRef errstr As String) As Boolean

            Dim tmp_int As Integer = 0

            '数値変換チェック
            If Int32.TryParse(chkstr, tmp_int) = False Then
                Return True
            End If

            '数値のみの場合は箇所分類が未設定なのでデフォルト値をセット
            errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
            chkstrafter = "箇所未登録" & tmp_int.ToString
            Return False

        End Function

        ''' <summary>
        ''' URL、メールアドレスをデータベース保存用に変換 20160829 メールアドレス、URLの正規化処理を追加
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <param name="chkstrafter"></param>
        ''' <param name="length"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_URL(ByVal chkstr As String, ByRef chkstrafter As String, ByVal length_str As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            '※構築中

            Dim tmp_str As String = ""
            Dim length_int As Integer = 0
            Dim tmp_int As Integer = 0

            '空文字チェック
            If chkstr = "" Then
                Return rtn
            End If

            ''最大サイズを数値変換
            'If Int32.TryParse(length_str, tmp_int) Then
            '    length_int = tmp_int
            'End If

            ''半角変換、半角変換不可文字は半角スペースに置換
            'Dim kana_obj As New DataChk
            'tmp_str = kana_obj.ToStrHankakuKana(chkstr)

            ''データベース保存用に変換
            'chkstrafter = ToStrSafeUrlEncode(tmp_str, length_int)

            ''変換後のチェック
            'If chkstr <> chkstrafter Then
            '    errstr = LOG_NAIYO_CHG_STR & "-" & LOG_HUBI_CHG_URLSTR & "-" & LOG_TAISYO_DEFAULT
            '    chkstrafter = chkstrafter
            '    rtn = False
            '    Return rtn
            'End If

            chkstrafter = chkstr

            Return rtn

        End Function

        ''' <summary>
        ''' 半角カナに変換する(革命10から引用) '20160829 メールアドレス、URLの正規化処理を追加 -add
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="flg"></param>
        ''' <param name="wohenkan"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ToStrHankakuKana(ByVal value As Object) As String
            Dim result As String = ToStr(value)

            'ひらがな->カタカナ
            '全角->半角
            result = StrConv(result, VbStrConv.Katakana Or VbStrConv.Narrow)

            '2bytes文字が残っていたら除去
            Dim bytes() = _sjisEncoding.GetBytes(result)
            If (bytes.Length <> result.Length) Then
                Dim target = result
                Dim resultsb = New System.Text.StringBuilder
                For Each this In target
                    If (_sjisEncoding.GetBytes(this).Length = 1) Then
                        resultsb.Append(this)
                    Else
                        resultsb.Append(" ")
                    End If
                Next
                result = resultsb.ToString
            End If

            Return result
        End Function

        ''' <summary>
        ''' 半角英数と一部記号以外の文字を^22のような問題のない文字に置換します。<br/>
        ''' データベース保存用です(Web用には使えません。)<br/>
        ''' 20160829 メールアドレス、URLの正規化処理を追加
        ''' System.Web.HttpUtilityを有効にするために参照の追加で「System.Web」を追加
        ''' </summary>
        Public Shared Function ToStrSafeUrlEncode(ByVal value As Object, ByVal length As Integer) As String
            Dim result As String = Typ.ToStr(value)
            If (IsStrMissing(value) = False) Then
                result = System.Web.HttpUtility.UrlEncode(SubstrByte(value, 0, length))
                result = result.Replace("%", "^")
            End If
            Return result
        End Function

        ''' <summary>
        ''' 各マスタ有無チェック '20160530 ログ修正
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist(ByVal sqlcnnv10 As SqlConnection, ByVal sql_select As String, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'マスタからコードを取得
            Dim tmp_list As New List(Of String)
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

            'データが存在しない場合は処理を抜ける
            If value = "" Or value = "0" Then
                Return rtn
            End If

            flg = DBExec.Exec_DataReader_Col_List(sql_select, sqlcnnv10, tmp_list)

            If tmp_list.Contains(value) Then
                chkvalue = value
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_RELEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 都道府県マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_Todofuken(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'マスタからコードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT ken_no FROM m_ken"
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    '20160530 ログ修正 -add sta
        '    If value = "" Then
        '        Return rtn
        '    End If
        '    '20160530 ログ修正 -add end

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_RELEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 市区町村マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_Sikucyoson(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'マスタからコードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT si_no FROM m_si "
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    '20160530 ログ修正 -add sta
        '    If value = "" Then
        '        Return rtn
        '    End If
        '    '20160530 ログ修正 -add end

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        ''' <summary>
        ''' 沿線マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_Ensen(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'マスタからコードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT ensen_no FROM m_ensen "
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

            '20160915 沿線マスタログ修正 -add sta
            If value = "" Then
                Return rtn
            End If
            '20160915 沿線マスタログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

            If tmp_list.Contains(value) Then
                chkvalue = value
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 駅マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_Eki(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'マスタからコードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT eki_no FROM m_ensen_eki "
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

            '20160915 沿線マスタログ修正 -add sta
            If value = "" Then
                Return rtn
            End If
            '20160915 沿線マスタログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

            If tmp_list.Contains(value) Then
                chkvalue = value
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        '金融機関

        ''' <summary>
        ''' 口座振替マスタ有無チェック '2016.04.26 メインの方へも反映させる修正 -chg sta
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_KozaFkae(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'マスタからコードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT fkae_no FROM m_fb_fkaejyoho "
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

            '20160530 ログ修正 -add sta
            If value = "" Or value = "0" Then
                Return rtn
            End If
            '20160530 ログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

            If tmp_list.Contains(value) Then
                chkvalue = value
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_RELEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 仲介業者マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_GyCyukai(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'コードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT gy_fudono FROM gydata_fudo "
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 施工業者マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_GySeko(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'コードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT gy_sekono FROM gydata_seko "
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 保守業者マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_GyHosyu(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'コードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT gy_sisetuno FROM gydata_sisetu "
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 家賃保証業者マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_GyHosyo(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'コードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT gy_hosyono FROM gydata_hosyo "
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        ''' <summary>
        ''' ライフライン業者マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_GyLifeline(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'ライフライン業者の種別から該当するレコードを指定する
            Dim tmp_where As String = ""
            Dim lifelinekbn As String = erritem.Replace("ライフライン", "")

            Select Case lifelinekbn
                Case "(電気)"
                    tmp_where = " WHERE lifeline_denkiflg = 1 "
                Case "(上水)"
                    tmp_where = " WHERE lifeline_josuidoflg = 1 "
                Case "(ガス)"
                    tmp_where = " WHERE lifeline_gasflg = 1 "
                Case "(排水)"
                    tmp_where = " WHERE lifeline_haisuiflg = 1 "
                Case "(灯油)"
                    tmp_where = " WHERE lifeline_toyuflg = 1 "
                Case "(その他1)", "(その他2)", "(その他3)"
                    tmp_where = " WHERE lifeline_other1flg = 1 "
            End Select

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT gy_lifelineno FROM gydata_lifeline " & tmp_where
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

            '20160530 ログ修正 -add sta
            If value = "" Or value = "0" Then
                Return rtn
            End If
            '20160530 ログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

            If tmp_list.Contains(value) Then
                chkvalue = value
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 自社マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_Jisya(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'コードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT jisya_no FROM jisyadata "
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        '20160530 ログ修正 -del sta
        ' ''' <summary>
        ' ''' 家賃入金口座マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_YatinKoza(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    'コードを取得
        '    Dim tmp_list As New List(Of String)
        '    Dim tmp_sql As String = " SELECT yatin_kozano FROM m_yatinkoza "
        '    Dim flg As Boolean = True
        '    Dim rtn As Boolean = True

        '    flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        '    If tmp_list.Contains(value) Then
        '        chkvalue = value
        '    Else
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = ""
        '        rtn = False
        '    End If

        '    Return rtn

        'End Function
        '20160530 ログ修正 -del end

        ''' <summary>
        ''' エリアマスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_Area(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT area_no FROM m_area WHERE area_name = '" & value & "'"
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

            '20160530 ログ修正 -add sta
            If value = "" Then
                Return rtn
            End If
            '20160530 ログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

            If tmp_list.Count <> 0 Then
                chkvalue = tmp_list.Item(0)
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        '2016.04.06 都市計画・用途地域を紐データから取得するように修正 -del sta
        ' ''' <summary>
        ' '''都市開発/用途地域マスタ有無チェック
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataMstExist_TosiYoto(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        '    Dim rtn As Boolean = True

        '    If value = "99" Then
        '        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        '        chkvalue = "-1"
        '        rtn = False
        '    Else
        '        chkvalue = value
        '    End If

        '    Return rtn

        'End Function
        '2016.04.06 都市計画・用途地域を紐データから取得するように修正 -del end

        ''' <summary>
        ''' 紐付け項目有無チェック (都市計画・用途地域用) 2016.04.06 都市計画・用途地域を紐データから取得するように修正
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_RelItem_Tosiyoto(ByVal hash_rel As Hashtable, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            If value = "" Then
                Return rtn
            End If

            Select Case value
                Case "都市計画の重複"  '用途地域に入れておいた文字列
                    errstr = LOG_NAIYO_ERR_OVERLAP & "-" & "2つ目の都市計画データは用途地域へ移行できません" & "-" & LOG_TAISYO_DEFAULT
                    chkvalue = ""
                    rtn = False
                Case "用途地域の重複"  '都市計画に入れておいた文字列
                    errstr = LOG_NAIYO_ERR_OVERLAP & "-" & "2つ目の用途地域データは都市計画へ移行できません" & "-" & LOG_TAISYO_DEFAULT
                    chkvalue = ""
                    rtn = False
                Case Else
                    If hash_rel.Contains(value) Then
                        chkvalue = hash_rel.Item(value)
                    Else
                        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                        chkvalue = ""
                        rtn = False
                    End If
            End Select

            Return rtn

        End Function

        ''' <summary>
        ''' 紐付け項目有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応
        ''' 　引数にデフォルト値「defvalue」を追加
        ''' </remarks>
        Public Shared Function Chk_DataMstExist_RelItem(ByVal hash_rel As Hashtable, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String, _
                                                        Optional ByVal defvalue As String = "") As Boolean

            Dim rtn As Boolean = True

            '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -add sta
            If erritem = "物件分類" OrElse erritem = "部屋分類" Then
                If value = "" Then
                    errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                    chkvalue = defvalue
                    rtn = False
                    Return rtn
                End If
            End If
            '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -add end

            '20160530 ログ修正 -add sta
            If value = "" Or value = "0" Then
                Return rtn
            End If
            '20160530 ログ修正 -add end

            'ダミーレコードとして挿入する必要がある場合は「挿入用」として抽出しておいたデータをそのまま移行する
            '(送金ルール入金項目のその他請求等)
            '2016.04.26 メインの方へも反映させる修正 -chg sta
            'If value = "挿入用" Then
            '    chkvalue = "0"
            '    Return rtn
            'End If
            If value <> value.Replace("挿入用", "") Then
                chkvalue = "0"
                Return rtn
            End If
            '2016.04.26 メインの方へも反映させる修正 -chg end

            If hash_rel.Contains(value) Then
                chkvalue = hash_rel.Item(value)
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 紐付け項目(物件部屋鍵)有無チェック '20160530 ログ修正 -add sta
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_RelItem_kagi(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            If value = "" Then
                Return rtn
            End If

            '20160927 汎用CV時の鍵情報の取得処理修正 -chg sta
            ''20160613 鍵情報の取得処理修正 -chg sta
            ''Select Case erritem
            ''    Case "共用鍵タイトル"
            ''        If Hash_KagiTitleKyoyo.Contains(value) Then
            ''            chkvalue = Hash_KagiTitleKyoyo.Item(value)
            ''        ElseIf Hash_KagiTitleSenyo.Contains(value) Then
            ''            '20160603 ユーザーデータ検証による修正 -chg sta
            ''            'errstr = LOG_NAIYO_ERR_SENYOKAGI & "-" & LOG_HUBI_SENYOKAGI & "-" & LOG_TAISYO_SENYOKAGI
            ''            errstr = ""
            ''            '20160603 ユーザーデータ検証による修正 -chg end
            ''        Else    '対象データが共用鍵かつ専用鍵タイトルにデータが含まれていない場合はエラーログ出力
            ''            errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            ''            chkvalue = ""
            ''            rtn = False
            ''        End If
            ''    Case "専用鍵タイトル"
            ''        If Hash_KagiTitleSenyo.Contains(value) Then
            ''            chkvalue = Hash_KagiTitleSenyo.Item(value)
            ''        ElseIf Hash_KagiTitleKyoyo.Contains(value) Then
            ''            '20160603 ユーザーデータ検証による修正 -chg sta
            ''            'errstr = LOG_NAIYO_ERR_KYOYOKAGI & "-" & LOG_HUBI_KYOYOKAGI & "-" & LOG_TAISYO_KYOYOKAGI
            ''            errstr = ""
            ''            '20160603 ユーザーデータ検証による修正 -chg end
            ''        Else    '対象データが専用鍵かつ共用鍵タイトルにデータが含まれていない場合はエラーログ出力
            ''            errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            ''            chkvalue = ""
            ''            rtn = False
            ''        End If
            ''End Select

            'If Hash_KagiTitle_V7to10.Contains(value) Then

            '    Dim tmp_value As String = Hash_KagiTitle_V7to10.Item(value)
            '    Dim tmp_str() As String = tmp_value.Split("-")
            '    Dim tmp_kbn As String = tmp_str(0)
            '    Dim tmp_kagino As String = tmp_str(1)

            '    Select Case erritem

            '        Case "共用鍵タイトル"
            '            If tmp_kbn = "1" Then
            '                chkvalue = tmp_kagino
            '                Return rtn
            '            End If
            '        Case "専用鍵タイトル"
            '            If tmp_kbn = "2" Then
            '                chkvalue = tmp_kagino
            '                Return rtn
            '            End If
            '    End Select

            'End If
            ''20160613 鍵情報の取得処理修正 -chg end
            Select Case CNVNO

                Case ConvertTypes._汎用
                    '20161028 物件/部屋鍵取得方法修正 -del sta
                    '汎用CVの場合はあらかじめデフォルト「鍵1」～「鍵50」を設定しているため鍵タイトルの有無の確認は行わないようにする
                    'Dim list_kagino As New List(Of String)

                    'Select Case erritem
                    '    Case "共用鍵タイトル"
                    '        Dim list_bkkagino As New List(Of String) From {"1", "2", "3"}
                    '        list_kagino = list_bkkagino
                    '    Case "専用鍵タイトル"
                    '        Dim list_hykagino As New List(Of String) From {"1", "2", "3", "4", "5", "6"}
                    '        list_kagino = list_hykagino
                    'End Select

                    ''マスタと照合
                    'If list_kagino.Contains(value) Then
                    '    chkvalue = value
                    'Else
                    '    errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                    '    chkvalue = ""
                    '    rtn = False
                    'End If
                    '20161028 物件/部屋鍵取得方法修正 -del end
                Case ConvertTypes._既存ユーザ用
                    If Hash_KagiTitle_V7to10.Contains(value) Then

                        Dim tmp_value As String = Hash_KagiTitle_V7to10.Item(value)
                        Dim tmp_str() As String = tmp_value.Split("-")
                        Dim tmp_kbn As String = tmp_str(0)
                        Dim tmp_kagino As String = tmp_str(1)

                        Select Case erritem

                            Case "共用鍵タイトル"
                                If tmp_kbn = "1" Then
                                    chkvalue = tmp_kagino
                                    Return rtn
                                End If
                            Case "専用鍵タイトル"
                                If tmp_kbn = "2" Then
                                    chkvalue = tmp_kagino
                                    Return rtn
                                End If
                        End Select

                    End If
            End Select
            '20160927 汎用CV時の鍵情報の取得処理修正 -chg end

            Return rtn

        End Function

        ''' <summary>
        ''' 紐付け項目(設備)有無チェック '20160530 ログ修正
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_RelItem_setubi(ByVal hash_rel As Hashtable, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            If value = "" Or value = "0" Then
                Return rtn
            End If

            '20160913_2 部屋設備移行処理の追加 -chg sta
            '設備で紐付けられていないデータはOKとしログを出力しない
            'If hash_rel.Contains(value) Then
            '    chkvalue = hash_rel.Item(value)
            'Else
            '    'errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            '    'chkvalue = ""
            '    'rtn = False
            'End If
            Select Case CNVNO
                Case ConvertTypes._既存ユーザ用
                    If hash_rel.Contains(value) Then
                        chkvalue = hash_rel.Item(value)
                    Else
                    End If
                Case ConvertTypes._汎用
                    chkvalue = value
            End Select
            '20160913_2 部屋設備移行処理の追加 -chg end

            Return rtn

        End Function

        ''' <summary>
        ''' 間取文字列有無チェック
        ''' </summary>
        ''' <param name="chk_m"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_MadoriStr(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            If value <> "" Then

                Select Case value
                    Case "R"
                        chkvalue = 1
                    Case "K"
                        chkvalue = 2
                    Case "DK"
                        chkvalue = 3
                    Case "LDK"
                        chkvalue = 4
                    Case "SK"
                        chkvalue = 5
                    Case "SDK"
                        chkvalue = 6
                    Case "SLDK"
                        chkvalue = 7
                    Case "LK"
                        chkvalue = 8
                    Case "SLK"
                        chkvalue = 9
                    Case "SR"
                        chkvalue = 10
                    Case Else
                        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                        chkvalue = -1
                        rtn = False
                End Select

            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 間取文字列有無チェック(内訳)
        ''' </summary>
        ''' <param name="chk_m"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_MadoriUtiwakeStr(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            '20160915 間取内訳情報の汎用/既存処理の分岐 -chg sta
            'If value <> "" Then

            '    Select Case value
            '        Case "和"
            '            chkvalue = 1
            '        Case "洋"
            '            chkvalue = 2
            '        Case "K"
            '            chkvalue = 8
            '        Case "L"
            '            chkvalue = 9
            '        Case "DK"
            '            chkvalue = 10
            '        Case "LD"
            '            chkvalue = 11
            '        Case "LDK"
            '            chkvalue = 12
            '        Case "ロフト"
            '            chkvalue = 3
            '        Case "S"
            '            chkvalue = 4
            '        Case "書斎"
            '            chkvalue = 5
            '        Case "サンルーム"
            '            chkvalue = 6
            '        Case "グルニエ"
            '            chkvalue = 7
            '        Case Else
            '            '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg sta
            '            'errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            '            If value.Replace("(対応不要)", "") <> value Then
            '                errstr = LOG_NAIYO_ERR_TAIONASI & "-" & LOG_HUBI_TAIONASI & "-" & LOG_TAISYO_TAIONASI
            '            Else
            '                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            '            End If
            '            '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg end
            '            chkvalue = ""
            '            rtn = False
            '    End Select

            'End If

            If value = "" Then
                Return rtn
            End If

            Select Case CNVNO
                Case ConvertTypes._既存ユーザ用
                    Select Case value
                        Case "和"
                            chkvalue = 1
                        Case "洋"
                            chkvalue = 2
                        Case "K"
                            chkvalue = 8
                        Case "L"
                            chkvalue = 9
                        Case "DK"
                            chkvalue = 10
                        Case "LD"
                            chkvalue = 11
                        Case "LDK"
                            chkvalue = 12
                        Case "ロフト"
                            chkvalue = 3
                        Case "S"
                            chkvalue = 4
                        Case "書斎"
                            chkvalue = 5
                        Case "サンルーム"
                            chkvalue = 6
                        Case "グルニエ"
                            chkvalue = 7
                        Case Else
                            '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg sta
                            'errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                            If value.Replace("(対応不要)", "") <> value Then
                                errstr = LOG_NAIYO_ERR_TAIONASI & "-" & LOG_HUBI_TAIONASI & "-" & LOG_TAISYO_TAIONASI
                            Else
                                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                            End If
                            '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg end
                            chkvalue = ""
                            rtn = False
                    End Select
                Case ConvertTypes._汎用
                    chkvalue = value
            End Select
            '20160915 間取内訳情報の汎用/既存処理の分岐 -chg end
            Return rtn

        End Function

        ''' <summary>
        ''' 契約分類マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_Kybrui(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT ky_ruino FROM m_ky_rui "
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

            If tmp_list.Contains(value) Then
                chkvalue = value
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 参照ファイル有無チェック
        ''' </summary>
        ''' <param name="chk_m"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataExist_FilePath(ByVal value As String, ByRef chkvalue As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            'ファイル有無確認
            If File.Exists(value) Then
                chkvalue = value
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTFILEDATA & "-" & LOG_HUBI_NOTEXISTDATA_FILEEXIST & "-" & LOG_TAISYO_DEFAULT
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 請求発生詳細情報を設定する '2016.04.26 メインの方へも反映させる修正
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="chkkbn"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_Data_SqHasseiItem(ByVal value As String, ByRef chkvalue As String, ByVal chkkbn As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            '連結しているデータを分割
            Dim tmp_str() As String = value.Split("-")

            '各変数へ格納
            Dim tmp_cvitem As String = tmp_str(0)       '移行対象項目
            Dim tmp_sqtuki As String = tmp_str(1)       '請求月 (前月、当月等)
            Dim tmp_sqkaisiymd As String = tmp_str(2)   '請求開始年月
            Dim tmp_sqkankaku As String = tmp_str(3)    '請求間隔

            '返却用変数
            Dim sqptn As String = "1"       'デフォルト値を設定しておく
            Dim sqinterval As String = "1"  'デフォルト値を設定しておく
            Dim sqnenyear As String = ""
            Dim sqtukiyear As String = ""

            Select Case tmp_cvitem
                Case "部屋入金項目情報"
                    '-------------------------
                    '請求間隔の判別
                    '-------------------------
                    If tmp_sqkankaku = "" OrElse tmp_sqkankaku = "0" OrElse tmp_sqkankaku = "1" Then
                        Return rtn
                    Else
                        rtn = False
                        Return rtn
                    End If

                Case "契約入金項目情報", "家主固定控除情報"
                    '-------------------------
                    '各設定値のチェック (設定値が全て揃っていないと移行できないため)
                    '-------------------------
                    If tmp_sqtuki = "" OrElse tmp_sqkaisiymd = "" OrElse tmp_sqkankaku = "" Then
                        Return rtn
                    End If

                    '-------------------------
                    '請求月の取得
                    '-------------------------
                    Dim tmp_chgsqtuki As Integer = 0    '請求月
                    Dim tmp_addtuki As Integer = 0
                    If tmp_sqtuki = "" OrElse Int32.TryParse(tmp_sqtuki, tmp_chgsqtuki) = False Then
                        rtn = False
                        Return rtn
                    Else
                        '変換
                        Select Case tmp_chgsqtuki
                            Case 1 : tmp_addtuki = -2
                            Case 2 : tmp_addtuki = -1
                            Case 3 : tmp_addtuki = 0
                            Case 4 : tmp_addtuki = 1
                            Case 5 : tmp_addtuki = 2
                        End Select
                    End If

                    '-------------------------
                    '請求開始年月の月部分を取得
                    '-------------------------
                    Dim tmp_chgsqkaisiymd As DateTime
                    Dim tmp_sqkaisituki As Integer = 0  '請求開始月
                    If tmp_sqkaisiymd = "" OrElse DateTime.TryParse(tmp_sqkaisiymd, tmp_chgsqkaisiymd) = False Then
                        rtn = False
                        Return rtn
                    Else
                        tmp_sqkaisituki = Int32.Parse(tmp_chgsqkaisiymd.Month)
                    End If

                    '-------------------------
                    '請求開始月に対する請求月の取得 (請求間隔が2ヶ月の場合の奇数偶数請求月設定用)
                    '-------------------------
                    Dim tmp_sqtaisyotuki As Integer = tmp_sqkaisituki + tmp_addtuki     '請求月を取得 (該当月に対する請求月)
                    If tmp_sqtaisyotuki <= 0 Then
                        tmp_sqtaisyotuki = tmp_sqtaisyotuki + 12    '負の請求月の調整 (前月請求で1月の場合は12月に調整)
                    End If

                    '-------------------------
                    '請求間隔判別
                    '-------------------------
                    Dim tmp_chgsqkankaku As Integer = 0

                    '空または数値以外の場合は処理を抜ける
                    If tmp_sqkankaku = "" OrElse Int32.TryParse(tmp_sqkankaku, tmp_chgsqkankaku) = False Then
                        rtn = False
                        Return rtn
                    End If

                    '請求間隔が13ヶ月以上または12を割った場合に余りが生じる場合は処理を抜ける
                    If tmp_chgsqkankaku >= 13 OrElse 12 Mod tmp_chgsqkankaku <> 0 Then
                        rtn = False
                        Return rtn
                    End If

                    '請求間隔による処理の分岐
                    If tmp_chgsqkankaku = 2 Then
                        '2ヶ月の場合は奇数月か偶数月か判別して処理を抜ける
                        If tmp_sqtaisyotuki Mod 2 <> 0 Then
                            '奇数月請求
                            sqinterval = "2"
                        ElseIf tmp_sqtaisyotuki Mod 2 = 0 Then
                            '偶数月請求
                            sqinterval = "3"
                        End If
                    ElseIf tmp_chgsqkankaku >= 3 Then
                        '3ヶ月以上の場合は詳細な設定にチェック
                        sqptn = "2"

                    End If

                    '年あたりの請求回数を取得
                    Dim tmp_sqcnt As Integer = 12 / tmp_chgsqkankaku
                    Dim tmp_listsqtukigrp As New List(Of Integer)
                    For cntii = 1 To tmp_sqcnt
                        If tmp_chgsqkankaku = 1 Then
                            tmp_listsqtukigrp.Add(cntii)
                        Else
                            tmp_listsqtukigrp.Add(((tmp_chgsqkankaku * (cntii - 1)) + tmp_sqkaisituki) Mod 12)
                        End If
                    Next

                    Call Set_Sqnen(tmp_listsqtukigrp, tmp_sqkaisituki, tmp_addtuki, sqnenyear)
                    Call Set_Sqtuki(tmp_listsqtukigrp, tmp_sqkaisituki, tmp_addtuki, sqtukiyear)

            End Select

            'チェック項目で返す値を設定
            Select Case chkkbn
                Case "請求パターン"
                    chkvalue = sqptn
                Case "請求間隔"
                    chkvalue = sqinterval
                Case "請求発生年"
                    chkvalue = sqnenyear
                Case "請求発生月"
                    chkvalue = sqtukiyear
            End Select

            Return rtn

        End Function

        ''' <summary>
        ''' 請求発生年の設定
        ''' </summary>
        ''' <param name="sqtukigrp"></param>
        ''' <param name="sqkaisituki"></param>
        ''' <param name="addtuki"></param>
        ''' <param name="sqnenyear"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_Sqnen(ByVal sqtukigrp As Object, ByVal sqkaisituki As Integer, ByVal addtuki As Integer, ByRef sqnenyear As String)

            '請求年の設定
            Dim tmp_sqnenyear(12) As String
            For Each sqtuki In sqtukigrp

                For cntii = 1 To 12

                    '20160603 ユーザーデータ検証による修正 -chg sta
                    '全ての該当月に対して「当年」を設定するように修正
                    'If cntii < sqkaisituki Then
                    '    tmp_sqnenyear(cntii) = "3"
                    'Else
                    '    tmp_sqnenyear(cntii) = "2"
                    'End If
                    tmp_sqnenyear(cntii) = "2"
                    '20160603 ユーザーデータ検証による修正 -chg end

                Next

            Next

            '移行用に成形
            Dim tmp_strsqnen As String = ""
            Dim tmp_strsqtuki As String = ""

            For cntii = 1 To 12
                tmp_strsqnen = tmp_strsqnen & "," & tmp_sqnenyear(cntii)
            Next

            sqnenyear = tmp_strsqnen.Remove(0, 1)

        End Sub

        ''' <summary>
        ''' 請求発生月の設定
        ''' </summary>
        ''' <param name="sqtukigrp"></param>
        ''' <param name="sqkaisituki"></param>
        ''' <param name="addtuki"></param>
        ''' <param name="sqtukiyear"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_Sqtuki(ByVal sqtukigrp As Object, ByVal sqkaisituki As Integer, ByVal addtuki As Integer, ByRef sqtukiyear As String)

            '請求月の設定
            Dim tmp_sqnentuki(12) As String
            For Each sqtuki In sqtukigrp

                Dim tmp_sqtaisyotuki As Integer = sqtuki + addtuki     '請求月を取得 (該当月に対する請求月)
                If tmp_sqtaisyotuki <= 0 Then
                    tmp_sqtaisyotuki = tmp_sqtaisyotuki + 12    '負の請求月の調整 (前月請求で1月の場合は12月に調整)
                End If

                tmp_sqnentuki(sqtuki) = tmp_sqtaisyotuki.ToString

            Next

            For cntii = 1 To 12

                If tmp_sqnentuki(cntii) Is Nothing Then
                    tmp_sqnentuki(cntii) = "13"
                End If

            Next

            '移行用に成形
            Dim tmp_strsqnen As String = ""
            Dim tmp_strsqtuki As String = ""

            For cntii = 1 To 12
                tmp_strsqtuki = tmp_strsqtuki & "," & tmp_sqnentuki(cntii)
            Next

            sqtukiyear = tmp_strsqtuki.Remove(0, 1)

        End Sub

        ''' <summary>
        ''' 自社口座マスタ有無チェック
        ''' 引数 sqlcnnv10 を追加 '20160928 自社口座取得方法の修正
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="chkkbn"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_Data_FBKoza(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal chkkbn As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim jisyano As String = ""
            Dim jisyakozano As String = ""
            Dim erritem As String = "自社口座"

            '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
            ''20160530 ログ修正 -add sta
            'If value = "" Then
            '    Return rtn
            'End If
            ''20160530 ログ修正 -add end

            'If Hash_Rel_JisyaKoza.Contains(value) = False Then
            '    errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            '    chkvalue = ""
            '    rtn = False
            '    Return rtn
            'Else
            '    Dim tmp_jisyainfo As String = Hash_Rel_JisyaKoza.Item(value)
            '    Dim tmp_item() As String = tmp_jisyainfo.Split("-")
            '    jisyano = tmp_item(0)
            '    jisyakozano = tmp_item(1)
            'End If

            'Select Case chkkbn
            '    Case "自社_口座用"
            '        chkvalue = jisyano
            '    Case "自社口座_口座用"
            '        chkvalue = jisyakozano
            'End Select

            '空データチェック
            If value = "" Then
                Return rtn
            End If

            '汎用/既存で処理を分岐
            Select Case CNVNO

                Case ConvertTypes._既存ユーザ用

                    '既存は紐付情報から取得
                    If Hash_Rel_JisyaKoza.Contains(value) = False Then
                        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                        chkvalue = ""
                        rtn = False
                        Return rtn
                    Else
                        Dim tmp_jisyainfo As String = Hash_Rel_JisyaKoza.Item(value)
                        Dim tmp_item() As String = tmp_jisyainfo.Split("-")
                        jisyano = tmp_item(0)
                        jisyakozano = tmp_item(1)
                    End If

                    Select Case chkkbn
                        Case "自社_口座用"
                            chkvalue = jisyano
                        Case "自社口座_口座用"
                            chkvalue = jisyakozano
                    End Select

                Case ConvertTypes._汎用

                    '汎用は直接設定された値から取得
                    '20160928 自社口座取得方法の修正 -chg sta
                    'chkvalue = value

                    'データが存在しない場合は処理を抜ける
                    If value = "-" Then
                        Return rtn
                    End If

                    '自社No、自社口座No分割
                    Dim tmp_str() As String = value.Split("-")
                    Dim tmp_jisyano As String = tmp_str(0)
                    Dim tmp_jisyakozano As String = tmp_str(1)
                    Dim chkstrafter As String = ""
                    Dim tmp_sql As String = ""
                    Dim tmp_list As New List(Of String)
                    Dim taisyovalu As String = ""

                    Select Case chkkbn
                        Case "自社_口座用"
                            '自社Noが無い場合は処理を抜ける
                            If tmp_jisyano = "" Then
                                chkvalue = ""
                                Return rtn
                            ElseIf Chk_DataNumeric(tmp_jisyano, chkstrafter, "1", "999999999", "", errstr) = False Then
                                chkvalue = ""
                                rtn = False
                                Return rtn
                            End If
                            tmp_sql = " SELECT jisya_no FROM jisyadata "
                            taisyovalu = tmp_jisyano

                            Call DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

                            If tmp_list.Contains(taisyovalu) Then
                                chkvalue = taisyovalu
                            Else
                                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                                chkvalue = ""
                                rtn = False
                            End If

                        Case "自社口座_口座用"
                            '自社口座Noチェック
                            If tmp_jisyakozano = "" Then
                                chkvalue = ""
                                Return rtn
                            ElseIf Chk_DataNumeric(tmp_jisyakozano, chkstrafter, "1", "999", "", errstr) = False Then
                                chkvalue = ""
                                rtn = False
                                Return rtn
                            End If
                            tmp_sql = " SELECT CONVERT(VARCHAR,jisya_no) + '-' + CONVERT(VARCHAR,jisya_kozano) FROM jisyadata_koza "
                            taisyovalu = value

                            Call DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

                            If tmp_list.Contains(taisyovalu) Then
                                chkvalue = tmp_jisyakozano
                            Else
                                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                                chkvalue = ""
                                rtn = False
                            End If

                    End Select
                    '20160928 自社口座取得方法の修正 -chg end
            End Select

            '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
            Return rtn

        End Function

        ''' <summary>
        ''' 家賃入金口座に紐付く家主口座マスタ有無チェック 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="chkkbn"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_Data_YatinOwKoza(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal chkkbn As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            
            '空データチェック
            If value = "" Then
                Return rtn
            End If

            '汎用/既存で処理を分岐
            Select Case CNVNO

                Case ConvertTypes._既存ユーザ用
                    '既存ユーザーの場合はチェックしない(存在するデータのみ抽出しているため)
                    chkvalue = value
                    Return rtn
                Case ConvertTypes._汎用

                    'データが存在しない場合は処理を抜ける([ow_no]-[ow_kozano]の形)
                    If value = "-" Then
                        Return rtn
                    End If

                    '家主No、家主口座No分割
                    Dim tmp_str() As String = value.Split("-")
                    Dim tmp_owno As String = tmp_str(0)
                    Dim tmp_owkozano As String = tmp_str(1)
                    Dim chkstrafter As String = ""
                    Dim tmp_sql As String = ""
                    Dim tmp_list As New List(Of String)
                    Dim taisyovalu As String = ""

                    Select Case chkkbn
                        Case "家主_口座用"
                            '家主Noが無い場合は処理を抜ける
                            If tmp_owno = "" Then
                                chkvalue = ""
                                Return rtn
                            ElseIf Chk_DataNumeric(tmp_owno, chkstrafter, "1", "999999999", "", errstr) = False Then
                                chkvalue = ""
                                rtn = False
                                Return rtn
                            End If

                            tmp_sql = " SELECT ow_no FROM owdata "
                            taisyovalu = tmp_owno

                            Call DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

                            If tmp_list.Contains(taisyovalu) Then
                                chkvalue = taisyovalu
                            Else
                                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & "家主" & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                                chkvalue = ""
                                rtn = False
                            End If

                        Case "家主口座_口座用"
                            '家主口座Noチェック
                            If tmp_owkozano = "" Then
                                chkvalue = ""
                                Return rtn
                            ElseIf Chk_DataNumeric(tmp_owkozano, chkstrafter, "1", "999999999", "", errstr) = False Then
                                chkvalue = ""
                                rtn = False
                                Return rtn
                            End If
                            tmp_sql = " SELECT CONVERT(VARCHAR,ow_no) + '-' + CONVERT(VARCHAR,ow_kozano) FROM owdata_koza "
                            taisyovalu = value

                            Call DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

                            If tmp_list.Contains(taisyovalu) Then
                                chkvalue = tmp_owkozano
                            Else
                                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & "家主口座" & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                                chkvalue = ""
                                rtn = False
                            End If

                    End Select

            End Select

            Return rtn

        End Function





        ''' <summary>
        ''' 画像判別 '20160525 クレーム関連ファイルの画像判別処理実装
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_PictureFile(ByVal value As String, ByRef chkvalue As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            '20160627 クレーム関連ファイル移行修正 -chg sta
            '画像は移行せずファイルパスをそのまま移行するように修正するためコメントアウト
            'For Each pickaku In List_ImportableImageFileAttributes
            '    If Strings.StrConv(Path.GetExtension(value), VbStrConv.Narrow Or VbStrConv.Lowercase, 0) = pickaku Then
            '        errstr = LOG_NAIYO_ERR_RELFILEPIC & "-" & LOG_HUBI_RELFILEPIC & "-" & LOG_TAISYO_RELFILEPIC
            '        chkvalue = ""
            '        rtn = False
            '        Exit For
            '    Else
            '        chkvalue = value
            '    End If
            'Next
            chkvalue = value
            '20160627 クレーム関連ファイル移行修正 -chg end

            Return rtn

        End Function

        ''' <summary>
        ''' 紐付け項目(FB関連)有無チェック '20160609 紐付設定値取得に伴う修正
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_RelItem_FB(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            If value = "" Then
                Return rtn
            End If

            Dim obj As Object = Nothing

            Select Case erritem
                Case "FB関連_口座振替"
                    obj = Hash_Rel_FBInfo_KozaFurikaeFmt
                Case "FB関連_総合振込"
                    obj = Hash_Rel_FBInfo_FuriIraiFmt
                Case "FB関連_入出金フォーマット"
                    obj = Hash_Rel_FBInfo_NsSettingFmt
                Case "FB関連_入出金自社No"
                    obj = Hash_Rel_FBInfo_NsSettingJisya
            End Select

            If obj.Contains(value) Then
                chkvalue = obj.Item(value)
            Else
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = ""
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 全体データチェック→チェックメソッド呼出
        ''' </summary>
        ''' <param name="tblfldname"></param>
        ''' <param name="value"></param>
        ''' <param name="errstr"></param>
        ''' <remarks></remarks>
        Public Shared Sub Call_ChkMethodMidTotal(ByVal tblfldname As String, ByVal value As String, ByRef errstr As String)

            Dim tmp_type As String = Hash_FiledTypeJp(tblfldname)
            Dim min As New Object
            Dim max As New Object
            Dim datemin As String = ""
            Dim datemax As String = ""

            Dim chkaftervalue As String = ""
            Dim defvalue As String = ""
            Dim normalflg As Boolean = True                     'kakaka 以下の条件を表など一覧にして仕様書に記載しておく

            '2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            'Dim tmp_size As Integer = Int32.Parse(Hash_FiledSizeJp(tblfldname))         'kakaka4 0322_1300 変換時の型落ちを考慮。
            Dim tmp_size As Integer = 0
            If Int32.TryParse(Hash_FiledSizeJp(tblfldname), tmp_size) = False Then
                errstr = LOG_NAIYO_ERR_DATACHK & "-" & LOG_HUBI_ERR_DATACHK_B & "-" & LOG_TAISYO_DATACHK_B
                normalflg = False
                Exit Sub
            End If
            '2016.03.28 型落ちを考慮した処理へ修正 -chg end

            Select Case tmp_type                                'kakaka 要確認：資料を見てチェックする(金丸)
                Case "int", "smallint", "bigint", "tinyint"
                    Call Chg_SizeToValue(tmp_size, min, max)
                    normalflg = Chk_DataNumeric(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
                Case "bit"
                    normalflg = Chk_DataBit(value, chkaftervalue, defvalue, errstr)
                Case "char", "varchar", "text"
                    normalflg = Chk_DataString(value, chkaftervalue, tmp_size, 0, errstr)
                Case "nchar", "nvarchar", "ntext", "sysname"
                    normalflg = Chk_DataString(value, chkaftervalue, tmp_size, 1, errstr)
                Case "datetime", "smalldatetime", "date", "datetime2", "datetimeoffset"
                    normalflg = Chk_DataDate(value, chkaftervalue, datemin, datemax, defvalue, errstr)
                Case "time"                                     'kakaka 注意：未実装
                    normalflg = Chk_DataTime(value, chkaftervalue, errstr)
                Case "money"
                    Call Chg_SizeToValue(tmp_size, min, max)
                    normalflg = Chk_DataMoney(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
                Case "decimal", "numeric"
                    Call Chg_SizeToValue(tmp_size, min, max)
                    normalflg = Chk_DataDecimal(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
                Case "float", "real"
                    normalflg = Chk_DataDouble(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
                Case "varbinary"

                Case "xml", "uniqueidentifier"
                    'プログラム内で生成しているためチェック不要
            End Select

            'hash_err.Add(tblfldname, )



        End Sub

        ''' <summary>
        ''' 個別データチェック→チェックメソッド呼出
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="type"></param>
        ''' <param name="min"></param>
        ''' <param name="max"></param>
        ''' <param name="defvalue"></param>
        ''' <param name="normalflg"></param>
        ''' <param name="errstr"></param>
        ''' <remarks></remarks>
        Public Shared Sub Call_ChkMethodPerItem(ByVal value As String, ByRef chkvalue As String, ByVal type As String, ByVal min As String, ByVal max As String, ByVal defvalue As String, ByRef normalflg As Boolean, ByRef errstr As String)

            '初期化                            'kakaka メモ：型チェック(ソースレビューでは全て必要なチェックがあることを前提とするので、確認しません)
            chkvalue = ""
            errstr = ""
            normalflg = True

            Select Case type
                Case "int", "smallint", "bigint", "tinyint"
                    normalflg = Chk_DataNumeric(value, chkvalue, min, max, defvalue, errstr)
                Case "bit"
                    normalflg = Chk_DataBit(value, chkvalue, defvalue, errstr)
                Case "char", "varchar", "text"
                    normalflg = Chk_DataString(value, chkvalue, max, 0, errstr)
                Case "nchar", "nvarchar", "ntext", "sysname"
                    normalflg = Chk_DataString(value, chkvalue, max, 1, errstr)
                Case "datetime", "smalldatetime", "date", "datetime2", "datetimeoffset"
                    normalflg = Chk_DataDate(value, chkvalue, min, max, defvalue, errstr)
                Case "time"
                    normalflg = Chk_DataTime(value, chkvalue, errstr)
                Case "money"
                    normalflg = Chk_DataMoney(value, chkvalue, min, max, defvalue, errstr)
                Case "decimal", "numeric"
                    normalflg = Chk_DataDecimal(value, chkvalue, min, max, defvalue, errstr)
                Case "float", "real"
                    normalflg = Chk_DataDouble(value, chkvalue, min, max, defvalue, errstr)
                Case "varbinary"

                Case "xml", "uniqueidentifier", "任意文字"  '20160927 ログの内容が不正になっているため修正 "任意文字"を追加
                    'プログラム内で生成しているためチェックせずそのまま返却
                    chkvalue = value

            End Select

        End Sub

        ''' <summary>
        ''' 有無参照データチェックメソッド呼出
        ''' '20160829 メールアドレス、URLの正規化処理を追加 Optionalで引数を追加
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="chkkbn"></param>
        ''' <param name="normalflg"></param>
        ''' <param name="errstr"></param>
        ''' <remarks>
        ''' 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応
        ''' 　引数にデフォルト値「defvalue」を追加
        ''' </remarks>
        Public Shared Sub Call_ChkMethodPerItem_MstExist(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal chkkbn As String, ByRef normalflg As Boolean, ByRef errstr As String, _
                                                         Optional ByVal maxsize As String = "", Optional defvalue As String = "")

            '初期化
            chkvalue = ""
            errstr = ""
            normalflg = True
            '20160530 ログ修正 -add
            Dim tmp_sql As String = ""

            Select Case chkkbn
                Case "都道府県"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_Todofuken(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT ken_no FROM m_ken "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end
                Case "市区町村"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_Sikucyoson(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT si_no FROM m_si "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end
                Case "金融機関"
                    '構築中
                    chkvalue = value
                Case "金融機関支店"
                    '構築中
                    chkvalue = value
                Case "仲介業者"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_GyCyukai(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT gy_fudono FROM gydata_fudo "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end
                Case "ライフライン(電気)", "ライフライン(上水)", "ライフライン(ガス)", "ライフライン(排水)", "ライフライン(灯油)", "ライフライン(その他1)", "ライフライン(その他2)", "ライフライン(その他3)"
                    normalflg = Chk_DataMstExist_GyLifeline(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                Case "家賃入金口座"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_YatinKoza(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT yatin_kozano FROM m_yatinkoza "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end
                Case "施工業者"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_GySeko(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT gy_sekono FROM gydata_seko "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end
                Case "保守業者"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_GyHosyu(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT gy_sisetuno FROM gydata_sisetu "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end
                Case "家賃保証業者"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_GyHosyo(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT gy_hosyono FROM gydata_hosyo "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end

                    '20160530 ログ修正 -add sta
                Case "修繕業者"
                    tmp_sql = " SELECT gy_szenno FROM gydata_szen "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -add end
                Case "自社"
                    '20160530 ログ修正 -chg sta
                    'normalflg = Chk_DataMstExist_Jisya(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    tmp_sql = " SELECT jisya_no FROM jisyadata "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                    '20160530 ログ修正 -chg end
                Case "自社口座"
                    '構築中
                    chkvalue = value
                Case "家主"
                    tmp_sql = " SELECT ow_no FROM owdata "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                Case "家主口座"
                    '構築中
                    chkvalue = value
                Case "契約者"
                    '構築中
                    chkvalue = value
                    '2016.04.06 都市計画・用途地域を紐データから取得するように修正 -chg sta
                    'Case "都市計画・用途地域"
                    '    normalflg = Chk_DataMstExist_TosiYoto(value, chkvalue, chkkbn, errstr)
                Case "都市計画", "用途地域"
                    normalflg = Chk_DataMstExist_RelItem_Tosiyoto(Hash_Rel_Tosiyotono, value, chkvalue, chkkbn, errstr)
                    '2016.04.06 都市計画・用途地域を紐データから取得するように修正 -chg end
                Case "エリア"
                    normalflg = Chk_DataMstExist_Area(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                Case "沿線"
                    normalflg = Chk_DataMstExist_Ensen(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                Case "駅"
                    normalflg = Chk_DataMstExist_Eki(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                Case "間取"
                    normalflg = Chk_DataMstExist_MadoriStr(value, chkvalue, chkkbn, errstr)
                Case "間取内訳"
                    normalflg = Chk_DataMstExist_MadoriUtiwakeStr(value, chkvalue, chkkbn, errstr)
                Case "契約分類"
                    '2016.04.06 契約分類を紐付データを元に移行 -chg sta
                    'normalflg = Chk_DataMstExist_Kybrui(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Kyrui, value, chkvalue, chkkbn, errstr)
                    '2016.04.06 契約分類を紐付データを元に移行 -chg end
                Case "物件分類"
                    '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg sta
                    'normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Bkrui, value, chkvalue, chkkbn, errstr)
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Bkrui, value, chkvalue, chkkbn, errstr, defvalue)
                    '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg end
                Case "構造"
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Kozo, value, chkvalue, chkkbn, errstr)
                Case "部屋分類"
                    '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg sta
                    'normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Hyrui, value, chkvalue, chkkbn, errstr)
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Hyrui, value, chkvalue, chkkbn, errstr, defvalue)
                    '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg end
                Case "パス"
                    normalflg = Chk_DataExist_FilePath(value, chkvalue, errstr)
                Case "設備"
                    normalflg = Chk_DataMstExist_RelItem_setubi(Hash_Rel_Setubi, value, chkvalue, chkkbn, errstr)
                Case "取引態様"
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Toritaiyo, value, chkvalue, chkkbn, errstr)
                Case "口座種別"
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Kozasyubetu, value, chkvalue, chkkbn, errstr)
                    '20161004 口座種別のデフォルト値設定処理追加 -add sta
                    If chkvalue = "" Then
                        chkvalue = "1"
                    End If
                    '20161004 口座種別のデフォルト値設定処理追加 -add end
                Case "入金項目"
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Nkinkomk, value, chkvalue, chkkbn, errstr)
                Case "入金区分"
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Nkinkbn, value, chkvalue, chkkbn, errstr)
                Case "変動費マスタ"
                    tmp_sql = " SELECT rule_no FROM m_hendorule "
                    normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, chkvalue, chkkbn, errstr)
                Case "メーター分類"   '要紐付け作成(入金項目名→メーター分類)   ※仮で紐付けファイルだけ用意して検証
                    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Hendometer, value, chkvalue, chkkbn, errstr)
                    '20160613 鍵情報の取得処理修正 -del sta
                    '    '2016.04.06 物件鍵情報の移行処理修正 -add sta
                    'Case "共用鍵タイトル"
                    '    '20160530 ログ修正 -chg sta
                    '    'normalflg = Chk_DataMstExist_RelItem(Hash_KagiTitleKyoyo, value, chkvalue, chkkbn, errstr)
                    '    normalflg = Chk_DataMstExist_RelItem_kagi(value, chkvalue, chkkbn, errstr)
                    '    '20160530 ログ修正 -chg end
                    'Case "専用鍵タイトル"
                    '    '20160530 ログ修正 -chg sta
                    '    'normalflg = Chk_DataMstExist_RelItem(Hash_KagiTitleSenyo, value, chkvalue, chkkbn, errstr)
                    '    normalflg = Chk_DataMstExist_RelItem_kagi(value, chkvalue, chkkbn, errstr)
                    '    '20160530 ログ修正 -chg end
                    '    '2016.04.06 物件鍵情報の移行処理修正 -add end
                    '    '2016.04.26 メインの方へも反映させる修正 -add sta
                    '20160613 鍵情報の取得処理修正 -del end


                    '20160613 鍵情報の取得処理修正 -add sta
                Case "共用鍵タイトル", "専用鍵タイトル"
                    normalflg = Chk_DataMstExist_RelItem_kagi(value, chkvalue, chkkbn, errstr)
                    '20160613 鍵情報の取得処理修正 -add end

                Case "請求パターン", "請求間隔", "請求発生年", "請求発生月"
                    normalflg = Chk_Data_SqHasseiItem(value, chkvalue, chkkbn, errstr)
                    '2016.04.26 メインの方へも反映させる修正 -add end


                    '2016.04.26 メインの方へも反映させる修正 -add sta
                Case "自社_口座用", "自社口座_口座用"
                    normalflg = Chk_Data_FBKoza(sqlcnnv10, value, chkvalue, chkkbn, errstr)    '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 引数 sqlcnnv10 を追加
                    '2016.04.26 メインの方へも反映させる修正 -add end

                Case "家主_口座用", "家主口座_口座用"    '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                    normalflg = Chk_Data_YatinOwKoza(sqlcnnv10, value, chkvalue, chkkbn, errstr)


                    '2016.04.26 メインの方へも反映させる修正 -add sta
                Case "口座振替"
                    normalflg = Chk_DataMstExist_KozaFkae(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                    '2016.04.26 メインの方へも反映させる修正 -add end

                    '20160609 紐付設定値取得に伴う修正 -del sta
                    '    '20160517 入出金取得情報の新規作成 -add sta
                    'Case "FBフォーマット"
                    '    normalflg = Chk_DataMstExist_RelItem(Hash_Rel_FBFmt, value, chkvalue, chkkbn, errstr)
                    '    '20160517 入出金取得情報の新規作成 -add end
                    '20160609 紐付設定値取得に伴う修正 -del end

                    '20160609 紐付設定値取得に伴う修正 -add sta
                Case "FB関連_口座振替", "FB関連_総合振込", "FB関連_入出金フォーマット", "FB関連_入出金自社No"
                    normalflg = Chk_DataMstExist_RelItem_FB(value, chkvalue, chkkbn, errstr)
                    '20160609 紐付設定値取得に伴う修正 -add end

                    '20160524 クレーム情報移行処理実装　-add sta
                Case "箇所分類"
                    normalflg = Chk_DataMstExist_RelItem(Hash_ClaimBruiKasyo, value, chkvalue, chkkbn, errstr)
                Case "クレーム分類"
                    normalflg = Chk_DataMstExist_RelItem(Hash_ClaimBruiClaim, value, chkvalue, chkkbn, errstr)
                    '20160524 クレーム情報移行処理実装　-add end

                    '20160525 クレーム関連ファイルの画像判別処理実装 -add sta
                Case "画像判別"
                    normalflg = Chk_PictureFile(value, chkvalue, errstr)
                    '20160525 クレーム関連ファイルの画像判別処理実装 -add end

                    '20160530 ログ修正 -add sta
                Case "担当者"
                    '構築中
                    chkvalue = value
                    '20160530 ログ修正 -add end
                Case "口座名義カナ"     '20160829 口座名義カナチェック機能の追加 -add
                    normalflg = Chk_Kana(value, chkvalue, errstr)
                Case "金融機関_自社口座", "金融機関_家主口座"             '20160829 自社口座にデフォルト値を設定する処理を追加 -add '20161005 自社、家主口座デフォルト値設定処理追加 "金融機関_家主口座" を追加 -add
                    '20161005 自社、家主口座デフォルト値設定処理追加 -chg sta
                    'normalflg = Chk_Jisyakoza(value, chkvalue, errstr)
                    normalflg = Chk_JisyaOwkoza(value, chkvalue, chkkbn, errstr)
                    '20161005 自社、家主口座デフォルト値設定処理追加 -chg end
                Case "修繕維持箇所分類"             '20160829 箇所未登録データにデフォルト値を設定 -add
                    normalflg = ChK_KasyoBrui(value, chkvalue, errstr)
                Case "URL"  '20160829 メールアドレス、URLの正規化処理を追加
                    normalflg = Chk_URL(value, chkvalue, maxsize, errstr)

            End Select

        End Sub

        ''' <summary>
        ''' 個別データチェックメイン処理(移行処理からの第一階層)
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <param name="keyvalue"></param>
        ''' <param name="list"></param>
        ''' <param name="hash_chkbefore"></param>
        ''' <param name="hash_chkafter"></param>
        ''' <param name="conditioncnt"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataPerItem(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByVal list_chkduplicate As List(Of String), ByVal list_basekeydata As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer, ByRef hash_log As Hashtable) As Boolean
            'kakaka ↑"keyvalue"は読み込んだセル値が渡されるが、どこで使用する？
            'kakaka4 0322_1435 chg s 
            'Dim rtn As Boolean = True
            Dim rtn As Boolean = False
            'kakaka4 0322_1435 chg e
            Dim condflg As Boolean = False
            Dim tmp_mainkey As String = ""
            Dim midkeytblfldname(10) As String
            Dim midkeyvalue(10) As String

            '2016.04.26 メインの方へも反映させる修正 -add
            Dim chk_duplicateflg As Boolean = True                          '重複チェック実行フラグ

            '20160928 親有無チェック実行フラグの設定処理追加 -add sta
            Dim parentchkflg As Boolean = True                              '親有無チェック実行フラグ

            '------------------------
            '各移行値チェック
            '------------------------

            For Each hashvalue In hash_chkbefore

                Dim tmp_key As String = hashvalue.Key                           '移行項目名 (アルファベット)                                              'kakaka 説明を入れてください。
                Dim get_key As String = tblname & "-" & tmp_key                 'テーブル名-フィールド名 (アルファベット)
                Dim tmp_value As String = hashvalue.Value                       '移行値
                Dim tmp_type As String = Hash_FiledTypeAlpha.Item(get_key)      'データ型
                Dim tmp_size_min As String = Hash_Min_Code.Item(get_key)        '最小値 (画面上)
                Dim tmp_size_max As String = Hash_Max_Code.Item(get_key)        '最大値 (画面上)
                Dim tmp_defvalue As String = Hash_DefaultValue.Item(get_key)    'デフォルト値
                Dim tmp_chkvalue As String = ""                                 'データチェック後の値の格納用
                Dim errstr As String = ""                                       'エラー文字列格納用 (ログ出力)

                Dim normalflg As Boolean = True
                Dim keyflg As Boolean = False
                Dim requiredflg As Boolean = False

                '2016.04.26 メインの方へも反映させる修正 -add sta
                '重複チェック実行フラグの設定
                '一部の移行項目は重複チェックできないため項目毎に処理を分ける
                Dim tmp_tblnamejp As String = Hash_TblName_AlphaToJp.Item(tblname)
                Select Case tmp_tblnamejp
                    Case "変動費検針情報", "家主固定控除情報", "家主請求控除情報", "その他請求情報"
                        chk_duplicateflg = False
                    Case Else
                        chk_duplicateflg = True
                End Select
                '2016.04.26 メインの方へも反映させる修正 -add end

                '20160928 親有無チェック実行フラグの設定処理追加 -add sta
                Select Case tmp_tblnamejp
                    Case "物件基本情報", "クレーム基本情報", "契約者基本情報", "家主基本情報", _
                         "仲介業者基本情報", "保険業者基本情報", "家賃保証業者基本情報", "修繕業者基本情報", "ライフライン業者情報", "施工業者情報", "施設保守業者情報", _
                         "自社基本情報", "自社担当者情報", "振込依頼人情報", "口座振替情報", "入出金取得情報", "家賃入金口座情報", "ANSERエリア情報", "ANSERアクセスポイント情報", "ANSER接続情報", _
                         "送信設定基本情報", "ポータル連動部屋分類情報", _
                         "バス交通マスタ", "鍵タイトルマスタ", "特約マスタ", "クレーム分類設定内容", "契約分類マスタ", "保険種類マスタ", "学校区マスタ", "エリアマスタ", "変動費設定内容", "備考タイトルマスタ", "画像タイトルマスタ"
                        parentchkflg = False
                    Case Else
                        parentchkflg = True
                End Select

                '20160928 親有無チェック実行フラグの設定処理追加 -add end

                'キーフィールドの判別
                If Hash_MidKey_FieldAlpha.Contains(get_key) Then

                    '2016.02.22 型落ちを考慮した処理へ修正 -chg sta
                    'Dim tmp_keyno As Integer = Int32.Parse(Hash_MidKey_FieldAlpha.Item(get_key))        'kakaka 型落ちは大丈夫？
                    'midkeytblfldname(tmp_keyno) = get_key
                    'midkeyvalue(tmp_keyno) = tmp_value
                    'keyflg = True
                    Dim tmp_nostr As String = Hash_MidKey_FieldAlpha.Item(get_key)
                    Dim tmp_keyno As Integer = 0
                    If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
                        midkeytblfldname(tmp_keyno) = get_key
                        midkeyvalue(tmp_keyno) = tmp_value
                        keyflg = True
                    End If
                    '2016.02.22 型落ちを考慮した処理へ修正 -chg end

                End If

                'キーではないが登録する際に必須となる項目の判別(契約情報の契約開始日等)                 'kakaka メモ：実際の革命にて登録を行って確認した結果。型による一定のチェックとは別。
                If List_Required_Field.Contains(get_key) Then                                           'kakaka メモ：作業ファイルへ予め用意。それがセットされたもの
                    requiredflg = True
                End If

                'データチェック                                                                        'kakaka メモ：型(共通)チェック
                Call DataChk.Call_ChkMethodPerItem(tmp_value, tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, normalflg, errstr)      'kakaka4 ここのログ内容は出力されない場合がある？

                '参照マスタ有無チェック                                                                'kakaka わかり易い文言にして下さい(参照マスタチェック(設定元マスタがあるかどうか))
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'チェック項目名称の外出し (後で使用するため)
                'If Hash_Mst_ReferenceAlpha.Contains(get_key) Then
                '    Dim chkkbn As String = Hash_Mst_ReferenceAlpha.Item(get_key)
                '    Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr)                          'kakaka4 ここのログ内容は出力されない場合がある？
                'End If
                Dim chkkbn As String = ""
                If Hash_Mst_ReferenceAlpha.Contains(get_key) Then
                    chkkbn = Hash_Mst_ReferenceAlpha.Item(get_key)
                    '20160531 不正金額の移行制御修正 -chg sta
                    'Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr)                          'kakaka4 ここのログ内容は出力されない場合がある？
                    If chkkbn = "金額" Then
                        '金額の場合はチェックで問題があった場合に後で移行制御を行うためここでは個別データチェックを行わない
                    ElseIf chkkbn = "URL" Then     '20160829 メールアドレス、URLの正規化処理を追加
                        Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr, tmp_size_max)
                    Else
                        '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg sta
                        'Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr)                          'kakaka4 ここのログ内容は出力されない場合がある？
                        Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr, "", tmp_defvalue)                          'kakaka4 ここのログ内容は出力されない場合がある？
                        '20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg end
                    End If
                    '20160531 不正金額の移行制御修正 -chg end
                End If
                '2016.04.26 メインの方へも反映させる修正 -chg end

                '2016.02.22 エラー処理をまとめる -chg sta
                '↓↓↓旧srcコメントアウト↓↓↓
                ''チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける                    'kakaka メモ：上記までの処理の結果で不備の場合、処理をしない
                'If (keyflg AndAlso Not normalflg) OrElse (keyflg AndAlso tmp_chkvalue = "") Then                   'kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
                '    errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                '    hash_log.Add(get_key, errstr)                                                      'kakaka エラーの場合の処理を纏められたらまとめて下さい(ここと、この下)
                '    rtn = False
                '    Return rtn
                'End If

                ''チェックデータが必須項目かつ空の場合処理を抜ける
                'If requiredflg AndAlso tmp_chkvalue = "" Then
                '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                '    hash_log.Add(get_key, errstr)
                '    rtn = False
                '    Return rtn
                'End If
                '↑↑↑旧srcコメントアウト↑↑↑

                'kakaka4 0322_1435 chg s
                'If (keyflg AndAlso normalflg = False) OrElse (keyflg AndAlso tmp_chkvalue = "") Then            'チェックデータがキーかつエラー、またはキーかつ空の場合は移行不可とする
                '    errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                '    rtn = False
                'ElseIf requiredflg AndAlso tmp_chkvalue = "" Then                                           'チェックデータが必須項目かつ空の場合は移行不可とする
                '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                '    rtn = False
                'End If
                'If rtn = False AndAlso errstr <> "" Then
                '    hash_log.Add(get_key, errstr)
                '    Return rtn
                'End If

                'チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける                    'kakaka メモ：上記までの処理の結果で不備の場合、処理をしない
                If (keyflg AndAlso normalflg = False) OrElse (keyflg AndAlso tmp_chkvalue = "") Then                   'kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
                    '20161014 ログ出力内容修正 -chg sta
                    'キーエラーを優先してログに出力する(データ型やサイズに関することに触れいているため)
                    ''20160530 ログ修正 -del ログの詳細な内容が表示されないためコメントアウト
                    ''errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                    hash_log.Clear()
                    errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY & "(空行の場合は調整不要です)"
                    '20161014 ログ出力内容修正 -chg end
                    hash_log.Add(get_key, errstr)                                                      'kakaka エラーの場合の処理を纏められたらまとめて下さい(ここと、この下)
                    Return rtn
                End If

                'チェックデータが必須項目かつ空の場合処理を抜ける
                If requiredflg AndAlso tmp_chkvalue = "" Then
                    errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                    hash_log.Add(get_key, errstr)
                    Return rtn
                End If
                'kakaka4 0322_1435 chg e
                '2016.02.22 エラー処理をまとめる -chg end

                '2016.04.26 メインの方へも反映させる修正 -add sta
                'キー、必須項目以外の項目でデータの移行を制御する特殊な場合 (増えた場合は外出しすること)
                Select Case chkkbn
                    Case "請求パターン", "請求間隔", "請求発生年", "請求発生月"
                        If normalflg = False Then
                            errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH_NOTDEFAULT & "-" & LOG_TAISYO_DEFAULT
                            hash_log.Add(get_key, errstr)
                            Return rtn
                        End If
                        '20160525 クレーム関連ファイルの画像判別処理実装 -add sta
                    Case "画像判別"
                        If normalflg = False Then
                            hash_log.Add(get_key, errstr)
                            Return rtn
                        End If
                        '20160525 クレーム関連ファイルの画像判別処理実装 -add end
                        '20160531 不正金額の移行制御修正 -add sta
                    Case "金額"
                        If normalflg = False Then
                            hash_log.Add(get_key, errstr)
                            Return rtn
                        End If
                        '20160531 不正金額の移行制御修正 -add end
                End Select
                '2016.04.26 メインの方へも反映させる修正 -add end

                '調整された値が存在する場合の処理
                If normalflg = False Then                                                                'kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
                    '調整フラグをONにする
                    condflg = True
                    'エラー内容をオブジェクトへ格納
                    hash_log.Add(get_key, errstr)
                End If

                'チェックした値をチェック済み格納用ハッシュテーブルへ格納
                hash_chkafter.Add(tmp_key, tmp_chkvalue)

            Next

            '------------------------
            '重複チェック
            '------------------------

            Dim tmp_tblfldvalue As String = ""
            Dim tmp_keyvalue As String = ""

            '20160530 ログ修正 -add sta
            'テーブル名(日本語)を取得しておく
            Dim tblnamejp As String = Hash_TblName_AlphaToJp.Item(tblname)
            '20160530 ログ修正 -add end

            '退避しておいたキーフィールドと値を取得
            For cntii = 1 To UBound(midkeyvalue)
                If midkeyvalue(cntii) <> "" Then
                    tmp_tblfldvalue = tmp_tblfldvalue & "/" & midkeytblfldname(cntii)
                    tmp_keyvalue = tmp_keyvalue & "-" & midkeyvalue(cntii)
                End If
            Next
            tmp_tblfldvalue = tmp_tblfldvalue.Remove(0, 1)                                          'kakaka メモ：先頭文字除去
            tmp_keyvalue = tmp_keyvalue.Remove(0, 1)

            '照合
            '2016.04.26 メインの方へも反映させる修正 -chg sta
            '重複チェック実行有無の条件追加
            'If list_chkduplicate.Contains(tmp_keyvalue) Then                                        'kakaka メモ：先で貯めこんでいた重複データを比較
            '    hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
            '    hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
            '    'kakaka4 0322_1435  del
            '    'rtn = False
            '    Return rtn
            'End If
            If chk_duplicateflg Then
                If list_chkduplicate.Contains(tmp_keyvalue) Then                                        'kakaka メモ：先で貯めこんでいた重複データを比較
                    '20160530 ログ修正 -chg sta
                    '重複エラーログを出力する項目を条件に追加
                    'hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
                    'hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
                    If tblnamejp = "学校区マスタ" Then
                        '重複ログを出力しない
                    Else
                        hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
                        hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
                    End If
                    '20160530 ログ修正 -chg end
                    Return rtn
                End If
            End If
            '2016.04.26 メインの方へも反映させる修正 -chg end

            '------------------------
            '親マスタ有無チェック
            '------------------------
            '2016.04.26 メインの方へも反映させる修正 -chg sta
            ''照合
            'If list_basekeydata.Count <> 0 Then
            '    If Not list_basekeydata.Contains(midkeyvalue(1)) Then
            '        hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
            '        hash_log.Add(midkeytblfldname(1), LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE)
            '        'kakaka4 del
            '        'rtn = False
            '        Return rtn
            '    End If
            'End If

            '20161208 部屋情報照合の際の全角半角統一処理の削除漏れ修正 -chg sta
            ''照合用に文字を編集する処理を追加 (全角半角対応)
            ''20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -chg sta
            ''Dim tmp_chgkeyvalu As String = StrConv(midkeyvalue(1), VbStrConv.Narrow)
            'Dim tmp_chgkeyvalu As String = ""
            'If tblnamejp = "備考入力補助リストマスタ" Then
            '    tmp_chgkeyvalu = StrConv(midkeyvalue(1) & "-" & midkeyvalue(2), VbStrConv.Narrow)
            'Else
            '    tmp_chgkeyvalu = StrConv(midkeyvalue(1), VbStrConv.Narrow)
            'End If
            ''20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -chg end
            'If parentchkflg Then     '20160928 親有無チェック実行フラグの設定処理追加 条件変更 list_basekeydata.Count <> 0 → parentchkflg -chg
            '    If Not list_basekeydata.Contains(tmp_chgkeyvalu) Then
            '        hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
            '        hash_log.Add(midkeytblfldname(1), LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE)
            '        'kakaka4 del
            '        'rtn = False
            '        Return rtn
            '    End If
            'End If
            ''2016.04.26 メインの方へも反映させる修正 -chg end

            Dim tmp_chgkeyvalu As String = ""
            If tblnamejp = "備考入力補助リストマスタ" Then
                tmp_chgkeyvalu = midkeyvalue(1) & "-" & midkeyvalue(2)
            Else
                tmp_chgkeyvalu = midkeyvalue(1)
            End If
            If parentchkflg Then
                If Not list_basekeydata.Contains(tmp_chgkeyvalu) Then
                    hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
                    hash_log.Add(midkeytblfldname(1), LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE)
                    Return rtn
                End If
            End If
            '20161208 部屋情報照合の際の全角半角統一処理の削除漏れ修正 -chg end

            '--------------------------------------------
            '調整フラグがTrueの場合、調整件数を追加する
            '--------------------------------------------

            If condflg Then
                conditioncnt = conditioncnt + 1
            End If

            'kakaka4 0322_1435 add
            rtn = True
            Return rtn

        End Function

        ''' <summary>
        ''' 数値型フィールドから取得したサイズから最大最小値を作成
        ''' </summary>
        ''' <param name="size"></param>
        ''' <param name="min"></param>
        ''' <param name="max"></param>
        ''' <remarks></remarks>
        Public Shared Sub Chg_SizeToValue(ByVal size As Integer, ByRef min As Object, ByRef max As Object)

            min = 2 ^ (size * 8 - 1) * (-1)
            max = 2 ^ (size * 8 - 1) - 1

        End Sub

        ''' <summary>
        ''' 中間ファイルデータチェックメイン処理 '20160812 汎用コンバート対応
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <param name="keyvalue"></param>
        ''' <param name="list"></param>
        ''' <param name="hash_chkbefore"></param>
        ''' <param name="hash_chkafter"></param>
        ''' <param name="conditioncnt"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataPerItem_Mid(ByVal tblname As String, ByRef hash_miditem As Hashtable, ByVal list_chkduplicate As List(Of String), _
                                                   ByRef midkeytblfldname As Object, ByRef midkeyvalue As Object, ByRef hash_log As Hashtable) As Boolean

            Dim rtn As Boolean = False
            Dim condflg As Boolean = False
            Dim tmp_mainkey As String = ""

            Dim chk_duplicateflg As Boolean = True                          '重複チェック実行フラグ(現時点で未使用)

            '------------------------
            '各移行値チェック
            '------------------------

            For Each hashvalue In hash_miditem

                Dim tmp_key As String = hashvalue.Key                           '移行項目名 (アルファベット)         
                '20160913_2 「@#@」を共通変数に変更 -chg sta
                'Dim get_key As String = tblname & "@#@" & tmp_key                 'テーブル名-フィールド名 (アルファベット)
                Dim get_key As String = tblname & STR_SPLIT_1 & tmp_key                 'テーブル名-フィールド名 (アルファベット)
                '20160913_2 「@#@」を共通変数に変更 -chg end
                Dim tmp_value As String = hashvalue.Value                       '移行値
                Dim tmp_type As String = Hash_ExistMiddatatype.Item(get_key)    'データ型
                Dim tmp_size_min As String = Hash_ExistMiddatamin.Item(get_key) '最小値 (画面上)
                Dim tmp_size_max As String = Hash_ExistMiddatamax.Item(get_key) '最大値 (画面上)
                Dim tmp_defvalue As String = Hash_ExistMiddatadef.Item(get_key) 'デフォルト値
                Dim tmp_chkvalue As String = ""                                 'データチェック後の値の格納用
                Dim errstr As String = ""                                       'エラー文字列格納用 (ログ出力)

                Dim normalflg As Boolean = True
                Dim keyflg As Boolean = False
                Dim requiredflg As Boolean = False

                '重複チェック実行フラグの設定(既存用の処理を流用)
                '一部の移行項目は重複チェックできないため項目毎に処理を分ける
                '20160927 ログの内容が不正になっているため修正 -del
                'Dim tmp_tblnamejp As String = Hash_TblName_AlphaToJp.Item(tblname)
                Select Case tblname   '20160927 ログの内容が不正になっているため修正 tmp_tblnamejp → tblname -chg
                    Case "変動費検針情報", "家主固定控除情報", "家主請求控除情報", "その他請求情報"
                        chk_duplicateflg = False
                    Case Else
                        chk_duplicateflg = True
                End Select

                'キーフィールドの判別
                If Hash_ExistMiddatakey.Contains(get_key) Then

                    Dim tmp_nostr As String = Hash_ExistMiddatakey.Item(get_key)
                    Dim tmp_keyno As Integer = 0
                    If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
                        midkeytblfldname(tmp_keyno) = get_key
                        midkeyvalue(tmp_keyno) = tmp_value
                        keyflg = True
                    End If

                End If

                'キーではないが登録する際に必須となる項目の判別(契約情報の契約開始日等)    
                If List_ExistMiddatareq.Contains(get_key) Then
                    requiredflg = True
                End If
                '20160927 ログの内容が不正になっているため修正 -chg sta
                ''参照マスタ有無チェックは移行データが必要のためここでは行わない                                                    
                'Dim chkkbn As String = ""
                'If Hash_ExistMiddataref.Contains(get_key) Then
                '    chkkbn = Hash_Mst_ReferenceAlpha.Item(get_key)
                'Else
                '    'データチェック
                '    Call DataChk.Call_ChkMethodPerItem(tmp_value, tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, normalflg, errstr)
                'End If
                Dim chkkbn As String = ""
                If Hash_ExistMiddataref.Contains(get_key) Then
                    chkkbn = Hash_ExistMiddataref.Item(get_key)
                End If
                'データチェック
                '20160927 ログの内容が不正になっているため修正 -add sta
                If (tblname = "送金ルール基本情報" Or tblname = "送金ルール送金先情報" Or tblname = "送金ルール入金項目情報" Or tblname = "送金ルール控除項目情報") And tmp_key = "物件No" Then
                    Dim tmp_str() As String = tmp_value.Split("-")
                    tmp_value = tmp_str(0)
                End If
                '20160927 ログの内容が不正になっているため修正 -add end
                Call DataChk.Call_ChkMethodPerItem(tmp_value, tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, normalflg, errstr)
                '20160927 ログの内容が不正になっているため修正 -chg end
                'チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける      
                If (keyflg AndAlso normalflg = False) OrElse (keyflg AndAlso tmp_chkvalue = "") Then
                    '20161028 ログ出力内容の修正 -chg sta
                    'キーがエラーになっている場合はその内容を優先する
                    'If errstr = "" Then
                    '    errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                    'End If
                    errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY & "(空行の場合は調整不要です)"
                    '20161028 ログ出力内容の修正 -chg end
                    '20160927 ログの内容が不正になっているため修正 キーが既にエラーになっている場合はその内容を優先する -add
                    hash_log.Clear()
                    hash_log.Add(get_key, errstr)
                    Return rtn
                End If

                'チェックデータが必須項目かつ空の場合処理を抜ける
                If requiredflg AndAlso tmp_chkvalue = "" Then
                    errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                    '20160927 ログの内容が不正になっているため修正 必須項目エラーの内容を優先する -add
                    hash_log.Clear()
                    hash_log.Add(get_key, errstr)
                    Return rtn
                End If

                'キー、必須項目以外の項目でデータの移行を制御する特殊な場合 (既存用の処理を流用)
                Select Case chkkbn
                    Case "請求パターン", "請求間隔", "請求発生年", "請求発生月"
                        If normalflg = False Then
                            errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH_NOTDEFAULT & "-" & LOG_TAISYO_DEFAULT
                            hash_log.Add(get_key, errstr)
                            Return rtn
                        End If
                    Case "画像判別"
                        If normalflg = False Then
                            hash_log.Add(get_key, errstr)
                            Return rtn
                        End If
                    Case "金額"
                        If normalflg = False Then
                            hash_log.Add(get_key, errstr)
                            Return rtn
                        End If
                End Select

                '調整された値が存在する場合の処理
                If normalflg = False Then
                    '調整フラグをONにする
                    condflg = True
                    'エラー内容をオブジェクトへ格納
                    hash_log.Add(get_key, errstr)
                End If

            Next

            '------------------------
            '重複チェック
            '------------------------
            Dim tmp_tblfldvalue As String = ""
            Dim tmp_keyvalue As String = ""

            '退避しておいたキーフィールドと値を取得
            For cntii = 1 To UBound(midkeyvalue)
                If midkeyvalue(cntii) <> "" Then
                    '20160927 ログの内容が不正になっているため修正 -chg sta
                    'tmp_tblfldvalue = tmp_tblfldvalue & "/" & midkeytblfldname(cntii)
                    tmp_tblfldvalue = tmp_tblfldvalue & STR_SPLIT_2 & midkeytblfldname(cntii)
                    '20160927 ログの内容が不正になっているため修正 -chg end
                    tmp_keyvalue = tmp_keyvalue & "-" & midkeyvalue(cntii)
                End If
            Next
            tmp_tblfldvalue = tmp_tblfldvalue.Remove(0, 1)
            tmp_keyvalue = tmp_keyvalue.Remove(0, 1)

            '照合
            If chk_duplicateflg Then
                If list_chkduplicate.Contains(tmp_keyvalue) Then
                    If tblname = "学校区マスタ" Then
                        '重複ログを出力しない
                    Else
                        hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
                        hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
                    End If
                    Return rtn
                End If
                list_chkduplicate.Add(tmp_keyvalue)
            End If

            rtn = True
            Return rtn

        End Function

        ''' <summary>
        ''' 汎用中間ファイルデータチェックメイン処理 '20160812 汎用コンバート対応
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <param name="keyvalue"></param>
        ''' <param name="list"></param>
        ''' <param name="hash_chkbefore"></param>
        ''' <param name="hash_chkafter"></param>
        ''' <param name="conditioncnt"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataPerItem_BaseMid(ByVal tblname As String, ByRef hash_miditem As Hashtable, ByVal list_chkduplicate As List(Of String), _
                                                   ByRef midkeytblfldname As Object, ByRef midkeyvalue As Object, ByRef hash_log As Hashtable) As Boolean

            Dim rtn As Boolean = False
            Dim condflg As Boolean = False
            Dim tmp_mainkey As String = ""

            Dim chk_duplicateflg As Boolean = True                          '重複チェック実行フラグ(現時点で未使用)

            '------------------------
            '各移行値チェック
            '------------------------

            For Each hashvalue In hash_miditem

                Dim tmp_key As String = hashvalue.Key                           '移行項目名 (アルファベット)     
                '20160913_2 「@#@」を共通変数に変更 -chg sta
                'Dim get_key As String = tblname & "@#@" & tmp_key                 'テーブル名-フィールド名 (アルファベット)
                Dim get_key As String = tblname & STR_SPLIT_1 & tmp_key                 'テーブル名-フィールド名 (アルファベット)
                '20160913_2 「@#@」を共通変数に変更 -chg end
                Dim tmp_value As String = hashvalue.Value                       '移行値
                Dim tmp_type As String = Hash_BaseMiddatatype.Item(get_key)     'データ型
                '20161006 部屋設備チェック処理の修正 -chg sta
                'Dim tmp_size_min As String = Hash_BaseMiddatamin.Item(get_key)  '最小値 (画面上)
                Dim tmp_size_min As String = IIf(Hash_BaseMiddatamin.Item(get_key) = "未入力", "0", Hash_BaseMiddatamin.Item(get_key))  '最小値 (画面上)
                '20161006 部屋設備チェック処理の修正 -chg end
                Dim tmp_size_max As String = Hash_BaseMiddatamax.Item(get_key)  '最大値 (画面上)
                Dim tmp_defvalue As String = Hash_BaseMiddatadef.Item(get_key)  'デフォルト値
                Dim tmp_chkvalue As String = ""                                 'データチェック後の値の格納用
                Dim errstr As String = ""                                       'エラー文字列格納用 (ログ出力)

                Dim normalflg As Boolean = True
                Dim keyflg As Boolean = False
                Dim requiredflg As Boolean = False

                'キーフィールドの判別
                If Hash_BaseMiddatakey.Contains(get_key) Then

                    Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
                    Dim tmp_keyno As Integer = 0
                    If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
                        midkeytblfldname(tmp_keyno) = get_key
                        midkeyvalue(tmp_keyno) = tmp_value
                        keyflg = True
                    End If

                End If

                'キーではないが登録する際に必須となる項目の判別(契約情報の契約開始日等)                 
                If List_ExistMiddatareq.Contains(get_key) Then
                    requiredflg = True
                End If

                '参照マスタ有無チェックは移行データが必要のためここでは行わない                                                    
                Dim chkkbn As String = ""
                If Hash_BaseMiddataref.Contains(get_key) Then
                    chkkbn = Hash_Mst_ReferenceAlpha.Item(get_key)
                Else
                    'データチェック                                                                       
                    Call DataChk.Call_ChkMethodPerItem(tmp_value, tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, normalflg, errstr)
                End If

                'チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける 
                If (keyflg AndAlso normalflg = False) OrElse (keyflg AndAlso tmp_chkvalue = "") Then
                    If errstr = "" Then
                        errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY & "(空行の場合は調整不要です)"
                    End If
                    '20160927 ログの内容が不正になっているため修正 キーが既にエラーになっている場合はその内容を優先する -add
                    hash_log.Clear()
                    hash_log.Add(get_key, errstr)
                    Return rtn
                End If

                'チェックデータが必須項目かつ空の場合処理を抜ける
                If requiredflg AndAlso tmp_chkvalue = "" Then
                    errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                    '20160927 ログの内容が不正になっているため修正 必須項目エラーの内容を優先する -add
                    hash_log.Clear()
                    hash_log.Add(get_key, errstr)
                    Return rtn
                End If

                '調整された値が存在する場合の処理
                If normalflg = False Then
                    '調整フラグをONにする
                    condflg = True
                    'エラー内容をオブジェクトへ格納
                    If hash_log.Contains(get_key) = False Then
                        hash_log.Add(get_key, errstr)
                    End If
                End If

            Next

            '------------------------
            '重複チェック
            '------------------------
            Dim tmp_tblfldvalue As String = ""
            Dim tmp_keyvalue As String = ""

            '退避しておいたキーフィールドと値を取得
            For cntii = 1 To UBound(midkeyvalue)
                If midkeyvalue(cntii) <> "" Then
                    tmp_tblfldvalue = tmp_tblfldvalue & "/" & midkeytblfldname(cntii)
                    tmp_keyvalue = tmp_keyvalue & "-" & midkeyvalue(cntii)
                End If
            Next
            tmp_tblfldvalue = tmp_tblfldvalue.Remove(0, 1)
            tmp_keyvalue = tmp_keyvalue.Remove(0, 1)

            '照合
            If chk_duplicateflg Then
                If list_chkduplicate.Contains(tmp_keyvalue) Then
                    If tblname = "学校区マスタ" Then
                        '重複ログを出力しない
                    Else
                        hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
                        hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
                    End If
                    Return rtn
                End If
                list_chkduplicate.Add(tmp_keyvalue)
            End If

            rtn = True
            Return rtn

        End Function

    End Class

#End Region

#Region "配列設定クラス"

    ''' <summary>
    ''' 【配列設定クラス】
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SetArrayData

        Private myStrings() As String
        Public Sub New(ByVal capacity As Integer)
            'コンストラクタで配列を作成
            myStrings = New String(capacity) {}
        End Sub

        '既定のプロパティを宣言
        Default Public Property Item(ByVal index As Integer) As String
            Get
                Return myStrings(index)
            End Get
            Set(ByVal Value As String)
                myStrings(index) = Value
            End Set
        End Property

    End Class

#End Region

#Region "ファイル処理クラス"

    ''' <summary>
    ''' 【ファイル処理クラス】
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FileMethod

        ''' <summary>
        ''' ファイル書込み処理
        ''' </summary>
        ''' <param name="str">出力文字列</param>
        ''' <param name="filepath">ファイルパス(例：C:\Temp\test.txt)</param>
        ''' <param name="append">ファイル追記 Ture.追記(Def) False.上書き</param>
        ''' <remarks>
        ''' ・指定書込ファイルが存在しない場合、新規ファイル作成<br/>
        ''' ・ShiftJisで書込み<br/>
        ''' ・書込みファイルパスがNULLの場合はログ出力を行わない<br/>
        ''' </remarks>
        Public Shared Sub FileWriteAppend(ByVal str As String, ByVal filepath As String, Optional ByVal append As Boolean = True)

            If Not filepath Is Nothing Then
                If filepath.Trim <> "" Then
                    Dim sw As New System.IO.StreamWriter(filepath, append, System.Text.Encoding.GetEncoding("shift_jis"))
                    sw.Write(str + sw.NewLine)
                    sw.Close()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Table/ViewからCSVファイル出力
        ''' Optional はログを出力する際に用いる 2016.02.22 エラー時の対処2(エラー内容を格納する引数追加)
        ''' </summary>
        ''' <param name="sqlcnn"></param>
        ''' <param name="csvfilepath"></param>
        ''' <param name="taisyoname"></param>
        ''' <param name="sortstr"></param>
        ''' <param name="sql"></param>
        ''' <param name="logflg">True:中間ファイル出力時  False:ログファイル出力時</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TblView_Output_CSV(ByVal sqlcnn As SqlConnection, ByVal csvfilepath As String, ByVal taisyoname As String, ByVal sortstr As String, ByRef errstr As String, _
                                                  Optional ByVal sql As String = "", Optional ByVal logflg As Boolean = False) As Boolean
            'kakaka4 0322_1745 変数logflgがTrueは中間ファイル出力時、Falseはログファイル出力時の内容で処理ということを明記しておく。
            Dim rtn As Boolean = True

            '抽出クエリ作成
            Dim tmp_str As String = ""
            Dim tmp_sql As String = ""
            Dim tmp_order As String = ""

            tmp_order = " ORDER BY " & sortstr
            If sql = "" Then
                tmp_sql = "SELECT * FROM " & taisyoname & tmp_order
            Else
                tmp_sql = sql
            End If

            'ファイル出力用コマンド文字列作成
            Dim tmp_cmdstr As String = ""
            Dim tmp_constr() As String = sqlcnn.ConnectionString.Split(";")
            Dim servername As String = " -S " & tmp_constr(1).Remove(0, InStr(tmp_constr(1), "=") + 1)
            Dim catalogname As String = " -d " & tmp_constr(2).Remove(0, InStr(tmp_constr(2), "=") + 1)
            Dim username As String = " -U " & tmp_constr(3).Remove(0, InStr(tmp_constr(3), "=") + 1)
            Dim password As String = " -P " & tmp_constr(4).Remove(0, InStr(tmp_constr(4), "=") + 1)
            Dim strtype As String = " -c"
            If logflg Then
                strtype = strtype & " -t,"  'ログファイルを出力する際はカンマ区切りにする
            End If
            Dim outtype As String = " queryout "
            '20160829 CSV出力エラー修正 -chg sta
            'tmp_cmdstr = """" & tmp_sql & """" & outtype & csvfilepath & servername & catalogname & username & password & strtype
            tmp_cmdstr = """" & tmp_sql & """" & outtype & """" & csvfilepath & """" & servername & catalogname & username & password & strtype
            '20160829 CSV出力エラー修正 -chg end
            '20160516 Windows認証でログファイルが出力されない現象の対応 -add sta
            If Njc.Frm.MainFrm.optV10Authent2.Checked Then
                tmp_cmdstr = tmp_cmdstr.Replace(username & password, " -T")
            End If
            '20160516 Windows認証でログファイルが出力されない現象の対応 -add end

            'ファイル出力処理
            Try
                Dim proc As New System.Diagnostics.Process()
                proc.StartInfo.FileName = "bcp"
                proc.StartInfo.Arguments = tmp_cmdstr

                proc.StartInfo.CreateNoWindow = True
                proc.StartInfo.UseShellExecute = False
                proc.Start()

                proc.WaitForExit()
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                '2016.02.22 エラー時の対処2 -add sta
                If logflg Then
                    errstr = LOG_HUBI_MAKEMIDFILE & vbCrLf & "エラー内容：" & ex.Message
                Else
                    errstr = LOG_HUBI_MAKELOGFILE & vbCrLf & "エラー内容：" & ex.Message
                End If
                '2016.02.22 エラー時の対処2 -add end
                rtn = False
            End Try

            Return rtn

        End Function

        ''' <summary>
        ''' CSV→中間ファイル書込
        ''' </summary>
        ''' <param name="csvfilepath"></param>
        ''' <param name="midfilepath"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CSV_To_Excel(ByVal csvfilepath As String, ByVal midfilepath As String, ByVal sheetname As String, ByVal colcnt As Integer, ByRef rowcnt As Integer)

            Dim rtn As Boolean = True

            Dim exlapp As Excel.Application
            Dim exlwbook As Excel.Workbooks
            Dim exlwbookmid As Excel.Workbook
            Dim exlwsheetmid As Excel.Worksheet

            'ファイル準備
            exlapp = CreateObject("Excel.Application")
            exlapp.Visible = False

            'CSV展開ファイルブック
            exlwbook = exlapp.Workbooks

            '----------------------
            '中間ファイル設定
            '----------------------
            'オブジェクト生成
            exlwbookmid = exlapp.Workbooks.Open(midfilepath)
            exlwsheetmid = exlapp.Worksheets(sheetname)

            '書込開始行の設定
            Dim startrow As Integer = EXISTMIDFILE_READWRITE_ROW

            '既存データの確認(最大行の取得)
            Dim maxrowcnt As Integer = exlwsheetmid.UsedRange.Rows.Count

            '既存データ範囲を取得
            Dim dataarea As String = startrow & ":" & maxrowcnt

            '初期化
            If maxrowcnt >= startrow Then
                exlwsheetmid.Rows(dataarea).Delete()
            End If

            '----------------------
            'CSV展開用ファイル設定
            '----------------------
            Dim arrays As Object
            Dim arrayItems(colcnt - 1) As Object

            For cntii = 0 To colcnt - 1
                arrayItems(cntii) = New Object() {cntii + 1, 2}
            Next
            arrays = arrayItems

            ''拡張子チェック (※出力CSVファイルの拡張子は「.tmp」で出力している)
            ''拡張子がCSVだとFieldInfoが使えないので一時ファイルにコピーして開く
            'Dim filePath As String = csvfilepath
            'If (String.Equals(System.IO.Path.GetExtension(csvfilepath), ".csv", StringComparison.OrdinalIgnoreCase)) Then
            '    filePath = System.IO.Path.GetTempFileName()
            '    System.IO.File.Copy(csvfilepath, filePath, True)
            'End If

            '----------------------------
            'CSV展開→中間ファイル書込
            '----------------------------
            Try
                'CSVファイル展開
                exlwbook.OpenText(csvfilepath, FieldInfo:=arrays, Tab:=True, DataType:=1)

                '最終セルの位置を取得
                Dim maxrow As Integer = exlapp.Worksheets(1).Range(MIDFILE_READWRITE_CELLSTA).SpecialCells(Excel.XlCellType.xlCellTypeLastCell).row
                Dim maxcol As Integer = exlapp.Worksheets(1).Range(MIDFILE_READWRITE_CELLSTA).SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column

                '中間ファイルへ書込
                Dim writecellsta As String = EXISTMIDFILE_READWRITE_COL & EXISTMIDFILE_READWRITE_ROW.ToString
                exlapp.Range(exlapp.Cells(1, 1), exlapp.Cells(maxrow, maxcol)).Copy(exlwsheetmid.Range(writecellsta))

                '書込行数取得
                rowcnt = exlwsheetmid.UsedRange.Rows.Count - (EXISTMIDFILE_READWRITE_ROW - 1)

                '保存
                exlwbookmid.Save()

            Catch ex As Exception
                'エラー処理
                Debug.WriteLine(ex.Message)
                rtn = False
            Finally
                exlwbookmid.Close()
                exlwbook.Close()
                exlapp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(exlwsheetmid)
                '20161014 DisconnectedContextエラー修正 -chg sta
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(exlwbookmid)
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(exlwbook)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(CType(exlwbookmid, Object))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(CType(exlwbook, Object))
                '20161014 DisconnectedContextエラー修正 -chg end
                System.Runtime.InteropServices.Marshal.ReleaseComObject(exlapp)
                exlwsheetmid = Nothing
                exlwbookmid = Nothing
                exlwbook = Nothing
                exlapp = Nothing
            End Try

            Return rtn

        End Function

    End Class

#End Region

#Region "DB処理関連クラス"

    Public Class DBExec

        ''' <summary>
        ''' DBより全レコードを取得して返却 [DataTable]
        ''' </summary>
        ''' <param name="sqlstr">取得クエリ文</param>
        ''' <param name="sqlcnn">接続情報</param>
        ''' <param name="readtbl">読込TBL情報 [out]</param>
        ''' <returns>レコード件数<br/></returns>
        ''' <remarks>
        ''' ・sqlcnnSH→sqlcnn<br/>
        ''' ・発行クエリより抽出した全データをDataTableへ全格納
        ''' </remarks>
        Public Shared Function Exec_DataTable( _
                                         ByVal sqlstr As String, _
                                         ByRef sqlcnn As System.Data.SqlClient.SqlConnection, _
                                         ByRef readtbl As DataTable, _
                                         Optional ByRef normalflg As Boolean = True _
                                         ) As Integer

            Dim adater As New SqlDataAdapter()
            Dim time_sta As DateTime
            '★★★ CommonRepositry.Exec_DataTableを参考
            Try
                readtbl.Clear()                                                     '蓄積データクリア
                adater.SelectCommand = New SqlCommand(sqlstr, sqlcnn)
                adater.SelectCommand.CommandTimeout = LimitTimeOut                  'タイムアウト設定
                time_sta = Now                                                      '開始時間
                adater.Fill(readtbl)                                                'データ取得
                Debug.Print(Typ.ToStr(Now - TimeOfDay))                             '経過時間
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                normalflg = False
            End Try

            Return readtbl.Rows.Count

        End Function

        ''' <summary>
        ''' DBより1データを取得して返却 [ExecuteScalar]
        ''' </summary>
        ''' <param name="sqlstr">取得クエリ文</param>
        ''' <param name="sqlcnn">接続情報</param>
        ''' <returns>取得値<br/></returns>
        ''' <remarks>
        ''' ・sqlcnnSH→sqlcnn<br/>
        ''' ・前提：発行クエリ(sqlstr)は1データ抽出クエリとします
        ''' </remarks>
        Public Shared Function Exec_Scalar( _
                                         ByVal sqlstr As String, _
                                         ByRef sqlcnn As System.Data.SqlClient.SqlConnection _
                                         ) As Object
            '★★★ CommonRepository.Exec_Scalarを参考
            Dim sqlcom As SqlCommand = New SqlCommand(sqlstr, sqlcnn)
            Dim reccnt As Object = IIf(IsDBNull(sqlcom.ExecuteScalar()), 0, sqlcom.ExecuteScalar())
            sqlcom.Dispose()
            Return reccnt

        End Function

        ''' <summary>
        ''' DBより1レコードを取得してHashTableへ格納 [ExecuteReader]
        ''' </summary>
        ''' <param name="sqlstr">抽出クエリ文</param>
        ''' <param name="sqlcnn">接続情報</param>
        ''' <param name="hashtbl">返却抽出データ</param>
        ''' <returns>抽出レコード結果：True.あり False.なし</returns>
        ''' <remarks>
        ''' ・前提：発行クエリ(sqlstr)は1レコード抽出クエリとします<br/>
        ''' </remarks>
        Public Shared Function Exec_DataReader( _
                                            ByVal sqlstr As String, _
                                            ByRef sqlcnn As System.Data.SqlClient.SqlConnection, _
                                            ByRef hashtbl As Hashtable _
                                            ) As Boolean
            '★★★ CommonRepository.Exec_DataReaderを参考
            Dim sqlcom As SqlCommand = New SqlCommand(sqlstr, sqlcnn)
            Dim sqldrd As SqlDataReader = sqlcom.ExecuteReader(CommandBehavior.SingleRow)
            Dim recumuflg As Boolean                                'True.レコードあり

            sqldrd.Read()
            recumuflg = sqldrd.HasRows
            If recumuflg Then
                hashtbl = New Hashtable
                For cnt As Integer = 0 To sqldrd.FieldCount - 1
                    hashtbl.Add(sqldrd.GetName(cnt), sqldrd.GetValue(cnt))
                Next
            End If

            sqlcom.Dispose()
            sqldrd.Close()
            Return recumuflg

        End Function

        ''' <summary>
        ''' DBより全レコードを取得して1列目のデータをリストへ格納
        ''' 前提：抽出したテーブルの最初の1列のみ取得する
        ''' </summary>
        ''' <param name="sqlstr"></param>
        ''' <param name="sqlcnn"></param>
        ''' <param name="list"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Exec_DataReader_Col_List( _
                                                    ByVal sqlstr As String, _
                                                    ByRef sqlcnn As System.Data.SqlClient.SqlConnection, _
                                                    ByRef list As List(Of String) _
                                                    ) As Boolean

            Dim sqlcom As SqlCommand = New SqlCommand(sqlstr, sqlcnn)
            Dim sqldrd As SqlDataReader = sqlcom.ExecuteReader()
            Dim recumuflg As Boolean                                'True.レコードあり

            recumuflg = sqldrd.HasRows
            If recumuflg Then
                list = New List(Of String)
                While sqldrd.Read
                    list.Add(sqldrd.GetValue(0))
                End While
            End If

            sqlcom.Dispose()
            sqldrd.Close()
            Return recumuflg

        End Function

        ''' <summary>
        ''' DBより全レコードを取得してハッシュテーブルへ格納
        ''' 前提：2列取得し、1列目の値をハッシュテーブルのキーとする
        ''' </summary>
        ''' <param name="sqlstr"></param>
        ''' <param name="sqlcnn"></param>
        ''' <param name="hash"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Exec_DataReader_Col_Hash( _
                                                    ByVal sqlstr As String, _
                                                    ByRef sqlcnn As System.Data.SqlClient.SqlConnection, _
                                                    ByRef hash As Hashtable _
                                                    ) As Boolean

            Dim sqlcom As SqlCommand = New SqlCommand(sqlstr, sqlcnn)
            Dim sqldrd As SqlDataReader = sqlcom.ExecuteReader()
            Dim recumuflg As Boolean                                'True.レコードあり

            recumuflg = sqldrd.HasRows
            If recumuflg Then
                hash = New Hashtable
                While sqldrd.Read
                    Dim valuekey As String = sqldrd.GetValue(0).ToString
                    Dim value As String = sqldrd.GetValue(1).ToString
                    '20160527 指摘事項対応 -chg sta
                    'hash.Add(valuekey, value)
                    If hash.Contains(valuekey) = False Then
                        hash.Add(valuekey, value)
                    End If
                    '20160527 指摘事項対応 -chg end
                End While
            End If

            sqlcom.Dispose()
            sqldrd.Close()
            Return recumuflg

        End Function

        ''' <summary>
        ''' クエリ実行
        ''' </summary>
        ''' <param name="sqlcnnv10">接続情報</param>
        ''' <param name="qry">クエリ文</param>
        ''' <param name="rowcnt">処理件数</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Exec_NonQuery(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection, _
                                      ByVal qry As String, _
                                      ByRef rowcnt As Integer
                                      ) As Boolean
            Dim rtn As Boolean = True
            Try
                Dim sqlcom = New SqlCommand(qry, sqlcnnv10)
                rowcnt = sqlcom.ExecuteNonQuery()
                sqlcom.Dispose()
            Catch ex As Exception
                rtn = False
            End Try
            Return rtn

        End Function

    End Class

#End Region

#Region "個別処理クラス"

    ''' <summary>
    ''' 【個別処理クラス】
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EtcMethod

        ''' <summary>
        ''' DB登録用パラメータ文字列作成
        ''' </summary>
        ''' <param name="maxcnt">パラメータ最大値</param>
        ''' <param name="addcnt">パラメータ開始番号</param>
        ''' <param name="skipstr">
        ''' スキップ番号(","で連結)<br/>
        ''' 例："5,12,18"
        ''' →Para5,Para12,Para18は作成しない<br/>
        ''' </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_ParaNum( _
                                          ByVal maxcnt As Integer, _
                                          Optional ByVal addcnt As Integer = 0, _
                                          Optional ByVal skipstr As String = "" _
                                          ) As String

            Dim sbobj As New System.Text.StringBuilder      'StringBuilderオブジェクト生成
            Dim sumcnt As Integer                           '作業用カウント値
            Dim skipcnt() As String                         '分割SKIPカウント値
            Dim skipflg As Boolean = False                  'スキップFLG：True.Skip
            '★★★ CommonRepository.Get_ParaNumを参考
            skipstr = IIf(skipstr = Nothing, "", skipstr)
            If skipstr.Trim <> "" Then
                skipcnt = skipstr.Split(",")
            End If

            addcnt = IIf(addcnt = 0, 0, addcnt - 1)
            For cnti As Integer = 1 To maxcnt
                sumcnt = cnti + addcnt

                If Not skipcnt Is Nothing Then
                    For cntj As Integer = 0 To UBound(skipcnt)
                        If cnti = CInt(skipcnt(cntj)) Then
                            skipflg = True
                            Exit For
                        End If
                    Next
                End If
                If skipflg = False Then
                    sbobj.Append("@para")
                    sbobj.Append(EtcMethod.RightStrMethod(("000" & sumcnt), 3))
                    If cnti <= maxcnt - 1 Then
                        sbobj.Append(",")
                    End If
                End If
                skipflg = False
            Next
            Return sbobj.ToString

        End Function

        ''' <summary>
        ''' 文字列抽出(RIGHTメソッド)
        ''' </summary>
        ''' <param name="str">対象文字列<br/></param>
        ''' <param name="lenght">抽出文字数<br/></param>
        ''' <returns>抽出文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function RightStrMethod(ByVal str As String, ByVal lenght As Integer) As String
            '★★★ RightStrMethodを参考
            Dim rtn As String

            rtn = str
            If lenght <= str.Length Then
                rtn = str.Substring(str.Length - lenght)
            End If
            Return rtn

        End Function

        ''' <summary>
        ''' 検索文字カウント数
        ''' </summary>
        ''' <param name="str">検索元文字列</param>
        ''' <param name="chr">検索文字</param>
        ''' <returns>検索文字カウント数</returns>
        ''' <remarks></remarks>
        Public Shared Function CountChar(ByVal str As String, ByVal chr As Char) As Integer

            Return str.Length - str.Replace(chr.ToString(), "").Length

        End Function

        ''' <summary>
        ''' 対象文字列のバイト数取得
        ''' </summary>
        ''' <param name="target">対象文字列</param>
        ''' <returns>半角1バイト、全角2バイトでカウントされたバイト数</returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LenB(ByVal target As String) As Integer

            Return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(target)

        End Function

        ''' <summary>
        ''' 英数字半角変換(一部記号を含む)
        ''' </summary>
        ''' <param name="str">対象文字列</param>
        ''' <returns>変換後文字列</returns>
        ''' <remarks>
        ''' ・変換文字列：０-９Ａ-Ｚａ-ｚ：，－　<br/>
        ''' </remarks>
        Public Shared Function Chg_NumNarrow(ByVal str As String) As String

            Dim reg As Regex = New Regex("[０-９Ａ-Ｚａ-ｚ：，－　]+")
            Dim output As String = reg.Replace(str, AddressOf RepNarrow)
            Return output

        End Function

        ''' <summary>
        ''' 半角文字列変換
        ''' </summary>
        ''' <param name="mat">一致文字列</param>
        ''' <returns>変換後文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function RepNarrow(ByVal mat As Match) As String

            Return Strings.StrConv(mat.Value, VbStrConv.Narrow, 0)

        End Function

        ''' <summary>
        ''' 値が基準値の倍数か判断  2016.02.22 プログレスバーの表示修正
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="base"></param>
        ''' <param name="chgvalue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Sub Get_MultipleFlg(ByVal value As Integer, ByVal base As Integer, Optional ByRef chgvalue As Integer = 0)

            Dim tmp_result As Integer = 0

            '値が0の時は処理を抜ける
            If value = 0 Then
                Exit Sub
            End If

            '基準値で割った値が整数か否か判別
            Dim tmp_value As String = (value / base).ToString
            Dim flg As Boolean = Int32.TryParse(tmp_value, tmp_result)

            '基準値の倍数の場合かつ割った値が必要な場合は返す
            If flg Then
                chgvalue = tmp_result
            End If

        End Sub

        ''' <summary>
        ''' 四捨五入処理 2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_RoundingValue(ByVal value As String)

            Dim tmp_dbl As Double = 0
            Dim rtn_str As String = ""

            If Double.TryParse(value, tmp_dbl) Then
                '四捨五入して取得
                rtn_str = Math.Round(tmp_dbl, MidpointRounding.AwayFromZero).ToString
            Else
                '変換できない場合は値をそのまま格納 (ログ出力させる)
                rtn_str = value
            End If

            Return rtn_str

        End Function

        ''' <summary>
        ''' 入金項目区分から区分名を取得 2016.04.06 入金項目読込処理の修正
        ''' </summary>
        ''' <param name="tukikbn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Nkinruiname(ByRef tukikbn As String) As String

            Dim rtn_string As String = ""

            Select Case tukikbn
                Case "1" : rtn_string = tukikbn & "." & "通常月"
                Case "2" : rtn_string = tukikbn & "." & "契約時"
                Case "3" : rtn_string = tukikbn & "." & "更新時"
                Case "4" : rtn_string = tukikbn & "." & "解約時"
                Case "5" : rtn_string = tukikbn & "." & "随時変動"
                Case "6" : rtn_string = tukikbn & "." & "その他"
                Case "7" : rtn_string = tukikbn & "." & "修繕"
                Case "8" : rtn_string = tukikbn & "." & "家主送金"
                Case "9" : rtn_string = tukikbn & "." & "家主控除"
            End Select

            Return rtn_string

        End Function

        ''' <summary>
        ''' 10設備マスタから項目guidを取得してハッシュテーブルへ格納(setubi_grpsortorder-setubi_sortorder-komok_sortorder,komk_guid) 2016.04.06 部屋設備情報の取得処理修正
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_SetubiMst(ByVal sqlcnnv10 As SqlConnection)

            Dim tmp_sql As String = ""

            '抽出クエリの作成
            tmp_sql = tmp_sql & " SELECT "
            tmp_sql = tmp_sql & " 	 CONVERT(varchar,setubi_grpsortorder) + '-' + "
            tmp_sql = tmp_sql & " 	 CONVERT(varchar,setubi_sortorder) + '-' + "
            tmp_sql = tmp_sql & " 	 CONVERT(varchar,komok_sortorder) AS キー "
            tmp_sql = tmp_sql & " 	,komok_guid AS [項目guid] "
            tmp_sql = tmp_sql & " /* "
            tmp_sql = tmp_sql & " 	,SE.setubi_guid "
            tmp_sql = tmp_sql & " 	,komok_guid  "
            tmp_sql = tmp_sql & " 	,setubi_grpsortorder "
            tmp_sql = tmp_sql & " 	,setubi_grpname  "
            tmp_sql = tmp_sql & " 	,setubi_sortorder "
            tmp_sql = tmp_sql & " 	,setubi_name "
            tmp_sql = tmp_sql & " 	,komok_sortorder "
            tmp_sql = tmp_sql & " 	,komok_name "
            tmp_sql = tmp_sql & " */ "
            tmp_sql = tmp_sql & " FROM m_setubi_grp AS SEG "
            tmp_sql = tmp_sql & " LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid "
            tmp_sql = tmp_sql & " LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid "
            tmp_sql = tmp_sql & " ORDER BY setubi_grpsortorder,setubi_sortorder,komok_sortorder "


            '設備のソートNoが重複しているデータが存在するためとりあえず重複を除去したクエリを作成しておく
            tmp_sql = ""
            tmp_sql = tmp_sql & " /*20160829 エレベーター移行対応 chg sta*/ "
            tmp_sql = tmp_sql & " /* "
            tmp_sql = tmp_sql & " SELECT [キー],[項目guid] FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & "     SELECT "
            tmp_sql = tmp_sql & "             ROW_NUMBER()OVER(PARTITION BY [キー] ORDER BY [キー]) AS [抽出用] "
            tmp_sql = tmp_sql & "         ,* "
            tmp_sql = tmp_sql & "     FROM "
            tmp_sql = tmp_sql & "     ( "
            tmp_sql = tmp_sql & "         SELECT "
            tmp_sql = tmp_sql & "             	CONVERT(varchar,setubi_grpsortorder) + '-' + "
            tmp_sql = tmp_sql & "             	CONVERT(varchar,setubi_sortorder) + '-' + "
            tmp_sql = tmp_sql & "             	CONVERT(varchar,komok_sortorder) AS [キー] "
            tmp_sql = tmp_sql & "             ,komok_guid AS [項目guid] "
            tmp_sql = tmp_sql & "         /* "
            tmp_sql = tmp_sql & "             ,SE.setubi_guid "
            tmp_sql = tmp_sql & "             ,komok_guid  "
            tmp_sql = tmp_sql & "             ,setubi_grpsortorder "
            tmp_sql = tmp_sql & "             ,setubi_grpname  "
            tmp_sql = tmp_sql & "             ,setubi_sortorder "
            tmp_sql = tmp_sql & "             ,setubi_name "
            tmp_sql = tmp_sql & "             ,komok_sortorder "
            tmp_sql = tmp_sql & "             ,komok_name "
            tmp_sql = tmp_sql & "         */ "
            tmp_sql = tmp_sql & "         FROM m_setubi_grp AS SEG "
            tmp_sql = tmp_sql & "         LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid "
            tmp_sql = tmp_sql & "         LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid "
            tmp_sql = tmp_sql & "     ) AS VW1 "
            tmp_sql = tmp_sql & " ) AS VW2 "
            tmp_sql = tmp_sql & " WHERE [抽出用] = 1 "
            tmp_sql = tmp_sql & " */ "
            tmp_sql = tmp_sql & " SELECT * FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT [キー],CONVERT(varchar(MAX),[項目guid]) AS [項目guid] FROM "
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		SELECT "
            tmp_sql = tmp_sql & " 				ROW_NUMBER()OVER(PARTITION BY [キー] ORDER BY [キー]) AS [抽出用] "
            tmp_sql = tmp_sql & " 			,* "
            tmp_sql = tmp_sql & " 		FROM "
            tmp_sql = tmp_sql & " 		( "
            tmp_sql = tmp_sql & " 			SELECT "
            tmp_sql = tmp_sql & " 					CONVERT(varchar,setubi_grpsortorder) + '-' + "
            tmp_sql = tmp_sql & " 					CONVERT(varchar,setubi_sortorder) + '-' + "
            tmp_sql = tmp_sql & " 					CONVERT(varchar,komok_sortorder) AS [キー] "
            tmp_sql = tmp_sql & " 				,komok_guid AS [項目guid] "
            tmp_sql = tmp_sql & " 			/* "
            tmp_sql = tmp_sql & " 				,SE.setubi_guid "
            tmp_sql = tmp_sql & " 				,komok_guid  "
            tmp_sql = tmp_sql & " 				,setubi_grpsortorder "
            tmp_sql = tmp_sql & " 				,setubi_grpname  "
            tmp_sql = tmp_sql & " 				,setubi_sortorder "
            tmp_sql = tmp_sql & " 				,setubi_name "
            tmp_sql = tmp_sql & " 				,komok_sortorder "
            tmp_sql = tmp_sql & " 				,komok_name "
            tmp_sql = tmp_sql & " 			*/ "
            tmp_sql = tmp_sql & " 			FROM m_setubi_grp AS SEG "
            tmp_sql = tmp_sql & " 			LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid "
            tmp_sql = tmp_sql & " 			LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid "
            tmp_sql = tmp_sql & " 		) AS VW1 "
            tmp_sql = tmp_sql & " 	) AS VW2 "
            tmp_sql = tmp_sql & " 	WHERE [抽出用] = 1 "
            tmp_sql = tmp_sql & " ) AS VW3 "
            tmp_sql = tmp_sql & " UNION SELECT '999-1' AS [キー],'1' AS [項目guid] "
            tmp_sql = tmp_sql & " UNION SELECT '999-2' AS [キー],'2' AS [項目guid] "
            tmp_sql = tmp_sql & " UNION SELECT '999-3' AS [キー],'3' AS [項目guid] "
            tmp_sql = tmp_sql & " /*20160829 エレベーター移行対応 chg end*/ "

            '初期化
            Hash_SetubiMst.Clear()

            'キーNoと項目guidを紐付けたハッシュテーブルを作成
            DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, Hash_SetubiMst)

        End Sub

        ''' <summary>
        ''' 10設備マスタから項目guidを取得してハッシュテーブルへ格納(ユーザー作成設備分) '20160829 設備の新規挿入処理を追加 -add
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_SetubiMst_UserMake(ByVal sqlcnnv10 As SqlConnection)

            Dim tmp_sql As String = ""

            '抽出クエリの作成
            tmp_sql = tmp_sql & " SELECT * FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & "     SELECT [キー],CONVERT(varchar(MAX),[項目guid]) AS [項目guid] FROM "
            tmp_sql = tmp_sql & "     ( "
            tmp_sql = tmp_sql & "         SELECT "
            tmp_sql = tmp_sql & "             	ROW_NUMBER()OVER(PARTITION BY [キー] ORDER BY [キー]) AS [抽出用] "
            tmp_sql = tmp_sql & "             ,* "
            tmp_sql = tmp_sql & "         FROM "
            tmp_sql = tmp_sql & "         ( "
            tmp_sql = tmp_sql & "             SELECT "
            tmp_sql = tmp_sql & "             	 CONVERT(varchar,setubi_name) + '-' + "
            tmp_sql = tmp_sql & "             	 CONVERT(varchar,komok_name) AS [キー] "
            tmp_sql = tmp_sql & "             	,komok_guid AS [項目guid] "
            tmp_sql = tmp_sql & "             FROM m_setubi_grp AS SEG "
            tmp_sql = tmp_sql & "             LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid "
            tmp_sql = tmp_sql & "             LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid "
            tmp_sql = tmp_sql & " 			WHERE setubi_grpname = 'ユーザー作成設備' "
            tmp_sql = tmp_sql & "         ) AS VW1 "
            tmp_sql = tmp_sql & "     ) AS VW2 "
            tmp_sql = tmp_sql & "     WHERE [抽出用] = 1 "
            tmp_sql = tmp_sql & " ) AS VW3 "

            '初期化
            Hash_SetubiMst_UserMake.Clear()

            'キーNoと項目guidを紐付けたハッシュテーブルを作成
            DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, Hash_SetubiMst_UserMake)

        End Sub

        ''' <summary>
        ''' 10鍵タイトルマスタから鍵Noとタイトル名を取得してハッシュテーブルへ格納 2016.04.06 物件鍵情報の移行処理修正
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_KagiTitleMst(ByVal sqlcnnv10 As SqlConnection, ByVal kagikbn As Integer)

            Dim tmp_sql As String = ""

            '抽出クエリの作成
            tmp_sql = tmp_sql & " SELECT kagi_name,kagi_no FROM m_kagi_title "
            tmp_sql = tmp_sql & " WHERE kagi_kbn = " & kagikbn.ToString
            tmp_sql = tmp_sql & " ORDER BY kagi_no "

            Select Case kagikbn
                Case 1
                    '初期化
                    Hash_KagiTitleKyoyo.Clear()
                    'キーNoと項目guidを紐付けたハッシュテーブルを作成
                    DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, Hash_KagiTitleKyoyo)
                Case 2
                    '初期化
                    Hash_KagiTitleSenyo.Clear()
                    'キーNoと項目guidを紐付けたハッシュテーブルを作成
                    DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, Hash_KagiTitleSenyo)
            End Select

        End Sub

        ''' <summary>
        ''' 10クレーム分類マスタから分類Noと分類名を取得してハッシュテーブルへ格納 '20160524 クレーム情報移行処理実装
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_ClaimBruiMst(ByVal sqlcnnv10 As SqlConnection, ByVal claimkbn As Integer)

            Dim tmp_sql As String = ""

            '抽出クエリの作成
            tmp_sql = tmp_sql & " SELECT claim_name,claim_ruino FROM m_claim_rui "
            tmp_sql = tmp_sql & " WHERE claim_ruikbn = " & claimkbn.ToString
            tmp_sql = tmp_sql & " ORDER BY claim_ruino "

            Select Case claimkbn
                Case 1
                    '初期化
                    Hash_ClaimBruiKasyo.Clear()
                    'キーNoと項目guidを紐付けたハッシュテーブルを作成
                    DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, Hash_ClaimBruiKasyo)
                Case 2
                    '初期化
                    Hash_ClaimBruiClaim.Clear()
                    'キーNoと項目guidを紐付けたハッシュテーブルを作成
                    DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, Hash_ClaimBruiClaim)
            End Select

        End Sub

        ''' <summary>
        ''' Trim処理を行い文字列として返却する 20160413 ユーザーデータテストでの修正
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_ShapeStr(ByVal value As Object) As String

            Dim rtn_str As String = ""
            Dim tmp_str As String = IIf(value Is Nothing, "", value)

            tmp_str = tmp_str.ToString.Trim

            rtn_str = tmp_str

            Return rtn_str

        End Function

        ''' <summary>
        ''' 抽出クエリの抽出項目に各文字列置換関数を付加して返却
        ''' </summary>
        ''' <param name="fldnamegrp">抽出対象をカンマ区切りで連結した文字列</param>
        ''' <param name="deleteflg">改行除去判別フラグ True:除去 False:改行用置換文字列に変換</param>
        ''' <param name="commaflg">カンマ変換フラグ True:全角変換 False:何もしない</param>
        ''' <returns>置換関数を付加した抽出対象の連結文字列</returns>
        ''' <remarks>
        ''' 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加
        ''' 　引数に Optional commaflg As Boolean = False を追加
        ''' 　CSV出力時の抽出クエリはカンマを全角変換するため追加したフラグで条件分岐を行う
        ''' </remarks>
        Public Shared Function Get_FldnameAddReplaceMethod(ByVal fldnamegrp As String, ByVal deleteflg As Boolean, Optional commaflg As Boolean = False) As String

            Dim rtn_str As String = ""
            Dim tmp_chgfldname As String = ""
            Dim tmp_fldname() As String = fldnamegrp.Split(",")
            Dim replacestr As String = ""

            '置換文字列設定 (引数のフラグで制御)
            If deleteflg Then
                replacestr = ""
            Else
                replacestr = LINE_BREAK
            End If

            For cntii = 0 To UBound(tmp_fldname)
                '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
                '↓↓↓旧srcコメントアウト↓↓↓
                'Dim tmp_str As String = tmp_fldname(cntii)
                ''tmp_str = "REPLACE(" & tmp_str & ", CHAR(13) + CHAR(10), '" & replacestr & "') AS " & tmp_str  '改行が残存している箇所が確認できたため変更
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''改行に加えてタブも確認できたため除去する
                ''tmp_str = "REPLACE(REPLACE(" & tmp_str & ", CHAR(13),'" & replacestr & "'), CHAR(10),'') AS " & tmp_str
                'tmp_str = "REPLACE(REPLACE(REPLACE(" & tmp_str & ", CHAR(13),'" & replacestr & "'), CHAR(10),''), CHAR(9),'') AS " & tmp_str
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                'tmp_chgfldname = tmp_chgfldname & "," & tmp_str
                '↑↑↑旧srcコメントアウト↑↑↑

                '※コメントが大きくなったのでまとめてコメントアウトしています。
                Dim tmp_str As String = tmp_fldname(cntii)
                If commaflg Then
                    tmp_str = "NULLIF(RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(" & tmp_str & ", CHAR(13),'" & replacestr & "'), CHAR(10),''), CHAR(9),''),'''','’'),',','，')),SPACE(0))"
                Else
                    tmp_str = "NULLIF(RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(" & tmp_str & ", CHAR(13),'" & replacestr & "'), CHAR(10),''), CHAR(9),''),'''','’')),SPACE(0))"
                End If
                tmp_str = tmp_str & " AS " & tmp_fldname(cntii)
                tmp_chgfldname = tmp_chgfldname & "," & tmp_str
                '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
            Next

            tmp_chgfldname = tmp_chgfldname.Remove(0, 1)

            rtn_str = tmp_chgfldname

            Return rtn_str

        End Function

        ''' <summary>
        ''' ハッシュテーブルのキーを半角変換して返す '2016.04.26 メインの方へも反映させる修正
        ''' </summary>
        ''' <param name="hash"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_HashKeyChg(ByVal hash As Hashtable) As Hashtable

            Dim rtn_hash As New Hashtable

            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
            'For Each item In hash

            '    Dim tmp_key As String = item.Key
            '    Dim tmp_value As String = item.Value

            '    tmp_key = StrConv(tmp_key, VbStrConv.Narrow)

            '    rtn_hash.Add(tmp_key, tmp_value)

            'Next
            rtn_hash = hash
            '20160819 部屋の半角変換処理によるエラー修正 -chg end

            Return rtn_hash

        End Function

        ''' <summary>
        ''' ハッシュテーブルのキーを取得してリストオブジェクトへ格納 '2016.04.26 メインの方へも反映させる修正
        ''' </summary>
        ''' <param name="hash"></param>
        ''' <param name="list_basedata"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_HashKeyToList(ByVal hash As Hashtable, ByRef list_basedata As List(Of String))

            For Each item In hash

                Dim tmp_key As String = item.Key
                list_basedata.Add(tmp_key)

            Next

        End Sub

        ''' <summary>
        ''' 画像ファイルの拡張子のリスト作成 '20160525 クレーム関連ファイルの画像判別処理実装
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_ImportableImageFileAttributesList()

            '画像ファイルの拡張子のリスト作成
            List_ImportableImageFileAttributes.Add(".jpg")
            List_ImportableImageFileAttributes.Add(".jpeg")
            List_ImportableImageFileAttributes.Add(".png")
            List_ImportableImageFileAttributes.Add(".gif")
            List_ImportableImageFileAttributes.Add(".tif")
            List_ImportableImageFileAttributes.Add(".tiff")
            List_ImportableImageFileAttributes.Add(".wmf")
            List_ImportableImageFileAttributes.Add(".bmp")

        End Sub

        '20160707 共通処理へ移動 -add sta
        '移動に伴い Private → Public Shared へ変更
        ''' <summary>
        ''' ファイル有無チェック
        ''' </summary>
        ''' <param name="filepath"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_FileExist(ByVal filepath As String) As Boolean

            If Not (File.Exists(filepath)) Then
                Return False
            End If
            Return True

        End Function

        ''' <summary>
        ''' フォルダ有無チェック 
        ''' </summary>
        ''' <param name="filepath"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DirExist(ByVal dirpath As String) As Boolean

            If Not (Directory.Exists(dirpath)) Then
                Return False
            End If
            Return True

        End Function

        ''' <summary>
        ''' ファイルオープンチェック '20160825 ファイルが開かれているかチェックする機能を追加 -add
        ''' </summary>
        ''' <param name="filepath"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_FileOpen(ByVal filepath As String) As Boolean

            Dim filename As String = Path.GetFileName(filepath)
            Dim filekaku As String = Path.GetExtension(filepath)
            Dim dirpath As String = filepath.Replace(filename, "")
            Dim chgfilename As String = Set_Path(dirpath, "temp" & filekaku)

            '念の為
            If Chk_FileExist(filepath) = False Then

            End If

            Try
                'ファイル名を変更して、使用中かチェックする
                File.Move(filepath, chgfilename)
                'ファイル名を元に戻す
                File.Move(chgfilename, filepath)
                'ファイル名の変更が成功したので、使用中ではない
                Return True
            Catch ex As Exception
                'ファイル名が変更できないので、使用中とする
                Return False
            End Try


            Return True

        End Function

        ''' <summary>
        ''' ファイル/フォルダのフルパスを設定  
        ''' </summary>
        ''' <param name="path"></param>
        ''' <param name="name"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Set_Path(ByVal path As String, ByVal name As String) As String

            Dim rtn_path As String = ""

            If EtcMethod.RightStrMethod(path, 1) = "\" Then
                rtn_path = path & name
            Else
                rtn_path = path & "\" & name
            End If

            Return rtn_path

        End Function
        '20160707 共通処理へ移動 -add end

        ''' <summary>
        ''' 親ヘッダーと子ヘッダーを結合してオブジェクトへ格納 20160913_2 部屋設備移行処理の追加
        ''' </summary>
        ''' <param name="objmain"></param>
        ''' <param name="objsub"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Set_Header(ByVal objmain As Object, ByVal objsub As Object) As Object

            Dim rtn_arry(0, objmain.Length) As String

            For cntii = 1 To objmain.Length

                Dim tmp_main As String = IIf(objmain(1, cntii) Is Nothing, "", objmain(1, cntii))
                Dim tmp_sub As String = IIf(objsub(1, cntii) Is Nothing, "", objsub(1, cntii))

                Dim tmp_total As String = ""
                If tmp_main <> "" And tmp_sub <> "" Then
                    tmp_total = tmp_main & STR_SPLIT_1 & tmp_sub
                ElseIf tmp_main <> "" And tmp_sub = "" Then
                    tmp_total = tmp_main
                ElseIf tmp_main = "" And tmp_sub <> "" Then
                    tmp_total = tmp_sub
                ElseIf tmp_main = "" And tmp_sub = "" Then
                    tmp_total = ""
                End If

                rtn_arry(0, cntii) = tmp_total

            Next

            Return rtn_arry

        End Function

        ''' <summary>
        ''' 汎用中間ファイルのヘッダー取得処理 '20161006 中間ファイルチェック処理の速度改善対応 -add
        ''' グループ名と設備項目名を
        ''' </summary>
        ''' <param name="solist_grp"></param>
        ''' <param name="solist_komk"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Set_SetubiHeader(ByVal solist_grp As SortedList(Of Integer, String), _
                                                ByRef solist_komk As SortedList(Of Integer, String)) As SortedList(Of Integer, String)

            Dim rtn_solist As New SortedList(Of Integer, String)

            For cntii = 0 To solist_grp.Count - 1

                Dim tmp_main As String = IIf(solist_grp(cntii) Is Nothing, "", solist_grp(cntii))
                Dim tmp_sub As String = IIf(solist_komk(cntii) Is Nothing, "", solist_komk(cntii))

                Dim tmp_total As String = ""
                If tmp_main <> "" And tmp_sub <> "" Then
                    tmp_total = tmp_main & STR_SPLIT_1 & tmp_sub
                ElseIf tmp_main <> "" And tmp_sub = "" Then
                    tmp_total = tmp_main
                ElseIf tmp_main = "" And tmp_sub <> "" Then
                    tmp_total = tmp_sub
                ElseIf tmp_main = "" And tmp_sub = "" Then
                    tmp_total = ""
                End If

                rtn_solist.Add(cntii, tmp_total)

            Next

            Return rtn_solist

        End Function




        ''' <summary>
        ''' 10設備マスタから項目guidを取得してハッシュテーブルへ格納(汎用コンバート用) '20160913_2 部屋設備移行処理の追加
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_SetubiMst_BaseMId(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT  "
            tmp_sql = tmp_sql & " 	 SEG.setubi_grpname + '@#@' + setubi_name + '@#@' + CONVERT(VARCHAR,komok_sortorder) "
            tmp_sql = tmp_sql & " 	,CONVERT(VARCHAR(MAX),komok_guid) "
            tmp_sql = tmp_sql & " FROM m_setubi_grp AS SEG "
            tmp_sql = tmp_sql & " LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid "
            tmp_sql = tmp_sql & " LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid "

            '初期化
            hash.Clear()

            'キーNoと項目guidを紐付けたハッシュテーブルを作成
            DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

        End Sub

    End Class

#End Region

#Region "共通クエリクラス"

    ''' <summary>
    ''' 【共通クエリクラス】
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DBQuery

        ''' <summary>
        ''' 削除(DELETE)クエリ
        ''' </summary>
        ''' <param name="reptbl">FROM句(例：m_bk_rui)</param>
        ''' <param name="repwhere">WHERE句 [optional]</param>
        ''' <returns>返却クエリ</returns>
        ''' <remarks>
        ''' ・DELETE FROM [reptbl] WHERE [repwhere]
        ''' </remarks>
        Public Shared Function Qry_DelInfo( _
                                           ByVal reptbl As String, _
                                           Optional ByVal repwhere As String = "" _
                                          ) As String
            '★★★ CommonRepository.Get_UseQueryを参考
            Dim qry_rtn As String                                   '返却クエリ文
            Dim qry_tblname As String = ""                          'FROM句
            Dim qry_where As String = ""                            'WHERE句

            qry_tblname = " DELETE FROM " & reptbl
            qry_where = IIf(repwhere.Trim = "", "", " WHERE " & repwhere)
            qry_rtn = qry_tblname & qry_where
            Return qry_rtn

        End Function

        ''' <summary>
        ''' テーブル/ビュー削除(DROP)クエリ
        ''' </summary>
        ''' <param name="targetname"></param>
        ''' <param name="tableflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Qry_DropInfo(ByVal targetname As String, ByVal tableflg As Boolean) As String

            Dim droptarget As String = ""
            Dim qry_rtn As String = ""

            '削除対象選択
            If tableflg Then
                droptarget = "TABLE"
            Else
                droptarget = "VIEW"
            End If

            qry_rtn = qry_rtn & "  IF EXISTS "
            qry_rtn = qry_rtn & " 	( "
            qry_rtn = qry_rtn & " 		SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[" & targetname & "]') "
            qry_rtn = qry_rtn & " 		AND OBJECTPROPERTY(id, N'Is" & droptarget & "') = 1 "
            qry_rtn = qry_rtn & " 	) "
            qry_rtn = qry_rtn & " DROP " & droptarget & " [dbo].[" & targetname & "]; "

            Return qry_rtn

        End Function

        ''' <summary>
        ''' ストアド削除(DROP)クエリ '20160829 口座名義カナチェック機能の追加
        ''' </summary>
        ''' <param name="targetname"></param>
        ''' <param name="tableflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Qry_DropFnInfo(ByVal targetname As String) As String

            Dim qry_rtn As String = ""
            qry_rtn = qry_rtn & " IF OBJECT_ID (N'[dbo].[" & targetname & "]', N'FN') IS NOT NULL "
            qry_rtn = qry_rtn & " DROP FUNCTION [dbo].[" & targetname & "]; "
            Return qry_rtn

        End Function

        ''' <summary>
        ''' 挿入クエリ(パラメータで記載)
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <param name="fldnamegrp"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Qry_Insert_ParameterSet(ByVal tblname As String, ByVal fldnamegrp As String, ByRef para() As String) As String

            Dim tmp_para As String = ""

            '取得したフィールド名(カンマで連結)を分割して配列へ格納
            Dim tmp_str() As String = fldnamegrp.Split(",")

            'テーブル名と各フィールドを「-」で連結してカンマ区切りで連結した文字列に変換
            For cntii = 0 To UBound(tmp_str)
                'tmp_para = tmp_para & ",@" & tblname & "-" & fldname(cntii)
                tmp_para = tmp_para & ",@" & tmp_str(cntii)
            Next
            tmp_para = tmp_para.Remove(0, 1)

            ''取得したテーブル/フィールド名("tbl-fld")を分割して配列へ格納
            'tblfldname = tmp_para.Split(",")

            'クエリ作成
            Dim tmp_sql As String = "INSERT INTO " & tblname & " (" & fldnamegrp & ") VALUES (" & tmp_para & ")"

            Return tmp_sql

        End Function

        ''' <summary>
        ''' テーブル列数取得クエリ
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Qry_GetColCount(ByVal tblname As String)

            Dim qry_rtn As String = ""

            qry_rtn = qry_rtn & " SELECT "
            qry_rtn = qry_rtn & "	COUNT(syscolumns.name) "
            qry_rtn = qry_rtn & " FROM "
            qry_rtn = qry_rtn & "	syscolumns "
            qry_rtn = qry_rtn & " INNER JOIN sysobjects ON "
            qry_rtn = qry_rtn & "	sysobjects.id = syscolumns.id "
            qry_rtn = qry_rtn & " WHERE "
            qry_rtn = qry_rtn & "	sysobjects.name = '"
            qry_rtn = qry_rtn & tblname
            qry_rtn = qry_rtn & "'"

            Return qry_rtn

        End Function

        ''' <summary>
        ''' テーブル列数取得クエリ
        ''' </summary>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Qry_GetRowCount(ByVal tblname As String)

            Dim qry_rtn As String = ""

            qry_rtn = qry_rtn & " SELECT COUNT(*) FROM "
            qry_rtn = qry_rtn & tblname

            Return qry_rtn

        End Function

        ''' <summary>
        ''' 革命10DBフィールド情報取得クエリ
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' 
        'Public Shared Function Qry_GetTableInfo() As String

        'CVDBInfoModuleへ移動


        '    Dim tmp_sql As String = ""

        '    '速度改善のためクエリの結合を減らす

        '    'tmp_sql = tmp_sql & " SELECT "
        '    'tmp_sql = tmp_sql & " 	 VW.* "
        '    'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用種別,'') AS CV用種別 "
        '    'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用項目名,'') AS CV用項目名 "
        '    'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最小値,'') AS CV用最小値 "
        '    'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最大値,'') AS CV用最大値 "
        '    'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用キーフラグ,'') AS CV用キーフラグ "
        '    'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用必須項目フラグ,'') AS CV用必須項目フラグ "
        '    'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV備考,'') AS CV備考 "
        '    'tmp_sql = tmp_sql & " FROM "
        '    'tmp_sql = tmp_sql & " ( "
        '    'tmp_sql = tmp_sql & " 	SELECT "
        '    'tmp_sql = tmp_sql & " 		 col.column_id AS [行No] "
        '    'tmp_sql = tmp_sql & " 		,CASE "
        '    'tmp_sql = tmp_sql & " 			WHEN index_column_id IS NULL THEN '' "
        '    'tmp_sql = tmp_sql & " 			ELSE '*' "
        '    'tmp_sql = tmp_sql & " 		 END AS [主キー] "
        '    'tmp_sql = tmp_sql & " 		,jpnname.value AS [種別] "
        '    'tmp_sql = tmp_sql & " 		,obj.NAME AS [TBL名] "
        '    'tmp_sql = tmp_sql & " 		,fldjpname.value AS [項目名] "
        '    'tmp_sql = tmp_sql & " 		,col.NAME AS [フィールド名] "
        '    'tmp_sql = tmp_sql & " 		,type_name(col.user_type_id) AS [属性] "
        '    'tmp_sql = tmp_sql & " 		,col.max_length AS [サイズ] "
        '    'tmp_sql = tmp_sql & " 		,CASE is_nullable "
        '    'tmp_sql = tmp_sql & " 			WHEN '1' THEN '○' "
        '    'tmp_sql = tmp_sql & " 			ELSE '' "
        '    'tmp_sql = tmp_sql & " 		 END AS [Null許容] "
        '    'tmp_sql = tmp_sql & " 		,fldcomment.value AS [説明] "
        '    'tmp_sql = tmp_sql & " 		,obj.id "
        '    'tmp_sql = tmp_sql & " 		,creater.value AS creator "
        '    'tmp_sql = tmp_sql & " 		,checker.value AS checker "
        '    'tmp_sql = tmp_sql & " 		,comment.value  AS [TBL説明] "
        '    'tmp_sql = tmp_sql & " 	FROM sys.sysobjects AS obj "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS jpnname ON obj.id = jpnname.major_id "
        '    'tmp_sql = tmp_sql & " 		AND jpnname.NAME = 'MS_Description' "
        '    'tmp_sql = tmp_sql & " 		AND jpnname.minor_id = 0 "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS creater ON obj.id = creater.major_id "
        '    'tmp_sql = tmp_sql & " 		AND creater.NAME = 'creator' "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS checker ON obj.id = checker.major_id "
        '    'tmp_sql = tmp_sql & " 		AND checker.NAME = 'checker' "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS comment ON obj.id = comment.major_id "
        '    'tmp_sql = tmp_sql & " 		AND comment.NAME = 'comment' "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.columns AS col ON obj.id = col.object_id "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS fldjpname ON col.object_id = fldjpname.major_id "
        '    'tmp_sql = tmp_sql & " 		AND col.column_id = fldjpname.minor_id "
        '    'tmp_sql = tmp_sql & " 		AND fldjpname.NAME = 'Jpfieldname' "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS fldcomment ON col.object_id = fldcomment.major_id "
        '    'tmp_sql = tmp_sql & " 		AND col.column_id = fldcomment.minor_id "
        '    'tmp_sql = tmp_sql & " 		AND fldcomment.NAME = 'MS_Description' "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.indexes AS I ON col.object_id = I.object_id "
        '    'tmp_sql = tmp_sql & " 		AND I.is_primary_key = 1 "
        '    'tmp_sql = tmp_sql & " 	LEFT JOIN sys.index_columns AS pkey ON I.object_id = pkey.object_id "
        '    'tmp_sql = tmp_sql & " 		AND col.column_id = pkey.column_id "
        '    'tmp_sql = tmp_sql & " 		AND I.index_id = pkey.index_id "
        '    'tmp_sql = tmp_sql & " 	WHERE obj.xtype = 'U' "
        '    'tmp_sql = tmp_sql & " ) AS VW "
        '    'tmp_sql = tmp_sql & " LEFT JOIN cv_dbinfo AS CVDB "
        '    'tmp_sql = tmp_sql & " ON  VW.TBL名 = CVDB.紐付用TBL名 "
        '    'tmp_sql = tmp_sql & " AND VW.フィールド名 = CVDB.紐付用フィールド名 "
        '    'tmp_sql = tmp_sql & " ORDER BY [TBL名],id,[行No] "

        '    tmp_sql = tmp_sql & " SELECT "
        '    tmp_sql = tmp_sql & " 	 VW.* "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用種別,'') AS CV用種別 "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用項目名,'') AS CV用項目名 "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最小値,'') AS CV用最小値 "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最大値,'') AS CV用最大値 "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.デフォルト値,'') AS デフォルト値 "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用キーフラグ,'') AS CV用キーフラグ "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用必須項目フラグ,'') AS CV用必須項目フラグ "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.革命マスタ参照区分,'') AS 革命マスタ参照区分 "
        '    tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV備考,'') AS CV備考 "
        '    tmp_sql = tmp_sql & " FROM "
        '    tmp_sql = tmp_sql & " ( "
        '    tmp_sql = tmp_sql & " 	SELECT "
        '    tmp_sql = tmp_sql & " 		 col.column_id AS [行No] "
        '    tmp_sql = tmp_sql & " 		,CASE "
        '    tmp_sql = tmp_sql & " 			WHEN index_column_id IS NULL THEN '' "
        '    tmp_sql = tmp_sql & " 			ELSE '*' "
        '    tmp_sql = tmp_sql & " 		 END AS [主キー] "
        '    tmp_sql = tmp_sql & " 		,jpnname.value AS [種別] "
        '    tmp_sql = tmp_sql & " 		,obj.NAME AS [TBL名] "
        '    tmp_sql = tmp_sql & " 		,fldjpname.value AS [項目名] "
        '    tmp_sql = tmp_sql & " 		,col.NAME AS [フィールド名] "
        '    tmp_sql = tmp_sql & " 		,type_name(col.user_type_id) AS [属性] "
        '    tmp_sql = tmp_sql & " 		,col.max_length AS [サイズ] "
        '    tmp_sql = tmp_sql & " 		,CASE is_nullable "
        '    tmp_sql = tmp_sql & " 			WHEN '1' THEN '○' "
        '    tmp_sql = tmp_sql & " 			ELSE '' "
        '    tmp_sql = tmp_sql & " 		 END AS [Null許容] "
        '    tmp_sql = tmp_sql & " 		 ,obj.id "
        '    tmp_sql = tmp_sql & " 		 /* "
        '    tmp_sql = tmp_sql & " 		,fldcomment.value AS [説明] "
        '    tmp_sql = tmp_sql & " 		,creater.value AS creator "
        '    tmp_sql = tmp_sql & " 		,checker.value AS checker "
        '    tmp_sql = tmp_sql & " 		,comment.value  AS [TBL説明] "
        '    tmp_sql = tmp_sql & " 		*/ "
        '    tmp_sql = tmp_sql & " 	FROM sys.sysobjects AS obj "
        '    tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS jpnname ON obj.id = jpnname.major_id "
        '    tmp_sql = tmp_sql & " 		AND jpnname.NAME = 'MS_Description' "
        '    tmp_sql = tmp_sql & " 		AND jpnname.minor_id = 0 "
        '    tmp_sql = tmp_sql & " 	LEFT JOIN sys.columns AS col ON obj.id = col.object_id "
        '    tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS fldjpname ON col.object_id = fldjpname.major_id "
        '    tmp_sql = tmp_sql & " 		AND col.column_id = fldjpname.minor_id "
        '    tmp_sql = tmp_sql & " 		AND fldjpname.NAME = 'Jpfieldname' "
        '    tmp_sql = tmp_sql & " 	LEFT JOIN sys.indexes AS I ON col.object_id = I.object_id "
        '    tmp_sql = tmp_sql & " 		AND I.is_primary_key = 1 "
        '    tmp_sql = tmp_sql & " 	LEFT JOIN sys.index_columns AS pkey ON I.object_id = pkey.object_id "
        '    tmp_sql = tmp_sql & " 		AND col.column_id = pkey.column_id "
        '    tmp_sql = tmp_sql & " 		AND I.index_id = pkey.index_id "
        '    tmp_sql = tmp_sql & " 	WHERE obj.xtype = 'U' "
        '    tmp_sql = tmp_sql & " ) AS VW "
        '    tmp_sql = tmp_sql & " LEFT JOIN cv_dbinfo AS CVDB "
        '    tmp_sql = tmp_sql & " ON  VW.TBL名 = CVDB.紐付用TBL名 "
        '    tmp_sql = tmp_sql & " AND VW.フィールド名 = CVDB.紐付用フィールド名 "
        '    tmp_sql = tmp_sql & " ORDER BY [TBL名],id,[行No] "

        '    Return tmp_sql

        'End Function

        ''' <summary>
        ''' 金融機関情報取得クエリ
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Qry_GetKinyuInfo() As String

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " SELECT "
            tmp_sql = tmp_sql & " 	 KINYU.kinyu_no "
            tmp_sql = tmp_sql & " 	,KINYU.kinyu_name "
            tmp_sql = tmp_sql & " 	,TEN.kinyu_tenno "
            tmp_sql = tmp_sql & " 	,TEN.kinyu_tenname "
            tmp_sql = tmp_sql & " FROM m_kinyu_ten AS TEN "
            tmp_sql = tmp_sql & " LEFT JOIN m_kinyu AS KINYU ON TEN.kinyu_no = KINYU.kinyu_no "
            tmp_sql = tmp_sql & " ORDER BY KINYU.kinyu_no,TEN.kinyu_tenno "

            Return tmp_sql

        End Function

    End Class

#End Region

#Region "履歴関連クラス"

    Public Class HistorySetting

        Public Property Created() As String
        Public Property CreatedHostName() As String
        Public Property CreatedOSUserName() As String
        Public Property CreatedAppUserName() As String
        Public Property Updated() As String
        Public Property UpdatedHostName() As String
        Public Property UpdatedOSUserName() As String
        Public Property UpdatedAppUserName() As String

        ''' <summary>
        ''' 履歴出力内容をセット
        ''' </summary>
        ''' <param name="recuser"></param>
        ''' <remarks></remarks>
        Public Sub Set_HistoryData(ByVal recuser As String)

            Me.Created = Now
            Me.CreatedHostName = recuser
            Me.CreatedOSUserName = Me.CreatedHostName
            Me.CreatedAppUserName = "標準ユーザ"
            Me.Updated = Me.Created
            Me.UpdatedHostName = Me.CreatedHostName
            Me.UpdatedOSUserName = Me.CreatedOSUserName
            Me.UpdatedAppUserName = Me.CreatedAppUserName

        End Sub

        ''' <summary>
        ''' タグ付きデータを作成
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ToXmlString() As String

            Dim sb = New System.Text.StringBuilder
            Dim xw As Xml.XmlWriter = System.Xml.XmlWriter.Create(sb)
            xw.WriteStartDocument()
            xw.WriteStartElement("history")

            '登録日
            If (Typ.IsDateNotNull(Me.Created) = True) Then
                xw.WriteStartElement("c")
                xw.WriteString(Me.Created)
                xw.WriteEndElement()
            End If

            '登録者
            If Me.CreatedHostName <> "" Then
                xw.WriteStartElement("chost")
                xw.WriteString(Me.CreatedHostName)
                xw.WriteEndElement()
            End If

            '登録者OSユーザ名
            If Me.CreatedOSUserName <> "" Then
                xw.WriteStartElement("cosuser")
                xw.WriteString(Me.CreatedOSUserName)
                xw.WriteEndElement()
            End If

            '権限
            If Me.CreatedAppUserName <> "" Then
                xw.WriteStartElement("cappuser")
                xw.WriteString(Me.CreatedAppUserName)
                xw.WriteEndElement()
            End If

            '更新日
            If (Typ.IsDateNotNull(Me.Updated) = True) Then
                xw.WriteStartElement("u")
                xw.WriteString(Me.Updated)
                xw.WriteEndElement()
            End If

            '更新者名
            If Me.UpdatedHostName <> "" Then
                xw.WriteStartElement("uhost")
                xw.WriteString(Me.UpdatedHostName)
                xw.WriteEndElement()
            End If

            '更新者OSユーザ名
            If Me.UpdatedOSUserName <> "" Then
                xw.WriteStartElement("uosuser")
                xw.WriteString(Me.UpdatedOSUserName)
                xw.WriteEndElement()
            End If

            '更新者権限
            If Me.UpdatedAppUserName <> "" Then
                xw.WriteStartElement("uappuser")
                xw.WriteString(Me.UpdatedAppUserName)
                xw.WriteEndElement()
            End If

            xw.WriteEndElement()
            xw.Flush()
            xw.Close()

            Dim str_xml As String = sb.ToString

            Return str_xml

        End Function

    End Class

#End Region

#Region "色変換クラス"

    Public Class Translate_Color

        ''' <summary>
        ''' RGB → ARGB 変換
        ''' </summary>
        ''' <param name="int"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Chg_ColorValue(ByVal int As Integer) As Integer

            Dim color As New System.Drawing.Color

            'RGB → ARGB 変換
            color = System.Drawing.ColorTranslator.FromOle(int)

            '変換値取得
            Dim chgvalue As Integer = color.ToArgb

            Return chgvalue

        End Function

    End Class

#End Region

#Region "Excelファイル設定クラス"

    Public Class ExcelFileManager
        '20160812 汎用コンバート対応 -del sta
        ' ''' <summary>
        ' ''' 書込時のExcelファイルオープン処理
        ' ''' </summary>
        ' ''' <param name="appli"></param>
        ' ''' <param name="wbook"></param>
        ' ''' <param name="wsheet"></param>
        ' ''' <param name="startrow"></param>
        ' ''' <param name="columncnt"></param>
        ' ''' <param name="maxrowcnt"></param>
        ' ''' <param name="dataarea"></param>
        ' ''' <param name="sheetname"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Function Set_ExcelFile_WriteOpen( _
        '                                        ByRef appli As Excel.Application, _
        '                                        ByRef wbook As Excel.Workbook, _
        '                                        ByRef wsheet As Excel.Worksheet, _
        '                                        ByRef startrow As Integer, _
        '                                        ByRef columncnt As Integer, _
        '                                        ByRef maxrowcnt As Integer, _
        '                                        ByRef dataarea As String, _
        '                                        ByVal filename As String, _
        '                                        ByVal sheetname As String _
        '                                        ) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim midfilepath As String = MiddleDirPath & "\" & filename & ".xlsx"

        '    Try
        '        'ファイル準備
        '        appli = CreateObject("Excel.Application")
        '        appli.Visible = False
        '        wbook = appli.Workbooks.Open(midfilepath)
        '        wsheet = wbook.Worksheets(sheetname)

        '        '書込開始行の設定
        '        startrow = MIDFILE_READWRITE_ROW

        '        '既存データの確認(最大行の取得)
        '        maxrowcnt = wsheet.UsedRange.Rows.Count

        '        '既存データ範囲を取得
        '        dataarea = startrow & ":" & maxrowcnt

        '        '初期化
        '        If maxrowcnt >= startrow Then
        '            wsheet.Rows(dataarea).Delete()
        '        End If

        '        'ヘッダの列数取得
        '        columncnt = wsheet.UsedRange.Columns.Count

        '    Catch ex As Exception
        '        'エラー処理を行うこと
        '        Debug.WriteLine(ex.Message)
        '        rtn = False
        '    End Try

        '    Return rtn

        'End Function
        '20160812 汎用コンバート対応 -del end
        ''' <summary>
        ''' 書込時のExcelファイルオープン処理 '20160812 汎用コンバート対応
        ''' </summary>
        ''' <param name="appli"></param>
        ''' <param name="wbook"></param>
        ''' <param name="wsheet"></param>
        ''' <param name="startrow"></param>
        ''' <param name="columncnt"></param>
        ''' <param name="maxrowcnt"></param>
        ''' <param name="dataarea"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_ExcelFile_WriteOpen( _
                                                ByRef appli As Excel.Application, _
                                                ByRef wbook As Excel.Workbook, _
                                                ByRef wsheet As Excel.Worksheet, _
                                                ByRef startrow As Integer, _
                                                ByRef columncnt As Integer, _
                                                ByRef maxrowcnt As Integer, _
                                                ByVal dirpath As String, _
                                                ByVal filename As String, _
                                                ByVal sheetname As String _
                                                ) As Boolean

            '20170601 未使用箇所のコメントアウト -del sta
            'Dim rtn As Boolean = True
            ''20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg sta
            ''Dim midfilepath As String = MiddleDirPath & "\" & filename & ".xlsx"
            'Dim midfilepath As String = dirpath & "\" & filename & ".xlsx"
            ''20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg end
            'Try
            '    'ファイル準備
            '    appli = CreateObject("Excel.Application")
            '    appli.Visible = False
            '    wbook = appli.Workbooks.Open(midfilepath)
            '    wsheet = wbook.Worksheets(sheetname)

            '    '書込開始行の設定
            '    '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
            '    'startrow = EXISTMIDFILE_READWRITE_ROW
            '    If filename = CV_FROM_MIDDLE Then
            '        startrow = BASEMIDFILE_READWRITE_ROW
            '    Else
            '        startrow = EXISTMIDFILE_READWRITE_ROW
            '    End If
            '    '20160912 中間ファイル作成に伴うプログラム修正 -chg end
            '    '既存データの確認(最大行の取得)
            '    maxrowcnt = wsheet.UsedRange.Rows.Count

            '    '既存データ範囲を取得
            '    Dim dataarea As String = startrow & ":" & maxrowcnt

            '    '初期化
            '    If maxrowcnt >= startrow Then
            '        wsheet.Rows(dataarea).Delete()
            '    End If

            '    'ヘッダの列数取得
            '    columncnt = wsheet.UsedRange.Columns.Count

            'Catch ex As Exception
            '    Debug.WriteLine(ex.Message)
            '    rtn = False
            'End Try

            'Return rtn
            '20170601 未使用箇所のコメントアウト -del end
        End Function

        ''' <summary>
        ''' 読込時のExcelファイルオープン処理
        ''' </summary>
        ''' <param name="appli"></param>
        ''' <param name="wbook"></param>
        ''' <param name="wsheet"></param>
        ''' <param name="startrow"></param>
        ''' <param name="columncnt"></param>
        ''' <param name="maxrowcnt"></param>
        ''' <param name="rowcnt"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_ExcelFile_ReadOpen( _
                                                ByRef appli As Excel.Application, _
                                                ByRef wbook As Excel.Workbook, _
                                                ByRef wsheet As Excel.Worksheet, _
                                                ByRef startrow As Integer, _
                                                ByRef columncnt As Integer, _
                                                ByRef maxrowcnt As Integer, _
                                                ByRef rowcnt As Integer, _
                                                ByVal dirpath As String, _
                                                ByVal filename As String, _
                                                ByVal sheetname As String _
                                                ) As Boolean

            Dim rtn As Boolean = True
            Dim filepath As String = dirpath & "\" & filename & ".xlsx"

            Try
                'ファイル準備
                appli = CreateObject("Excel.Application")
                appli.Visible = False
                wbook = appli.Workbooks.Open(filepath)
                wsheet = wbook.Worksheets(sheetname)

                '読込開始行の設定
                '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
                'If startrow = 0 Then
                '    startrow = MIDFILE_READWRITE_ROW
                'End If
                If filename = CV_FROM_MIDDLE Then                       '20160913 CNVNOで分岐ではなくファイル名で分岐したのは？(他の箇所も同様)
                    startrow = BASEMIDFILE_READWRITE_ROW
                Else
                    startrow = EXISTMIDFILE_READWRITE_ROW
                End If
                '20160912 中間ファイル作成に伴うプログラム修正 -chg end
                '既存データの確認(ヘッダを含めた最大行の取得)
                maxrowcnt = wsheet.UsedRange.Rows.Count

                'ヘッダの列数取得
                columncnt = wsheet.UsedRange.Columns.Count

                '行数取得(全行 - 読込開始行 + 1)
                rowcnt = maxrowcnt - startrow + 1

            Catch ex As Exception
                '20160907 EXCELが開いたままになっている不具合 -coment：エラーの場合の処理(ログ出力)がない？
                'エラー処理を行うこと
                Debug.WriteLine(ex.Message)
                rtn = False
            End Try

            Return rtn


        End Function

        ''' <summary>
        ''' 中間ファイルの件数を取得する処理 20160905 汎用コンバートの件数表示修正
        ''' </summary>
        ''' <param name="appli"></param>
        ''' <param name="wbook"></param>
        ''' <param name="wsheet"></param>
        ''' <param name="startrow"></param>
        ''' <param name="columncnt"></param>
        ''' <param name="maxrowcnt"></param>
        ''' <param name="rowcnt"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_ExcelFile_ReadOpen_BaseMidCnt( _
                                                ByRef appli As Excel.Application, _
                                                ByRef wbook As Excel.Workbook, _
                                                ByRef wsheet As Excel.Worksheet, _
                                                ByRef startrow As Integer, _
                                                ByRef maxrowcnt As Integer, _
                                                ByVal dirpath As String, _
                                                ByVal filename As String, _
                                                ByVal hash_sheet As Hashtable, _
                                                ByRef hash_cnt As Hashtable _
                                                ) As Boolean

            Dim rtn As Boolean = True
            Dim filepath As String = dirpath & "\" & filename & ".xlsx"

            Try
                'ファイル準備
                appli = CreateObject("Excel.Application")
                appli.Visible = False
                wbook = appli.Workbooks.Open(filepath)

                For Each baseitem In hash_sheet

                    Dim tmp_lbl As Label = baseitem.Key
                    Dim tmp_sheetname As String = baseitem.Value

                    'シート選択
                    wsheet = wbook.Worksheets(tmp_sheetname)

                    '読込開始行の設定
                    '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
                    'If tmp_sheetname = "部屋設備情報" Then
                    '    startrow = 3
                    'Else
                    '    startrow = EXISTMIDFILE_READWRITE_ROW
                    'End If
                    startrow = BASEMIDFILE_READWRITE_ROW
                    '20160912 中間ファイル作成に伴うプログラム修正 -chg end
                    '既存データの確認(ヘッダを含めた最大行の取得)
                    maxrowcnt = wsheet.UsedRange.Rows.Count

                    '行数取得(全行 - 読込開始行 + 1)
                    Dim rowcnt As Integer = maxrowcnt - startrow + 1

                    'ハッシュテーブルへ再格納
                    hash_cnt.Add(tmp_lbl, rowcnt)

                Next

            Catch ex As Exception
                '20160907 EXCELが開いたままになっている不具合 -coment：エラーの場合の処理(ログ出力)がない？
                'エラー処理を行うこと
                Debug.WriteLine(ex.Message)
                rtn = False
            End Try

            Return rtn

        End Function

        ''' <summary>
        ''' Excelファイルクローズ処理(読込時)
        ''' </summary>
        ''' <param name="appli"></param>
        ''' <param name="wbook"></param>
        ''' <param name="wsheet"></param>
        ''' <remarks></remarks>
        Public Sub Set_ExcelFile_ReadClose(ByRef appli As Excel.Application, ByRef wbook As Excel.Workbook, ByRef wsheet As Excel.Worksheet)

            Try
                wbook.Close()
                appli.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appli)      'Excelオブジェクトの開放(ロック明示的解放)
                wsheet = Nothing
                wbook = Nothing
                appli = Nothing
            Catch ex As Exception
                '20160907 EXCELが開いたままになっている不具合 -coment：エラーの場合の処理(ログ出力)がない？
                Debug.WriteLine(ex.Message)
            End Try

        End Sub

        ''' <summary>
        ''' Excelファイルクローズ処理(書込時)
        ''' </summary>
        ''' <param name="appli"></param>
        ''' <param name="wbook"></param>
        ''' <param name="wsheet"></param>
        ''' <remarks></remarks>
        Public Sub Set_ExcelFile_WriteClose(ByRef appli As Excel.Application, ByRef wbook As Excel.Workbook, ByRef wsheet As Excel.Worksheet)

            Try
                wbook.Save()
                wbook.Close()
                appli.Quit()
                '20161014 DisconnectedContextエラー修正 -chg sta
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(CType(wbook, Object))
                '20161014 DisconnectedContextエラー修正 -chg end
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appli)
                wsheet = Nothing
                wbook = Nothing
                appli = Nothing
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try

        End Sub

        ''' <summary>
        ''' Excel読込時のオープン処理 '20161004 既存中間→DB書込処理修正 -add
        ''' </summary>
        ''' <param name="filepath"></param>
        ''' <param name="sheetname"></param>
        ''' <param name="dt"></param>
        ''' <param name="con_read"></param>
        ''' <param name="cmd_read"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExcelFile_ReadOpen(ByVal filepath As String, ByVal filename As String, ByVal sql As String, _
                                           ByRef dt As DataTable, ByRef con_read As OleDbConnection) As Boolean

            Dim dr As OleDbDataReader
            Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "    'EXCEL2007以上(xlsx)
            Dim tmpextend As String = "Excel 8.0;HDR=YES;"
            Dim cmd_read As New OleDbCommand()
            Dim openfile As String = EtcMethod.Set_Path(filepath, filename & ".xlsx")

            Try

                '接続文字列生成
                con_read.ConnectionString = _
                    "Provider=" & tmpprovider & _
                    "Data Source=" & openfile & ";" & _
                    "Extended Properties=" & """" & tmpextend & """"

                '接続設定
                cmd_read.Connection = con_read

                '接続オープン処理
                con_read.Open()

                cmd_read = con_read.CreateCommand
                cmd_read.CommandText = sql
                dt = New DataTable
                dr = cmd_read.ExecuteReader
                dt.Load(dr)

            Catch ex As Exception

                Return False

            End Try

            Return True

        End Function

        ''' <summary>
        ''' Excel読込時のオープン処理 20161012 中間ファイル件数表示速度改善
        ''' データテーブル等のオブジェクトは返さずにそのまま開くだけ
        ''' </summary>
        ''' <param name="filepath"></param>
        ''' <param name="sheetname"></param>
        ''' <param name="dt"></param>
        ''' <param name="con_read"></param>
        ''' <param name="cmd_read"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExcelFile_ReadOpenOnly(ByVal filepath As String, ByVal filename As String, ByRef con_read As OleDbConnection) As Boolean


            Dim dr As OleDbDataReader
            Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "    'EXCEL2007以上(xlsx)
            Dim tmpextend As String = "Excel 8.0;HDR=YES;"
            Dim cmd_read As New OleDbCommand()
            Dim openfile As String = EtcMethod.Set_Path(filepath, filename & ".xlsx")

            Try

                '接続文字列生成
                con_read.ConnectionString = _
                    "Provider=" & tmpprovider & _
                    "Data Source=" & openfile & ";" & _
                    "Extended Properties=" & """" & tmpextend & """"

                '接続設定
                cmd_read.Connection = con_read

                '接続オープン処理
                con_read.Open()

                '20161104 中間ファイル読込エラー時の処理対応 -chg sta
                'Catch ex As Exception

                '    Return False

                'アクセスが拒否された場合
                'Catch ex As UnauthorizedAccessException
                '    MsgResult = MessageBox.Show("フォルダ作成時にアクセスが拒否されました。" & vbCrLf & _
                '                                "コンバーターのインストール先にアクセス権限を与えるか、インストール先を変更して再度セットアップを行って下さい。", _
                '                                "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '    Return False

                '無効なフォルダパスが指定された場合
            Catch ex As FileNotFoundException
                MsgResult = MessageBox.Show("無効なファイルパスが指定されました。" & vbCrLf & _
                                            "関連ファイルが削除された可能性があります。再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False

                'その他(アクセスが拒否された場合を含む)
            Catch ex As Exception
                MsgResult = MessageBox.Show(filename & "へのアクセスが拒否されました。" & vbCrLf & _
                                            filename & "の読み取り専用チェックを外すか、セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
                '20161104 中間ファイル読込エラー時の処理対応 -chg end
            End Try

            Return True

        End Function

        ''' <summary>
        ''' Excel読込時のクローズ処理 '20161004 既存中間→DB書込処理修正 -add
        ''' </summary>
        ''' <param name="con_read"></param>
        ''' <remarks></remarks>
        Public Sub ExcelFile_ReadClose(ByRef con_read As OleDbConnection)
            con_read.Close()
        End Sub

    End Class

#End Region

#Region "住所変換関連クラス"

    Public Class AddressChange

        ''' <summary>
        ''' 住所分割＋チェック処理
        ''' </summary>
        ''' <param name="chkstr">指定住所</param>
        ''' <param name="sqlcnn">接続情報</param>
        ''' <returns>分割住所情報</returns>
        ''' <remarks>
        ''' ・都道府県市区町村丁番地その他情報分割
        ''' ・例：chkstr = "宮崎県都城市"
        ''' </remarks>
        Public Shared Function Get_Address( _
                                            ByVal chkstr As String, _
                                            ByRef sqlcnn As System.Data.SqlClient.SqlConnection _
                                            ) As Hashtable

            Dim readtbl As New DataTable                                    'DataTableオブジェクト生成
            Dim fldname As String                                           '項目名
            Dim fldvalu As String                                           '項目値
            Dim rtnhash As New Hashtable                                    '返却用抽出データ格納変数
            Dim tmphash As New Hashtable                                    '作業用抽出データ格納変数

            '都道府県市情報抽出クエリ
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT (SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) AS ken_name,* "
            tmp_sql = tmp_sql & " FROM m_si "
            tmp_sql = tmp_sql & " WHERE '"
            tmp_sql = tmp_sql & chkstr
            tmp_sql = tmp_sql & "' LIKE ((SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) + si_name + '%') "

            '件数取得
            Dim reccnt As Integer = 0
            reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnn, readtbl)

            '20160622 住所分割処理対応 -chg sta
            '既存処理は都道府県がヒットしない場合全てその他へ移行していたが分割処理自体は行うように修正
            'If reccnt <> 0 Then        '抽出件数あり
            '    '抽出データありの場合(都道府県市情報あり)
            '    For cntii = 0 To readtbl.Columns.Count - 1
            '        '項目名取得
            '        fldname = readtbl.Columns(cntii).ColumnName
            '        '登録値取得
            '        fldvalu = N3Lib.Utys.Typ.ToStr(readtbl.Rows(0).Item(cntii))
            '        'データ格納
            '        rtnhash.Add(fldname, fldvalu)
            '    Next
            '    tmphash = Get_CyoBanti(chkstr.Replace(rtnhash("ken_name"), "").Replace(rtnhash("si_name"), ""))
            '    rtnhash.Add("mati", tmphash("mati"))
            '    rtnhash.Add("cyome", tmphash("cyome"))
            '    rtnhash.Add("banti", tmphash("banti"))
            '    rtnhash.Add("etc", tmphash("etc"))
            '    rtnhash.Add("chomeptn", tmphash("chomeptn"))
            'Else
            '    '抽出データなしの場合
            '    rtnhash.Add("etc", chkstr)
            '    'ログ出力処理

            'End If
            Dim tmp_addresstotal As String = ""
            Dim tmp_mati As String = ""
            Dim tmp_address As String = ""

            If reccnt <> 0 Then        '抽出件数あり
                '抽出データありの場合(都道府県市情報あり)
                For cntii = 0 To readtbl.Columns.Count - 1
                    '項目名取得
                    fldname = readtbl.Columns(cntii).ColumnName
                    '登録値取得
                    fldvalu = N3Lib.Utys.Typ.ToStr(readtbl.Rows(0).Item(cntii))
                    'データ格納
                    rtnhash.Add(fldname, fldvalu)
                Next

                tmp_addresstotal = chkstr.Replace(rtnhash("ken_name"), "").Replace(rtnhash("si_name"), "")
            Else
                tmp_addresstotal = chkstr
            End If

            tmp_mati = AddressChange.Get_Cyotiiki(tmp_addresstotal)
            If tmp_mati <> "" Then
                tmp_address = tmp_addresstotal.Replace(tmp_mati, "")
            Else
                tmp_address = tmp_addresstotal
            End If
            tmphash = AddressChange.Get_CyoBanti(tmp_address)
            rtnhash.Add("mati", tmp_mati)
            rtnhash.Add("cyome", tmphash("cyome"))
            rtnhash.Add("banti", tmphash("banti"))
            rtnhash.Add("etc", tmphash("etc"))
            rtnhash.Add("chomeptn", tmphash("chomeptn"))
            '20160622 住所分割処理対応 -chg end

            Return rtnhash

        End Function

        '20160622 住所分割処理対応 -del sta
        ' ''' <summary>
        ' ''' 町村丁番地その他情報取得
        ' ''' </summary>
        ' ''' <param name="chkstr">都道府県市区除去済み住所データ</param>
        ' ''' <returns>町丁番地その他情報</returns>
        ' ''' <remarks>
        ' ''' ・全角文字を半角文字へ変換して処理<br/>
        ' ''' 
        ' ''' ※例(テスト用) 
        ' '''     chkstr = "宮崎県都城市妻ヶ丘1-2"
        ' '''     chkstr = "鹿児島県鹿児島市吉野町7815"
        ' '''     chkstr = "熊本県熊本区１２丁目３番地４５６　1aiueo34"
        ' '''     chkstr = "熊本県熊本区１-4-9　1aiueo34"
        ' '''     chkstr = "北海道古宇郡神恵内村aaa 12-33a"
        ' '''     chkstr = 2345
        ' '''     chkstr = "Flat D, 73/F, Block 2, The Arch,"
        ' '''     chkstr = "テスト3丁目-2-3"
        ' ''' </remarks>
        'Public Shared Function Get_CyoBanti(ByVal chkstr As String) As Hashtable

        '    Dim idxban As Integer                                                   '検索文字位置
        '    Dim tmpAddrMati As String = ""                                          '丁番地以前格納
        '    Dim tmpAddrKari As String = ""                                          '丁番地以降格納
        '    Dim tmpAddrChome As String = ""                                         '丁目
        '    Dim tmpAddrBan As String = ""                                           '番地
        '    Dim tmpAddrEtc As String = ""                                           'その他
        '    Dim tmpChomePtn As Integer = -1                                         '丁目パターン(固定)
        '    Dim chkflg As Boolean = False                                           '番地チェック：True.該当あり,False.該当なし

        '    '全角文字を半角へ変換
        '    chkstr = StrConv(chkstr, VbStrConv.Narrow)

        '    '番地チェック
        '    For chkLen As Integer = 1 To Len(chkstr)
        '        '1文字ずつチェック
        '        Dim chklenstr As String = Mid(chkstr, chkLen, 1)

        '        If IsNumeric(chklenstr) Then
        '            '数値の場合(丁番地以降)
        '            If chkLen > 1 Then
        '                tmpAddrMati = N3Lib.Utys.Substr(chkstr, 0, chkLen - 1)
        '                tmpAddrKari = chkstr.Replace(tmpAddrMati, "")
        '            Else
        '                tmpAddrKari = chkstr
        '            End If

        '            '町丁番地(配列(0))
        '            Dim tmpAddrTemp1 As Array = tmpAddrKari.Replace("　", " ").Split(" ")
        '            'その他(配列(1)～)
        '            If UBound(tmpAddrTemp1) > 0 Then
        '                For idxban = 1 To UBound(tmpAddrTemp1)
        '                    tmpAddrEtc = Trim(tmpAddrEtc & " " & tmpAddrTemp1(idxban))
        '                Next
        '            End If

        '            '"丁目"位置 
        '            idxban = IIf(tmpAddrTemp1(0).IndexOf("丁目") = 0, tmpAddrTemp1(0).IndexOf("丁"), tmpAddrTemp1(0).IndexOf("丁目"))
        '            If idxban <> -1 Then
        '                '"丁目"連結文字列チェック
        '                tmpAddrChome = N3Lib.Utys.Substr(tmpAddrTemp1(0), 0, idxban)
        '                tmpAddrBan = tmpAddrTemp1(0).ToString.Replace(tmpAddrChome.ToString & "丁目", "")
        '                tmpAddrBan = tmpAddrBan.ToString.Replace(tmpAddrChome.ToString & "丁", "")
        '                tmpAddrBan = tmpAddrBan.TrimStart("-")
        '                tmpChomePtn = 1
        '            Else
        '                '"-"連結文字列チェック
        '                Dim tmpAddrTemp2 As Array = tmpAddrTemp1(0).ToString.Split("-")
        '                For cnt As Integer = 0 To UBound(tmpAddrTemp2)
        '                    Select Case cnt
        '                        Case 0
        '                            tmpAddrChome = tmpAddrTemp2(cnt)
        '                        Case Else
        '                            tmpAddrBan = tmpAddrBan & "-" & tmpAddrTemp2(cnt)
        '                            tmpAddrBan = tmpAddrBan.TrimStart("-")
        '                    End Select
        '                Next
        '            End If

        '            '処理を抜ける
        '            chkflg = True
        '            Exit For
        '        End If
        '    Next

        '    '番地チェック該当なしの場合
        '    If chkflg = False Then
        '        tmpAddrMati = chkstr
        '    End If

        '    '[丁目など]設定
        '    If tmpAddrChome.Trim <> "" Then
        '        tmpChomePtn = 1
        '    Else
        '        tmpChomePtn = -1
        '    End If

        '    '返却値
        '    Dim rtnhash As New Hashtable
        '    rtnhash.Add("mati", tmpAddrMati)
        '    rtnhash.Add("cyome", tmpAddrChome)
        '    rtnhash.Add("banti", tmpAddrBan)
        '    rtnhash.Add("etc", tmpAddrEtc)
        '    rtnhash.Add("chomeptn", tmpChomePtn)
        '    Return rtnhash

        'End Function
        '20160622 住所分割処理対応 -del end

        ''' <summary>
        ''' 文字列が連結された状態の住所から町地域のみ抽出する '20160622 住所分割処理対応
        ''' </summary>
        ''' <param name="chkstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Cyotiiki(ByVal chkstr As String) As String

            Dim rtn_str As String = ""
            Dim tmpAddrMati As String = ""

            If chkstr.Trim = "" Then
                Return rtn_str
            End If

            For cntii = 1 To Len(chkstr)

                '1文字ずつチェック
                Dim chklenstr As String = Mid(chkstr, cntii, 1)

                If IsNumeric(chklenstr) Then

                    If cntii > 1 Then
                        tmpAddrMati = N3Lib.Utys.Substr(chkstr, 0, cntii - 1)
                        rtn_str = tmpAddrMati
                        Return rtn_str
                    Else
                        '1文字目が数値の場合は町地域が無く丁目から始まっていると見なす(町地域が = "" とする)
                        Return rtn_str
                    End If

                End If

            Next

            Return rtn_str

        End Function

        ''' <summary>
        ''' 丁番地分割処理(革命V7から引用して一部編集) '20160622 住所分割処理対応
        ''' </summary>
        ''' <param name="strtotal"></param>
        ''' <param name="strcyo"></param>
        ''' <param name="strbanti"></param>
        ''' <param name="stretc"></param>
        ''' <remarks></remarks>
        Public Shared Function Get_CyoBanti(ByVal strtotal As String) As Hashtable

            Dim rtn_hash As New Hashtable
            Dim strTmpOrg As String = ""
            Dim strTmp1 As String = ""
            Dim strTmp2 As String = ""
            Dim strTmp3 As String = ""
            Dim strTmp4 As String = ""

            Dim i1 As Integer = 0
            Dim i2 As Integer = 0
            Dim i3 As Integer = 0

            Dim a1 As Integer = 0
            Dim a2 As Integer = 0
            Dim a3 As Integer = 0
            Dim a4 As Integer = 0

            Dim e1 As Integer = 0
            Dim e2 As Integer = 0

            Dim rtn As Boolean = True

            If strtotal = "" Then
                Return rtn_hash
            End If

            strTmpOrg = StrConv(strtotal, vbNarrow)

            strTmpOrg = Replace(strTmpOrg, "町目", "丁目")

            strTmpOrg = Replace(strTmpOrg, "一", "1")
            strTmpOrg = Replace(strTmpOrg, "二", "2")
            strTmpOrg = Replace(strTmpOrg, "三", "3")
            strTmpOrg = Replace(strTmpOrg, "四", "4")
            strTmpOrg = Replace(strTmpOrg, "五", "5")
            strTmpOrg = Replace(strTmpOrg, "六", "6")
            strTmpOrg = Replace(strTmpOrg, "七", "7")
            strTmpOrg = Replace(strTmpOrg, "八", "8")
            strTmpOrg = Replace(strTmpOrg, "九", "9")

            strTmpOrg = Replace(strTmpOrg, "ｰ", "-")
            strTmpOrg = Replace(strTmpOrg, "ー", "-")
            strTmpOrg = Replace(strTmpOrg, "-", "-")
            strTmpOrg = Replace(strTmpOrg, "－", "-")
            strTmpOrg = Replace(strTmpOrg, "―", "-")
            strTmpOrg = Replace(strTmpOrg, "‐", "-")

            '不正な文字が入っていないかチェック
            For cntii = 1 To Len(strTmpOrg)
                strTmp1 = Mid(strTmpOrg, cntii, 1)

                Select Case strTmp1
                    Case "", " ", "-", "丁", "目", "番", "地", "割", "号"
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    Case Else
                        strTmp4 = Mid$(strTmpOrg, cntii) 'その他住所

                        If cntii > 1 Then
                            strTmpOrg = strTmpOrg.Substring(0, cntii - 1)
                            Exit For
                        End If
                End Select
            Next

            strTmp1 = ""
            strTmp2 = ""
            strTmp3 = ""

            '「丁目」「丁」「地割」の判別
            Dim tmp_cyostr As String = ""
            Dim strcyoptn As String = ""

            If strTmpOrg <> strTmpOrg.Replace("丁目", "") Then        '「丁目」が含まれるか判定
                tmp_cyostr = "丁目"
                strcyoptn = "1"
            ElseIf strTmpOrg <> strTmpOrg.Replace("丁", "") Then      '「丁」が含まれるか判定
                tmp_cyostr = "丁"
                strcyoptn = "2"
            ElseIf strTmpOrg <> strTmpOrg.Replace("地割", "") Then    '「地割」が含まれるか判定
                tmp_cyostr = "地割"
                strcyoptn = "3"
            ElseIf strTmpOrg <> strTmpOrg.Replace("-", "") Then       '「-」が含まれるか判定
                tmp_cyostr = "-"
                strcyoptn = "1"
            Else
                tmp_cyostr = "丁目"
                strcyoptn = "-1"
            End If

            a1 = InStr(1, strTmpOrg, tmp_cyostr)
            a2 = InStr(1, strTmpOrg, "-")
            a3 = InStr(1, strTmpOrg, "番地")
            a4 = InStr(1, strTmpOrg, "番")

            '「丁目」「丁」「地割」が存在する
            '"-" が "番地" or "番" の前に存在する

            If a1 <> 0 Or (a2 <> 0 And a3 <> 0 And a4 <> 0 And (a2 < a3 And a2 < a4)) Or _
               (a2 <> 0 And a3 = 0 And a4 = 0) Then
                If a1 <> 0 Then
                    strTmp1 = strTmpOrg.Substring(0, a1 - 1)
                    strTmpOrg = Mid(strTmpOrg, a1 + Len(tmp_cyostr))
                ElseIf a2 <> 0 Then
                    strTmp1 = strTmpOrg.Substring(0, a2 - 1)
                    strTmpOrg = Mid(strTmpOrg, a2 + Len("-"))
                End If

                i1 = InStr(1, strTmpOrg, "番地")
                i3 = InStr(1, strTmpOrg, "番")
                i2 = InStr(1, strTmpOrg, "-")

                If i1 <> 0 Or i2 <> 0 Or i3 <> 0 Then
                    '20160829 住所分割処理修正 -chg sta
                    'If i1 <> 0 Then
                    '    strTmp2 = strTmpOrg.Substring(0, i1 - 1)
                    '    strTmpOrg = Mid(strTmpOrg, i1 + Len("番地"))
                    'ElseIf i2 <> 0 Then
                    '    strTmp2 = strTmpOrg.Substring(0, i2 - 1)
                    '    strTmpOrg = Mid(strTmpOrg, i2 + Len("-"))
                    'ElseIf i3 <> 0 Then
                    '    strTmp2 = strTmpOrg.Substring(0, i3 - 1)
                    '    strTmpOrg = Mid(strTmpOrg, i3 + Len("番"))
                    'End If
                    'strTmp3 = strTmpOrg
                    strTmp2 = strTmpOrg
                    '20160829 住所分割処理修正 -chg end
                    rtn = True
                Else
                    If a1 = 0 Then
                        strTmp2 = strTmp1
                        strTmp3 = strTmpOrg
                        strTmp1 = ""
                    Else
                        strTmp2 = strTmpOrg
                    End If
                    rtn = True
                End If

            Else
                i1 = InStr(1, strTmpOrg, "番地")
                i3 = InStr(1, strTmpOrg, "番")
                i2 = InStr(1, strTmpOrg, "-")

                If i1 <> 0 Or i2 <> 0 Or i3 <> 0 Then

                    If i1 <> 0 Then
                        strTmp2 = strTmpOrg.Substring(0, i1 - 1)
                        strTmp3 = Mid(strTmpOrg, i1 + Len("番地"))
                    ElseIf i2 <> 0 Then
                        strTmp2 = strTmpOrg.Substring(0, i2 - 1)
                        strTmp3 = Mid(strTmpOrg, i2 + Len("-"))
                    ElseIf i3 <> 0 Then
                        strTmp2 = strTmpOrg.Substring(0, i3 - 1)
                        strTmp3 = Mid(strTmpOrg, i3 + Len("番"))
                    End If

                    rtn = True
                Else
                    strTmp4 = strTmpOrg
                    rtn = True
                End If
            End If

            '例外
            '    If strTmp3 = "" And Val(strTmp1) > 50 Then
            '        strTmp3 = strTmp2
            '        strTmp2 = strTmp1
            '        strTmp1 = ""
            '    End If

            strTmp1 = Trim(Replace(strTmp1, tmp_cyostr, ""))
            strTmp1 = Trim(Replace(strTmp1, "番", ""))
            strTmp1 = Trim(Replace(strTmp1, "号", ""))
            strTmp1 = Trim(Replace(strTmp1, "-", ""))
            '20160829 住所分割処理修正 -del sta
            'strTmp2 = Trim(Replace(strTmp2, "番", ""))
            'strTmp2 = Trim(Replace(strTmp2, tmp_cyostr, ""))
            'strTmp2 = Trim(Replace(strTmp2, "号", ""))
            'strTmp2 = Trim(Replace(strTmp2, "-", ""))
            '20160829 住所分割処理修正 -del end
            Dim strcyo As String = strTmp1
            Dim strbanti As String = strTmp2
            '20160829 住所分割処理修正 -chg sta
            'Dim stretc As String = (strTmp3 & " " & strTmp4).Trim
            Dim stretc As String = strTmp4.Trim
            '20160829 住所分割処理修正 -chg end
            rtn_hash.Add("cyome", strcyo)
            rtn_hash.Add("banti", strbanti)
            rtn_hash.Add("etc", stretc)
            rtn_hash.Add("chomeptn", strcyoptn)

            Return rtn_hash

        End Function

    End Class

    Public Class AddressChange_old

        ''' <summary>
        ''' 住所分割＋チェック処理
        ''' </summary>
        ''' <param name="chkstr">指定住所</param>
        ''' <param name="sqlcnn">接続情報</param>
        ''' <returns>分割住所情報</returns>
        ''' <remarks>
        ''' ・都道府県市区町村丁番地その他情報分割
        ''' ・例：chkstr = "宮崎県都城市"
        ''' </remarks>
        Public Shared Function Get_Address( _
                                            ByVal chkstr As String, _
                                            ByRef sqlcnn As System.Data.SqlClient.SqlConnection _
                                            ) As Hashtable

            Dim readtbl As New DataTable                                    'DataTableオブジェクト生成
            Dim fldname As String                                           '項目名
            Dim fldvalu As String                                           '項目値
            Dim rtnhash As New Hashtable                                    '返却用抽出データ格納変数
            Dim tmphash As New Hashtable                                    '作業用抽出データ格納変数

            '都道府県市情報抽出クエリ
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT (SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) AS ken_name,* "
            tmp_sql = tmp_sql & " FROM m_si "
            tmp_sql = tmp_sql & " WHERE '"
            tmp_sql = tmp_sql & chkstr
            tmp_sql = tmp_sql & "' LIKE ((SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) + si_name + '%') "

            '件数取得
            Dim reccnt As Integer = 0
            reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnn, readtbl)

            If reccnt <> 0 Then        '抽出件数あり
                '抽出データありの場合(都道府県市情報あり)
                For cntii = 0 To readtbl.Columns.Count - 1
                    '項目名取得
                    fldname = readtbl.Columns(cntii).ColumnName
                    '登録値取得
                    fldvalu = N3Lib.Utys.Typ.ToStr(readtbl.Rows(0).Item(cntii))
                    'データ格納
                    rtnhash.Add(fldname, fldvalu)
                Next
                tmphash = Get_CyoBanti(chkstr.Replace(rtnhash("ken_name"), "").Replace(rtnhash("si_name"), ""))
                rtnhash.Add("mati", tmphash("mati"))
                rtnhash.Add("cyome", tmphash("cyome"))
                rtnhash.Add("banti", tmphash("banti"))
                rtnhash.Add("etc", tmphash("etc"))
                rtnhash.Add("chomeptn", tmphash("chomeptn"))
            Else
                '抽出データなしの場合
                rtnhash.Add("etc", chkstr)
                'ログ出力処理

            End If

            Return rtnhash

        End Function

        ''' <summary>
        ''' 町村丁番地その他情報取得
        ''' </summary>
        ''' <param name="chkstr">都道府県市区除去済み住所データ</param>
        ''' <returns>町丁番地その他情報</returns>
        ''' <remarks>
        ''' ・全角文字を半角文字へ変換して処理<br/>
        ''' 
        ''' ※例(テスト用) 
        '''     chkstr = "宮崎県都城市妻ヶ丘1-2"
        '''     chkstr = "鹿児島県鹿児島市吉野町7815"
        '''     chkstr = "熊本県熊本区１２丁目３番地４５６　1aiueo34"
        '''     chkstr = "熊本県熊本区１-4-9　1aiueo34"
        '''     chkstr = "北海道古宇郡神恵内村aaa 12-33a"
        '''     chkstr = 2345
        '''     chkstr = "Flat D, 73/F, Block 2, The Arch,"
        '''     chkstr = "テスト3丁目-2-3"
        ''' </remarks>
        Public Shared Function Get_CyoBanti(ByVal chkstr As String) As Hashtable

            Dim idxban As Integer                                                   '検索文字位置
            Dim tmpAddrMati As String = ""                                          '丁番地以前格納
            Dim tmpAddrKari As String = ""                                          '丁番地以降格納
            Dim tmpAddrChome As String = ""                                         '丁目
            Dim tmpAddrBan As String = ""                                           '番地
            Dim tmpAddrEtc As String = ""                                           'その他
            Dim tmpChomePtn As Integer = -1                                         '丁目パターン(固定)
            Dim chkflg As Boolean = False                                           '番地チェック：True.該当あり,False.該当なし

            '全角文字を半角へ変換
            chkstr = StrConv(chkstr, VbStrConv.Narrow)

            '番地チェック
            For chkLen As Integer = 1 To Len(chkstr)
                '1文字ずつチェック
                Dim chklenstr As String = Mid(chkstr, chkLen, 1)

                If IsNumeric(chklenstr) Then
                    '数値の場合(丁番地以降)
                    If chkLen > 1 Then
                        tmpAddrMati = N3Lib.Utys.Substr(chkstr, 0, chkLen - 1)
                        tmpAddrKari = chkstr.Replace(tmpAddrMati, "")
                    Else
                        tmpAddrKari = chkstr
                    End If

                    '町丁番地(配列(0))
                    Dim tmpAddrTemp1 As Array = tmpAddrKari.Replace("　", " ").Split(" ")
                    'その他(配列(1)～)
                    If UBound(tmpAddrTemp1) > 0 Then
                        For idxban = 1 To UBound(tmpAddrTemp1)
                            tmpAddrEtc = Trim(tmpAddrEtc & " " & tmpAddrTemp1(idxban))
                        Next
                    End If

                    '"丁目"位置 
                    idxban = IIf(tmpAddrTemp1(0).IndexOf("丁目") = 0, tmpAddrTemp1(0).IndexOf("丁"), tmpAddrTemp1(0).IndexOf("丁目"))
                    If idxban <> -1 Then
                        '"丁目"連結文字列チェック
                        tmpAddrChome = N3Lib.Utys.Substr(tmpAddrTemp1(0), 0, idxban)
                        tmpAddrBan = tmpAddrTemp1(0).ToString.Replace(tmpAddrChome.ToString & "丁目", "")
                        tmpAddrBan = tmpAddrBan.ToString.Replace(tmpAddrChome.ToString & "丁", "")
                        tmpAddrBan = tmpAddrBan.TrimStart("-")
                        tmpChomePtn = 1
                    Else
                        '"-"連結文字列チェック
                        Dim tmpAddrTemp2 As Array = tmpAddrTemp1(0).ToString.Split("-")
                        For cnt As Integer = 0 To UBound(tmpAddrTemp2)
                            Select Case cnt
                                Case 0
                                    tmpAddrChome = tmpAddrTemp2(cnt)
                                Case Else
                                    tmpAddrBan = tmpAddrBan & "-" & tmpAddrTemp2(cnt)
                                    tmpAddrBan = tmpAddrBan.TrimStart("-")
                            End Select
                        Next
                    End If

                    '処理を抜ける
                    chkflg = True
                    Exit For
                End If
            Next

            '番地チェック該当なしの場合
            If chkflg = False Then
                tmpAddrMati = chkstr
            End If

            '[丁目など]設定
            If tmpAddrChome.Trim <> "" Then
                tmpChomePtn = 1
            Else
                tmpChomePtn = -1
            End If

            '返却値
            Dim rtnhash As New Hashtable
            rtnhash.Add("mati", tmpAddrMati)
            rtnhash.Add("cyome", tmpAddrChome)
            rtnhash.Add("banti", tmpAddrBan)
            rtnhash.Add("etc", tmpAddrEtc)
            rtnhash.Add("chomeptn", tmpChomePtn)
            Return rtnhash

        End Function

    End Class


#End Region

#Region "革命10DBフィールドタイプ格納用ハッシュテーブル作成クラス"

    Public Class GetFieldInfo

        ''' <summary>
        ''' 革命10DBフィールド情報(サイズ/型)を取得してハッシュテーブルへ格納 接続先不正によるエラー時の処理追加 (Sub→Functionに変更し、正常終了の判別を行う)
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="strsql"></param>
        ''' <param name="hash_type"></param>
        ''' <param name="hash_size"></param>
        ''' <remarks></remarks>
        Public Shared Function Get_HashFieldInfo(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection, ByVal strsql As String) As Boolean

            Dim readtbl As New DataTable
            Dim reccnt As Integer
            Dim fldname As String
            Dim fldvalue As String
            Dim tmp_tblname As String = ""
            Dim tmp_tblname_cv As String = ""
            Dim tmp_fldname As String = ""
            Dim tmp_fldname_cv As String = ""
            Dim tmp_type As String = ""
            Dim tmp_size As String = ""
            Dim tmp_size_cv_max As String = ""
            Dim tmp_size_cv_min As String = ""
            Dim tmp_defaultvalue As String = ""
            Dim tmp_keyno As String = ""
            Dim tmp_keyflg As String = ""
            Dim tmp_cvkeyno As String = ""
            Dim tmp_requiredflg As String = ""
            Dim tmp_mstreference As String = ""
            '2016.03.28 接続先不正によるエラー時の処理追加 -add
            Dim rtn As Boolean = True

            '--------------------------------
            '革命10フィールド情報取得
            '--------------------------------
            reccnt = DBExec.Exec_DataTable(strsql, sqlcnnv10, readtbl)

            '2016.03.28 接続先不正によるエラー時の処理追加 -add sta
            'DB情報取得件数→接続情報が不正の可能性があるため処理を抜ける
            If reccnt = 0 Then
                rtn = False
                Return rtn
            End If
            '2016.03.28 接続先不正によるエラー時の処理追加 -add end

            '2016.02.22 中間ファイルヘッダーエラー修正が生じる現象の対応 -add sta
            '--------------------------------
            '初期化
            '--------------------------------
            Hash_FiledTypeAlpha.Clear()
            Hash_FiledTypeJp.Clear()
            Hash_FiledSizeAlpha.Clear()
            Hash_FiledSizeJp.Clear()
            Hash_Min_Code.Clear()
            Hash_Max_Code.Clear()
            Hash_TblName_AlphaToJp.Clear()
            Hash_TblName_JpToAlpha.Clear()
            Hash_FldName_AlphaToJp.Clear()
            Hash_FldName_JpToAlpha.Clear()
            Hash_DefaultValue.Clear()
            Hash_MidKey_FieldAlpha.Clear()
            Hash_MidKey_FieldJp.Clear()
            Hash_Mst_ReferenceAlpha.Clear()
            Hash_FldName_JpToAlpha_Mid.Clear()  '20160812 汎用コンバート対応 -add
            '2016.02.22 中間ファイルヘッダーエラー修正が生じる現象の対応 -add end

            '--------------------------------
            '読込開始
            '--------------------------------
            For cntii As Integer = 0 To readtbl.Rows.Count - 1

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Exit Function
                End If

                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                    '項目名取得
                    fldname = readtbl.Columns(cntjj).ColumnName

                    '登録値取得
                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                    '各項目値→変数格納
                    Select Case fldname
                        Case "行No"
                            tmp_keyno = fldvalue.Trim
                        Case "主キー"
                            tmp_keyflg = fldvalue.Trim
                        Case "TBL名"
                            tmp_tblname = fldvalue.Trim
                        Case "フィールド名"
                            tmp_fldname = fldvalue.Trim
                        Case "属性"
                            tmp_type = fldvalue.Trim
                        Case "サイズ"
                            tmp_size = fldvalue.Trim
                        Case "CV用種別"
                            tmp_tblname_cv = fldvalue.Trim
                        Case "CV用項目名"
                            tmp_fldname_cv = fldvalue.Trim
                        Case "CV用最小値"
                            tmp_size_cv_min = fldvalue.Trim
                        Case "CV用最大値"
                            tmp_size_cv_max = fldvalue.Trim
                        Case "デフォルト値"
                            tmp_defaultvalue = fldvalue.Trim
                        Case "CV用キーNo"
                            tmp_cvkeyno = fldvalue.Trim
                        Case "CV用必須項目フラグ"
                            tmp_requiredflg = fldvalue.Trim
                        Case "有無確認区分"
                            tmp_mstreference = fldvalue.Trim
                    End Select

                Next

                '--------------------------------
                '各オブジェクトへ格納
                '--------------------------------

                '革命のフィールド名とフィールドタイプ取得
                Hash_FiledTypeAlpha.Add(tmp_tblname & "-" & tmp_fldname, tmp_type)

                'コンバート用フィールド名とフィールドタイプ取得
                If tmp_tblname_cv <> "" Then
                    Hash_FiledTypeJp.Add(tmp_tblname_cv & "-" & tmp_fldname_cv, tmp_type)
                End If

                '革命のフィールド名とフィールドサイズ取得
                Hash_FiledSizeAlpha.Add(tmp_tblname & "-" & tmp_fldname, tmp_size)

                'コンバート用フィールド名とフィールドサイズ取得
                If tmp_tblname_cv <> "" Then
                    Hash_FiledSizeJp.Add(tmp_tblname_cv & "-" & tmp_fldname_cv, tmp_size)       'kakaka メモ：Alpha,Jp各、サイズ等の情報を付加
                End If

                '20160812 汎用コンバート対応 -add sta
                'コンバート用フィールド名(日本語)と革命のフィールド名(アルファベット)取得
                If tmp_tblname_cv <> "" Then
                    Hash_FldName_JpToAlpha_Mid.Add(tmp_tblname_cv & "-" & tmp_fldname_cv, tmp_fldname)
                End If
                '20160812 汎用コンバート対応 -add end

                '革命上の最小値取得
                If tmp_size_cv_min <> "" Then
                    Hash_Min_Code.Add(tmp_tblname & "-" & tmp_fldname, tmp_size_cv_min)
                End If

                '革命上の最大値取得
                If tmp_size_cv_max <> "" Then
                    Hash_Max_Code.Add(tmp_tblname & "-" & tmp_fldname, tmp_size_cv_max)
                End If

                '革命のテーブル名(アルファベット)とコンバート用テーブル名(日本語)取得
                If tmp_tblname <> "" And Not Hash_TblName_AlphaToJp.Contains(tmp_tblname) Then
                    Hash_TblName_AlphaToJp.Add(tmp_tblname, tmp_tblname_cv)
                End If

                'コンバート用テーブル名(日本語)と革命のテーブル名(アルファベット)取得
                If tmp_tblname_cv <> "" And Not Hash_TblName_JpToAlpha.Contains(tmp_tblname_cv) Then
                    Hash_TblName_JpToAlpha.Add(tmp_tblname_cv, tmp_tblname)
                End If

                '革命のフィールド名(アルファベット)とコンバート用フィールド名(日本語)取得
                If tmp_tblname_cv <> "" Then
                    Hash_FldName_AlphaToJp.Add(tmp_tblname & "-" & tmp_fldname, tmp_fldname_cv)
                End If

                'コンバート用フィールド名(日本語)と革命のフィールド名(アルファベット)取得
                If tmp_tblname_cv <> "" Then
                    Hash_FldName_JpToAlpha.Add(tmp_tblname_cv & "-" & tmp_fldname_cv, tmp_tblname & "-" & tmp_fldname)
                End If

                '革命のフィールド名とデフォルト値取得
                If tmp_defaultvalue <> "" Then
                    Hash_DefaultValue.Add(tmp_tblname & "-" & tmp_fldname, tmp_defaultvalue)
                End If

                '移行元のキーNoを革命のフィールド名で取得
                If tmp_cvkeyno <> "" Then
                    Hash_MidKey_FieldAlpha.Add(tmp_tblname & "-" & tmp_fldname, tmp_cvkeyno)
                End If

                '移行元のキーNoをコンバート用フィールド名で取得
                If tmp_cvkeyno <> "" Then
                    Hash_MidKey_FieldJp.Add(tmp_tblname_cv & "-" & tmp_fldname_cv, tmp_cvkeyno)
                End If

                'キーではないが必須になっている項目を取得(物件名称等)
                If tmp_requiredflg <> "" Then
                    List_Required_Field.Add(tmp_tblname & "-" & tmp_fldname)
                End If

                'コンバート用フィールド名(日本語)を取得
                If tmp_tblname_cv <> "" Then
                    List_FldNameJp.Add(tmp_tblname_cv & "-" & tmp_fldname_cv)
                End If

                '有無確認区分を取得
                If tmp_mstreference <> "" Then
                    Hash_Mst_ReferenceAlpha.Add(tmp_tblname & "-" & tmp_fldname, tmp_mstreference)
                End If

            Next

            '2016.03.28 接続先不正によるエラー時の処理追加 -add
            Return rtn

        End Function

    End Class

#End Region

#Region "金融機関マスタ格納用ハッシュテーブル作成クラス"

    Public Class GetBankInfo

        ''' <summary>
        ''' 金融機関/金融機関支店情報を取得してハッシュテーブルへ格納
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="tmp_sql"></param>
        ''' <remarks></remarks>
        Public Shared Sub Get_HashKinyu(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection, ByVal tmp_sql As String)

            Dim readtbl As New DataTable
            Dim reccnt As Integer
            Dim fldname As String
            Dim fldvalue As String
            Dim tmp_kinyuno As String = ""
            Dim tmp_kinyuname As String = ""
            Dim tmp_tenno As String = ""
            Dim tmp_tenname As String = ""

            '金融機関情報取得
            reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl)

            '読込開始
            For cntii As Integer = 0 To readtbl.Rows.Count - 1

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Exit Sub
                End If

                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                    '項目名取得
                    fldname = readtbl.Columns(cntjj).ColumnName

                    '登録値取得
                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                    '各項目値→変数格納
                    Select Case fldname
                        Case "kinyu_no"
                            tmp_kinyuno = fldvalue
                        Case "kinyu_name"
                            tmp_kinyuname = fldvalue
                        Case "kinyu_tenno"
                            tmp_tenno = fldvalue
                        Case "kinyu_tenname"
                            tmp_tenname = fldvalue
                    End Select

                Next

                'ハッシュテーブルへ格納
                If Not Hash_Kinyu.Contains(tmp_kinyuno) Then
                    Hash_Kinyu.Add(tmp_kinyuno, tmp_kinyuname)
                End If
                Hash_KinyuTen.Add(tmp_kinyuno & "-" & tmp_tenno, tmp_kinyuname & "-" & tmp_tenname)

            Next

        End Sub

    End Class

#End Region

#Region "フィールド/移行値紐付けハッシュテーブル作成クラス"

    Public Class GetHashFldToValue

        ''' <summary>
        ''' フィールド名と移行値を紐付けてハッシュテーブルへ格納
        ''' </summary>
        ''' <param name="fldnamegrp"></param>
        ''' <param name="model_cvitem"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Hash_fldvalue(ByVal fldnamegrp As String, ByVal model_cvitem As Object) As Hashtable

            Dim rtn_hash As New Hashtable
            Dim modeltype As Type = model_cvitem.GetType
            Dim fldname() As String = fldnamegrp.Split(",")

            For cntii = 0 To UBound(fldname)

                Dim tmp_variname As String = "Vari_" & fldname(cntii).Substring(0, 1).ToUpper & fldname(cntii).Substring(1)
                Dim propertyinfo As Object = modeltype.GetProperty(tmp_variname)
                Dim targetproperty As Object = propertyinfo.getValue(model_cvitem, Nothing)
                Dim targetvalue As String = targetproperty

                rtn_hash.Add(fldname(cntii), targetvalue)

            Next

            Return rtn_hash

        End Function

        ''' <summary>
        ''' フィールド名と移行値を紐付けてハッシュテーブルへ格納
        ''' 備考等、複数行挿入する場合
        ''' </summary>
        ''' <param name="fldnamegrp"></param>
        ''' <param name="model_cvitem"></param>
        ''' <param name="rowindex"></param>
        ''' <param name="colindex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Hash_fldvalue(ByVal fldnamegrp As String, ByVal model_cvitem As Object, ByVal rowindex As Integer, ByVal colindex As Integer) As Hashtable

            Dim rtn_hash As New Hashtable
            Dim modeltype As Type = model_cvitem.GetType
            Dim fldname() As String = fldnamegrp.Split(",")

            For cntii = 0 To UBound(fldname)

                Dim tmp_variname As String = "Vari_" & fldname(cntii).Substring(0, 1).ToUpper & fldname(cntii).Substring(1) & colindex.ToString
                Dim propertyinfo As Object = modeltype.GetProperty(tmp_variname)
                Dim targetproperty As Object = propertyinfo.getValue(model_cvitem, Nothing)
                Dim targetvalue As String = targetproperty(rowindex)

                rtn_hash.Add(fldname(cntii), targetvalue)

            Next

            Return rtn_hash

        End Function

        ''' <summary>
        ''' フィールド名と移行値を紐付けてハッシュテーブルへ格納(配列同士を紐付け)
        ''' </summary>
        ''' <param name="fldname"></param>
        ''' <param name="fldvalue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Hash_fldvalue(ByVal tblname As String, ByVal fldname As Object, ByVal fldvalue As Object) As Hashtable

            Dim rtn_hash As New Hashtable

            For cntii = 0 To UBound(fldname)

                '2016.02.22 セル取得方法修正 -chg sta
                'rtn_hash.Add(tblname & "-" & fldname(cntii), fldvalue(cntii))
                rtn_hash.Add(tblname & "-" & fldname(1, cntii + 1), fldvalue(1, cntii + 1))
                '2016.02.22 セル取得方法修正 -chg end
            Next

            Return rtn_hash

        End Function

        ''' <summary>
        ''' Xpathと移行値を紐付けてハッシュテーブルへ格納 '20160616 初期設定情報移行処理実装
        ''' </summary>
        ''' <param name="fldnamegrp"></param>
        ''' <param name="model_cvitem"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Hash_fldvalue(ByVal hash_varitopath As Hashtable, ByVal model_cvitem As Object) As Hashtable

            Dim rtn_hash As New Hashtable
            Dim modeltype As Type = model_cvitem.GetType
            'Dim fldname() As String = fldnamegrp.Split(",")

            'modeltype.GetProperty()

            'For cntii = 0 To UBound(fldname)

            '    Dim tmp_variname As String = "Vari_" & fldname(cntii).Substring(0, 1).ToUpper & fldname(cntii).Substring(1)
            '    Dim propertyinfo As Object = modeltype.GetProperty(tmp_variname)
            '    Dim targetproperty As Object = propertyinfo.getValue(model_cvitem, Nothing)
            '    Dim targetvalue As String = targetproperty

            '    rtn_hash.Add(fldname(cntii), targetvalue)

            'Next

            Return rtn_hash

        End Function

        ''' <summary>
        ''' 既存用中間ファイルのフィールド名と移行値を紐付けてハッシュテーブルへ格納 '20160812 汎用コンバート対応 -add
        ''' </summary>
        ''' <param name="fldname"></param>
        ''' <param name="fldvalue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Hash_fldvalue_MidheaderToValue(ByVal tblname As String, ByVal fldname As Object, ByVal fldvalue As Object) As Hashtable

            Dim rtn_hash As New Hashtable

            For cntii = 0 To fldname.Length - 1
                Dim midfldname As String = fldname(1, cntii + 1)
                rtn_hash.Add(midfldname, fldvalue(1, cntii + 1))
            Next

            Return rtn_hash

        End Function

        ''' <summary>
        ''' 汎用用中間ファイルのフィールド名と移行値を紐付けてハッシュテーブルへ格納(部屋設備用) '20160913_2 部屋設備移行処理の追加 -add
        ''' </summary>
        ''' <param name="fldname"></param>
        ''' <param name="fldvalue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Hash_fldvalue_MidheaderToValue_Setubi(ByVal tblname As String, ByVal fldname As Object, ByVal fldvalue As Object) As Hashtable

            Dim rtn_hash As New Hashtable

            For cntii = 1 To fldname.Length - 1
                Dim midfldname As String = fldname(0, cntii)
                rtn_hash.Add(midfldname, fldvalue(1, cntii))
            Next

            Return rtn_hash

        End Function

    End Class

#End Region

#Region "ログ出力設定クラス"

    Public Class LogSetting

        ''' <summary>
        ''' ログ出力テーブルCREATEクエリ生成
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LogTblCreateQry(ByVal midchkflg As Boolean) As String

            Dim tmp_sql As String = ""
            Dim taisyotbl As String = ""

            If midchkflg Then
                taisyotbl = LOG_TMP_MIDCHKTABLENAME
            Else
                taisyotbl = LOG_TMP_TABLENAME
            End If

            tmp_sql = tmp_sql & " CREATE TABLE " & taisyotbl
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & LOG_HEADER1 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER2 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER3 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER4 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER5 & " NCHAR(200), "               '2016.03.28 許容文字数を拡張
            tmp_sql = tmp_sql & LOG_HEADER6 & " NCHAR(100), "                'kakaka 50で足りる？足りればOK
            tmp_sql = tmp_sql & LOG_HEADER7 & " NCHAR(100), "
            tmp_sql = tmp_sql & LOG_HEADER8 & " NCHAR(200), "                'kakaka 50で足りる？足りればOK
            tmp_sql = tmp_sql & LOG_HEADER9 & " NCHAR(200), "                'kakaka 50で足りる？足りればOK
            tmp_sql = tmp_sql & LOG_HEADER10 & " NCHAR(200), "
            tmp_sql = tmp_sql & LOG_HEADER11 & " NCHAR(200), "
            tmp_sql = tmp_sql & LOG_HEADER12 & " NCHAR(50), "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	); "

            Return tmp_sql

        End Function

        ''' <summary>
        ''' ログ出力テーブルINSERTクエリ生成
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LogTblInsertQry(ByVal logvalue As String, ByVal midchkflg As Boolean) As String

            Dim tmp_sql As String = ""
            Dim taisyotbl As String = ""

            If midchkflg Then
                taisyotbl = LOG_TMP_MIDCHKTABLENAME
            Else
                taisyotbl = LOG_TMP_TABLENAME
            End If

            tmp_sql = " INSERT INTO " & taisyotbl & " ( " & LOG_HEADER_TOTAL & ") VALUES ("
            tmp_sql = tmp_sql & logvalue
            tmp_sql = tmp_sql & ");"

            Return tmp_sql

        End Function

        ''' <summary>
        ''' ログ出力テーブルSELECTクエリ作成
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LogTblSelectQry(ByVal midchkflg As Boolean) As String

            Dim tmp_sql As String = ""
            Dim taisyotbl As String = ""

            If midchkflg Then
                taisyotbl = LOG_TMP_MIDCHKTABLENAME
            Else
                taisyotbl = LOG_TMP_TABLENAME
            End If

            tmp_sql = tmp_sql & " SELECT "
            tmp_sql = tmp_sql & "     RTRIM(" & LOG_HEADER1 & ") AS " & LOG_HEADER1
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER2 & ") AS " & LOG_HEADER2
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER3 & ") AS " & LOG_HEADER3
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER4 & ") AS " & LOG_HEADER4
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER5 & ") AS " & LOG_HEADER5
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER6 & ") AS " & LOG_HEADER6
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER7 & ") AS " & LOG_HEADER7
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER8 & ") AS " & LOG_HEADER8
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER9 & ") AS " & LOG_HEADER9
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER10 & ") AS " & LOG_HEADER10
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER11 & ") AS " & LOG_HEADER11
            If Dev_CVFlg Then                                                                       'kakaka メモ：TBL名は通常非表示にする
                tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER12 & ") AS " & LOG_HEADER12
            End If
            tmp_sql = tmp_sql & " FROM " & taisyotbl
            tmp_sql = tmp_sql & " ORDER BY " & LOG_HEADER1

            Return tmp_sql

        End Function

        ''' <summary>
        ''' ログに出力する値を設定
        ''' </summary>
        ''' <param name="typeno"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Set_LogValue(ByVal typeno As Integer, _
                                            Optional ByVal syorikomok As String = "-", _
                                            Optional ByVal syoriitem As String = "-", _
                                            Optional ByVal syorikekka As String = "-", _
                                            Optional ByVal taisyodata As String = "-", _
                                            Optional ByVal hubigein As String = "-", _
                                            Optional ByVal taisyo As String = "-", _
                                            Optional ByVal beforechgvalue As String = "-", _
                                            Optional ByVal afterchgvalue As String = "-", _
                                            Optional ByVal tblname As String = "-") As String

            Dim itemcnt As Integer = 12
            Dim tmp_logitem(itemcnt - 1) As String
            Dim tmp_str As String = ""
            Dim rtn_logvalue As String = ""
            '20161108_2 レビュー結果戻り修正時に気付いた箇所の修正 -chg sta
            '余計な「/」が入っていたため除去
            '日付と時間の間に空白があるとCSVファイルをExcelで開いた際に日付の表示形式が
            '不正になるので「_」を入れる
            ''20161025 ログ出力日時をミリ秒まで拡張する -chg sta
            ''tmp_logitem(0) = String.Format(Now)
            'tmp_logitem(0) = DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss fff")
            ''20161025 ログ出力日時をミリ秒まで拡張する -chg end
            tmp_logitem(0) = DateTime.Now.ToString("yyyy/MM/dd_HH:mm:ss fff")
            '20161108_2 レビュー結果戻り修正時に気付いた箇所の修正 -chg end
            tmp_logitem(3) = syorikomok
            tmp_logitem(4) = syoriitem
            tmp_logitem(5) = syorikekka
            '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
            'tmp_logitem(6) = taisyodata
            tmp_logitem(6) = taisyodata.Replace(",", "，").Replace("'", "’")
            '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
            tmp_logitem(7) = hubigein
            tmp_logitem(8) = taisyo
            '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
            'tmp_logitem(9) = beforechgvalue
            'tmp_logitem(10) = afterchgvalue
            tmp_logitem(9) = beforechgvalue.Replace(",", "，").Replace("'", "’")
            tmp_logitem(10) = afterchgvalue.Replace(",", "，").Replace("'", "’")
            '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
            tmp_logitem(11) = tblname

            Select Case typeno
                Case 1  '処理開始
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYU_STA
                Case 9  '処理終了
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYU_END

                Case 11 '移行元読込時
                    tmp_logitem(1) = ""
                    tmp_logitem(2) = ""

                Case 21 '中間ファイル作成時
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYU_MIDMAKE

                Case 31 '中間ファイル読込時
                    tmp_logitem(1) = ""
                    tmp_logitem(2) = ""

                Case 41 'DB書込時エラー
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYU_DBWRITE

                Case 42 'DB書込時エラー
                    tmp_logitem(1) = LOG_RUI_TYUUI
                    tmp_logitem(2) = LOG_SYU_DBWRITE

                Case 99 '中断
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYU_STOP

                Case 101    '中間ファイルヘッダーチェック
                    tmp_logitem(1) = LOG_RUI_KEIKOKU
                    tmp_logitem(2) = LOG_SYU_MIDHEADERCHK
                Case 102    '中間ファイルデータチェック
                    tmp_logitem(1) = LOG_RUI_TYUUI
                    tmp_logitem(2) = LOG_SYU_MIDDATACHK
                    
                Case 901    '汎用→既存U用中間ファイルコピー                   '20161009 ログ出力処理追加 -add
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYORIKOMK_MIDCOPY
                Case 905    '汎用→既存U用中間ファイルチェック                 '20161009 ログ出力処理追加 -add
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYORIKOMK_MIDCHECK
                Case 902    '汎用→既存U用中間ファイルデータ調整               '20161009 ログ出力処理追加 -add
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYORIKOMK_MIDCONDI
                Case 903    '汎用→既存U用中間ファイルコピー(個別)             '20161009 ログ出力処理追加 -add
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYORIKOMK_MIDONECOPY
                Case 904    '汎用→既存U用中間ファイルコピーエラー             '20161009 ログ出力処理追加 -add
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_SYORIKOMK_ERR_MIDONECOPY

                Case Else

            End Select

            For cntii = 0 To itemcnt - 1

                tmp_str = tmp_str & ",'" & tmp_logitem(cntii) & "'"

            Next

            rtn_logvalue = tmp_str.Remove(0, 1)

            Return rtn_logvalue

        End Function

        ''' <summary>
        ''' 移行時に生じたログ内容を各変数へ格納　　　　　　　　　　　　　　　'kakaka 手順・フローなどは"remarks"に記載下さい。(他の箇所についても同様)
        ''' 別メソッド(同一クラス内)にて成形                                  'kakaka 注意：現時点ではいいが、各タグ(paramなど)の内容も記載下さい。(他の箇所についても同様)
        ''' 挿入用クエリへ加工
        ''' ソートリストへ格納
        ''' </summary>
        ''' <param name="hash_log"></param>
        ''' <param name="hash_beforechk"></param>
        ''' <param name="hash_afterchk"></param>
        ''' <param name="taisyodata"></param>
        ''' <param name="tblname"></param>
        ''' <param name="logcnt"></param>
        ''' <param name="sortlist"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_Log_Value_KomkErr(ByVal hash_log As Hashtable, ByVal hash_beforechk As Hashtable, ByVal hash_afterchk As Hashtable, _
                                                ByVal taisyodata As String, ByVal tblname As String, ByRef logcnt As Integer, _
                                                ByRef sortlist As SortedList(Of Integer, String))

            For Each logvalue In hash_log

                Dim log_key As String = logvalue.Key
                Dim log_value As String = logvalue.Value
                Dim log_totalstr As String = ""

                If log_value.Trim <> "" Then '20160530 ログ修正 ログが空でない場合の条件追加

                    '処理項目を分割
                    Dim tmp_syori() As String = log_key.Split("/")

                    '日本語テーブル名を取得して「処理項目_大分類」へ格納
                    Dim tmp_tblfldname As String = tmp_syori(0)                 '※テーブル名は共通なので配列のインデックスは任意の値
                    Dim tmp_tblname() As String = tmp_tblfldname.Split("-")
                    Dim syorikomk As String = Hash_TblName_AlphaToJp(tmp_tblname(0))

                    '日本語フィールド名を取得して「処理項目_小分類」へ格納
                    Dim tmp_syoriitem As String = ""
                    For cntii = 0 To UBound(tmp_syori)
                        Dim tmp_str As String = Hash_FldName_AlphaToJp(tmp_syori(cntii))
                        tmp_syoriitem = tmp_syoriitem & "、" & tmp_str
                    Next
                    tmp_syoriitem = tmp_syoriitem.Remove(0, 1)
                    Dim syoriitem As String = tmp_syoriitem

                    '「-」で連結されたログ内容を分割
                    Dim tmp_log() As String = log_value.Split("-")

                    'ログの中身を各変数へ格納
                    Dim syorikekka As String = tmp_log(0)
                    Dim hubinaiyo As String = tmp_log(1)
                    Dim taisyo As String = tmp_log(2)

                    '調整前後の値を変数へ格納
                    '20160829 口座名義カナチェック機能の追加 -chg sta
                    'Dim tmp_before As String = hash_beforechk(tmp_tblfldname(1))
                    'Dim tmp_after As String = hash_afterchk(tmp_tblfldname(1))
                    Dim tmp_before As String = IIf(hash_beforechk(tmp_tblname(1)) Is Nothing, "", hash_beforechk(tmp_tblname(1)))
                    Dim tmp_after As String = IIf(hash_afterchk(tmp_tblname(1)) Is Nothing, "", hash_afterchk(tmp_tblname(1)))
                    '20160829 口座名義カナチェック機能の追加 -chg end
                    Dim tmp_empty As String = "-"
                    Dim beforechgvalue As String = ""
                    Dim afterchgvalue As String = ""
                    '20160829 口座名義カナチェック機能の追加 -chg sta
                    'If hash_afterchk.Count = 0 Then         '移行されなかった場合(チェック後ハッシュテーブルにデータが存在しない)は値を設定しない
                    '    beforechgvalue = tmp_empty
                    '    afterchgvalue = tmp_empty
                    'ElseIf tmp_before = tmp_after Then      '調整前後で値が同じ場合は値を設定しない
                    '    beforechgvalue = tmp_empty
                    '    afterchgvalue = tmp_empty
                    'ElseIf tmp_before <> tmp_after Then     '調整前後で値が異なる場合は値を設定する(出力時を考慮しカンマを全角に変換しておく)
                    '    beforechgvalue = tmp_before.Replace(",", "，")
                    '    afterchgvalue = tmp_after.Replace(",", "，")
                    'End If
                    If tmp_before = tmp_after Then          '調整前後で値が同じ場合は値を設定しない
                        beforechgvalue = tmp_empty
                        afterchgvalue = tmp_empty
                    ElseIf tmp_before <> tmp_after Then     '調整前後で値が異なる場合は値を設定する(出力時を考慮しカンマを全角に変換しておく)
                        '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
                        '最後にまとめて変換処理を行うためここではそのまま設定するように修正
                        'beforechgvalue = tmp_before.Replace(",", "，")
                        'afterchgvalue = tmp_after.Replace(",", "，")
                        beforechgvalue = tmp_before
                        afterchgvalue = tmp_after
                        '20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
                    End If
                    '20160829 口座名義カナチェック機能の追加 -chg end
                    '最終的なログ文字列群を取得
                    log_totalstr = Set_LogValue(42, syorikomk, syoriitem, syorikekka, taisyodata, hubinaiyo, taisyo, beforechgvalue, afterchgvalue, tblname)

                    '挿入用クエリへ加工
                    Dim log_sql As String = Get_LogTblInsertQry(log_totalstr, False)

                    'ログカウントを更新してリストへ格納
                    logcnt = logcnt + 1
                    sortlist.Add(logcnt, log_sql)

                End If

            Next

        End Sub

    End Class

#End Region

#Region "DB書込処理クラス"

    Public Class CVDBInsert

        ''' <summary>
        ''' 革命10DBへのINSERT処理実行
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="tblname"></param>
        ''' <param name="fldname"></param>
        ''' <param name="hash_cvitem"></param>
        ''' <remarks></remarks>
        Public Shared Sub Cnv_Db(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByVal fldname As String, ByVal hash_cvitem As Hashtable, ByRef flg As Boolean)

            '初期値設定
            flg = True

            'パラメータ生成
            Dim repprm As String = CVDBInsert.Get_Parameter(fldname)

            'フィールド名を分割して配列へ格納
            Dim tmp_fld() As String = fldname.Split(",")

            'クエリ作成
            Dim tmpcomtxt As String = "INSERT INTO " & tblname & " (" & fldname & ") VALUES (" & repprm & ")"
            Dim tmpsqlcom As New SqlCommand(tmpcomtxt, sqlcnnv10)

            'DB挿入値取得
            For cntjj = 0 To UBound(tmp_fld)
                Dim cv_value As String = hash_cvitem(tmp_fld(cntjj))
                CVDBInsert.SqlComAdd(tmpsqlcom, IIf(cv_value = Nothing, System.DBNull.Value, cv_value), tblname & "-" & tmp_fld(cntjj), "@" & tmp_fld(cntjj))
            Next

            '挿入処理
            Try
                Dim rowsAffected As Integer = tmpsqlcom.ExecuteNonQuery()
            Catch ex As Exception                                               'kakaka INSERT失敗を重複と判断、呼び出し先にてnormalflgへセットして重複データのカウントを取っている？↓
                flg = False                                                     'kakaka 重複エラーではない場合も重複エラーとカウントされてる？
            Finally
                '終了処理
                tmpsqlcom.Dispose()
            End Try

        End Sub

        ''' <summary>
        ''' 革命10フィールドタイプ設定
        ''' </summary>
        ''' <param name="tmpsqlcom"></param>
        ''' <param name="objval"></param>
        ''' <param name="sqlprm"></param>
        ''' <remarks></remarks>
        Public Shared Sub SqlComAdd(tmpsqlcom As SqlCommand, objval As Object, tblfldname As String, sqlprm As String)

            Dim spltyp As Integer
            Dim tmp_type As String = Hash_FiledTypeAlpha(tblfldname)

            Select Case tmp_type
                Case "int"
                    spltyp = SqlDbType.Int
                Case "smallint"
                    spltyp = SqlDbType.SmallInt
                Case "varchar"
                    spltyp = SqlDbType.VarChar
                Case "xml"
                    spltyp = SqlDbType.Xml
                Case "uniqueidentifier"
                    spltyp = SqlDbType.UniqueIdentifier
                Case "datetime"
                    spltyp = SqlDbType.DateTime
                Case "money"
                    spltyp = SqlDbType.Money
                Case "nvarchar"
                    spltyp = SqlDbType.NVarChar
                Case "float"
                    spltyp = SqlDbType.Float
                Case "varbinary"
                    spltyp = SqlDbType.VarBinary
                Case "nchar"
                    spltyp = SqlDbType.NChar
                Case "decimal"
                    spltyp = SqlDbType.Decimal
                Case "date"
                    spltyp = SqlDbType.Date
                    'Case "sysname"
                    '    spltyp = SqlDbType.sysname
                Case "bit"
                    spltyp = SqlDbType.Bit
                Case "bigint"
                    spltyp = SqlDbType.BigInt
            End Select

            tmpsqlcom.Parameters.AddWithValue(sqlprm, spltyp)
            tmpsqlcom.Parameters(sqlprm).Value = objval

        End Sub

        ''' <summary>
        ''' カンマで連結されたフィールド名群の各フィールド名に「@」を付加
        ''' </summary>
        ''' <param name="fldnamegrp"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Parameter(ByVal fldnamegrp As String) As String

            Dim rtn_str As String = ""
            Dim tmp_str As String = ""
            Dim tmp_parastr() As String = fldnamegrp.Split(",")

            For cntii = 0 To UBound(tmp_parastr)
                tmp_str = tmp_str & ",@" & tmp_parastr(cntii)
            Next

            rtn_str = tmp_str.Remove(0, 1)

            Return rtn_str

        End Function

    End Class

#End Region

#Region "部屋間取関連クラス"

    Public Class MadoriConv

        ''' <summary>
        ''' 間取り分割処理
        ''' </summary>
        ''' <param name="chkmadori"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' ・1データのみ対象<br/>
        '''   例：「2LDK」…「2」と「LDK」に分割
        '''       「3DK+4LDK」…「3」と「DK+4LDK」に分割
        '''       「LDK」…「」と「LDK」に分割
        '''       「345」…「345」と「」に分割
        ''' </remarks>
        Public Shared Sub SplitMadori(ByVal madoristr As String, ByRef madoricnt As String, ByRef madorinaiyo As String)

            Dim numflg As Boolean = False                                           '全数値FLG…True.末尾
            Dim madori_num As String = ""
            Dim madori_nai As String = ""
            Dim tmpmadori As String = ""                                            '作業用：文字列編集用
            Dim tmpstr As String = ""                                               '作業用：結合用
            Dim tmpchr As String = ""                                               '作業用：1文字チェック

            tmpmadori = madoristr.ToString.Trim

            If tmpmadori <> "" Then

                tmpmadori = (StrConv(tmpmadori, VbStrConv.Narrow).ToUpper)          '半角大文字変換

                Dim len As Integer = tmpmadori.Length                               'チェック文字数

                For cnt As Integer = 0 To len - 1
                    '1文字ずつチェック
                    tmpchr = tmpmadori.Substring(cnt, 1)
                    If IsNumeric(tmpchr) Then
                        tmpstr = tmpstr & tmpchr
                    Else
                        Exit For
                    End If
                    If cnt = len - 1 Then
                        '全て数値の場合
                        numflg = True
                    End If
                Next

                madori_num = tmpstr

                If madori_num <> "" Then
                    madori_nai = IIf(numflg = True, "", tmpmadori.Replace(madori_num, ""))
                Else
                    madori_nai = tmpmadori
                End If

            End If

            madoricnt = madori_num
            madorinaiyo = madori_nai

        End Sub

        ''' <summary>
        ''' 間取り分割処理(内訳)
        ''' </summary>
        ''' <param name="madoristr"></param>
        ''' <param name="madorino"></param>
        ''' <param name="madorikbn"></param>
        ''' <param name="madorijo"></param>
        ''' <remarks></remarks>
        Public Shared Sub SplitMadori(ByVal madoristr As String, ByRef madorino() As String, ByRef madorikbn() As String, ByRef madorijo() As String)

            '空文字チェック
            If madoristr = "" Then
                Exit Sub
            End If

            '半角スペースで連結された文字列を分割
            Dim tmp_madorigrp() As String = madoristr.Split(" ")

            ReDim madorino(UBound(tmp_madorigrp))
            ReDim madorikbn(UBound(tmp_madorigrp))
            ReDim madorijo(UBound(tmp_madorigrp))

            For cntii = 0 To UBound(tmp_madorigrp)
                '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg sta
                ''20160822 部屋間取内訳情報文字列分割処理修正 -chg sta
                ''空文字の場合の処理を入れる
                ' ''間取型*畳数の形になっているデータをさらに分割する
                ''Dim tmp_madoripart() As String = tmp_madorigrp(cntii).Split("*")

                ' ''各配列へ格納
                ''madorino(cntii) = cntii + 1
                ''madorikbn(cntii) = tmp_madoripart(0)
                ''madorijo(cntii) = tmp_madoripart(1)

                ''間取No格納(歯抜で登録されたデータは詰めずにそのまま移行する)
                'madorino(cntii) = cntii + 1

                ''データが存在する場合は分割して格納する
                'If tmp_madorigrp(cntii) <> "" Then

                '    '間取型*畳数の形になっているデータをさらに分割する
                '    Dim tmp_madoripart() As String = tmp_madorigrp(cntii).Split("*")

                '    '各配列へ格納
                '    madorikbn(cntii) = tmp_madoripart(0)
                '    madorijo(cntii) = tmp_madoripart(1)

                'End If
                ''20160822 部屋間取内訳情報文字列分割処理修正 -chg end


                '間取No格納(歯抜で登録されたデータは詰めずにそのまま移行する)
                madorino(cntii) = cntii + 1

                'データが存在する場合は分割して格納する
                If tmp_madorigrp(cntii) <> "" Then

                    '「*」が含まれているかチェック
                    If tmp_madorigrp(cntii).Replace("*", "") = tmp_madorigrp(cntii) Then
                        '「*」が含まれていない場合は「対応不要」を付加して格納
                        madorikbn(cntii) = tmp_madorigrp(cntii) & "(対応不要)"
                        madorijo(cntii) = ""
                    Else
                        '間取型*畳数の形になっている場合はさらに分割する
                        Dim tmp_madoripart() As String = tmp_madorigrp(cntii).Split("*")

                        '各配列へ格納
                        madorikbn(cntii) = tmp_madoripart(0)
                        madorijo(cntii) = tmp_madoripart(1)
                    End If

                End If
                '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg end
            Next

        End Sub

    End Class

#End Region

#Region "都市計画/用途地域関連クラス"

    Public Class TosiYotoConv

        Public Shared Sub Get_TosiYoto(ByVal tosiyoto As String, ByRef tosiyotokbn As Integer, ByRef tosiyotono As String)

            Select Case tosiyoto
                Case "第一種低層住居専用地域"
                    tosiyotokbn = 2
                    tosiyotono = "1"
                Case "第二種低層住居専用地域"
                    tosiyotokbn = 2
                    tosiyotono = "2"
                Case "第一種中高層住居専用地域"
                    tosiyotokbn = 2
                    tosiyotono = "3"
                Case "第二種中高層住居専用地域"
                    tosiyotokbn = 2
                    tosiyotono = "4"
                Case "第一種住居地域"
                    tosiyotokbn = 2
                    tosiyotono = "5"
                Case "第二種住居地域"
                    tosiyotokbn = 2
                    tosiyotono = "6"
                Case "準住居地域"
                    tosiyotokbn = 2
                    tosiyotono = "7"
                Case "近隣商業地域"
                    tosiyotokbn = 2
                    tosiyotono = "8"
                Case "商業地域"
                    tosiyotokbn = 2
                    tosiyotono = "9"
                Case "準工業地域"
                    tosiyotokbn = 2
                    tosiyotono = "10"
                Case "工業地域"
                    tosiyotokbn = 2
                    tosiyotono = "11"
                Case "工業専用地域"
                    tosiyotokbn = 2
                    tosiyotono = "12"
                Case "指定なし"
                    tosiyotono = "13"
                    tosiyotokbn = 2
                Case "都市計画区域外"
                    tosiyotokbn = 1
                    tosiyotono = "7"
                Case ""
                    tosiyotokbn = 0
                    tosiyotono = "-1"
                Case Else
                    tosiyotokbn = 3
                    tosiyotono = "99"
            End Select

        End Sub

    End Class

#End Region

#Region "紐付項目取得関連クラス"

    Public Class SetRelItemToObject

        ''' <summary>
        ''' 入金項目紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Nkinkomk()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "入金項目マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '既存データ有無確認(入金項目は他の項目でも参照するため)
            If Hash_Rel_Nkinkomk.Count <> 0 Then
                Exit Sub
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldnkinname As String = ""
                Dim tmp_newnkinno As String = ""
                Dim tmp_nkinkbn As String = ""
                Dim tmp_hendometerno As String = ""
                Dim tmp_newnkinzkseino As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元入金項目名称"
                            tmp_oldnkinname = fldvalue.Trim
                        Case "賃貸革命10入金項目No"
                            tmp_newnkinno = fldvalue.Trim
                            '2016.04.06 入金項目読込処理の修正 -add sta
                        Case "移行元入金項目区分"
                            tmp_nkinkbn = fldvalue.Trim
                        Case "賃貸革命10変動費メーター分類No"
                            tmp_hendometerno = fldvalue.Trim
                            '2016.04.06 入金項目読込処理の修正 -add end
                            '20160617 マイナス金額移行処理修正 -add sta
                        Case "賃貸革命10入金項目属性No"
                            tmp_newnkinzkseino = fldvalue.Trim
                            '20160617 マイナス金額移行処理修正 -add end
                    End Select

                Next

                '入金項目をハッシュテーブル格納
                If tmp_newnkinno <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Nkinkomk.Add(tmp_oldnkinname & "-" & tmp_nkinkbn, tmp_newnkinno)
                    If Hash_Rel_Nkinkomk.Contains(tmp_oldnkinname & "-" & tmp_nkinkbn) = False Then
                        Hash_Rel_Nkinkomk.Add(tmp_oldnkinname & "-" & tmp_nkinkbn, tmp_newnkinno)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

                '2016.04.06 入金項目読込処理の修正 -add sta
                '変動費メーター分類をハッシュテーブル格納
                If tmp_nkinkbn = "5.随時変動" And tmp_hendometerno <> "" Then
                    Hash_Rel_Hendometer.Add(tmp_oldnkinname & "-" & tmp_nkinkbn, tmp_hendometerno)
                End If
                '2016.04.06 入金項目読込処理の修正 -add end

                '20160617 マイナス金額移行処理修正 -add sta
                If Hash_Rel_NkinkomkZksei.Contains(tmp_oldnkinname & "-" & tmp_nkinkbn) = False Then
                    Hash_Rel_NkinkomkZksei.Add(tmp_oldnkinname & "-" & tmp_nkinkbn, tmp_newnkinzkseino)
                End If
                '20160617 マイナス金額移行処理修正 -add end

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 入金区分紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Nkinkbn()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "入金区分マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '既存データ有無確認(入金項目は他の項目でも参照するため)
            If Hash_Rel_Nkinkbn.Count <> 0 Then
                Exit Sub
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldnkinkbnname As String = ""
                Dim tmp_newnkinkbnno As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元入金区分名称"
                            tmp_oldnkinkbnname = fldvalue.Trim
                        Case "賃貸革命10入金区分No"
                            tmp_newnkinkbnno = fldvalue.Trim
                    End Select

                Next

                'ハッシュテーブル格納
                If tmp_newnkinkbnno <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
                    If Hash_Rel_Nkinkbn.Contains(tmp_oldnkinkbnname) = False Then
                        Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 設備紐付情報取得 (V7のデータと10の設備Noを紐付)
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Setubi_V7To10No()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "設備マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_setubigrpname As String = ""
                Dim tmp_setubiname As String = ""
                '2016.04.06 部屋設備情報の取得処理修正 -chg sta
                'Dim tmp_setubiguid As String = ""
                Dim tmp_setubigrpno As String = ""
                Dim tmp_setubino As String = ""
                Dim tmp_setubikomkno As String = ""
                '2016.04.06 部屋設備情報の取得処理修正 -chg end

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元設備グループ名称"
                            tmp_setubigrpname = fldvalue.Trim
                        Case "移行元設備名称"
                            tmp_setubiname = fldvalue.Trim
                            '2016.04.06 部屋設備情報の取得処理修正 -chg sta
                            'Case "賃貸革命10項目GUID"
                            '    tmp_setubiguid = fldvalue.Trim
                        Case "賃貸革命10設備グループNo"
                            tmp_setubigrpno = fldvalue.Trim
                        Case "賃貸革命10設備No"
                            tmp_setubino = fldvalue.Trim
                        Case "賃貸革命10項目No"
                            tmp_setubikomkno = fldvalue.Trim
                            '2016.04.06 部屋設備情報の取得処理修正 -chg end
                    End Select

                Next

                'ハッシュテーブル格納
                '2016.04.06 部屋設備情報の取得処理修正 -chg sta
                'If tmp_setubiguid <> "" Then
                '    Hash_Rel_Setubi.Add(tmp_setubigrpname & "-" & tmp_setubiname, tmp_setubiguid)
                'End If
                Dim tmp_V7setubikey As String = tmp_setubigrpname & "-" & tmp_setubiname
                '20160829 エレベーター移行対応 -chg sta
                'Dim tmp_10setubikey As String = tmp_setubigrpno & "-" & tmp_setubino & "-" & tmp_setubikomkno
                Dim tmp_10setubikey As String = ""

                If tmp_setubigrpno <> "999" And (tmp_setubigrpno = "" Or tmp_setubino = "" Or tmp_setubikomkno = "") Then
                    tmp_10setubikey = ""
                ElseIf tmp_setubigrpno = "999" And tmp_setubino <> "" Then
                    tmp_10setubikey = tmp_setubigrpno & "-" & tmp_setubino
                Else
                    tmp_10setubikey = tmp_setubigrpno & "-" & tmp_setubino & "-" & tmp_setubikomkno
                End If
                '20160829 エレベーター移行対応 -chg end
                '20160829 エレベーター移行対応 -chg sta
                'If Hash_SetubiMid.Contains(tmp_V7setubikey) = False Then
                '    Hash_SetubiMid.Add(tmp_V7setubikey, tmp_10setubikey)
                'End If
                If Hash_SetubiMid.Contains(tmp_V7setubikey) = False Or tmp_setubigrpno = "999" Then
                    Hash_SetubiMid.Add(tmp_V7setubikey, tmp_10setubikey)
                End If
                '20160829 エレベーター移行対応 -chg end
                '2016.04.06 部屋設備情報の取得処理修正 -chg end
            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 契約分類紐付情報取得 2016.04.06 契約分類を紐付データを元に移行
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Kyrui()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "契約分類マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '初期化
            Hash_Rel_Kyrui.Clear()
            Hash_Rel_Kyrui_Teikisyakuyakbn.Clear()

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldkyruiname As String = ""
                Dim tmp_newkyruino As String = ""
                Dim tmp_newkyruiteisyakukbn As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元契約分類名称"
                            tmp_oldkyruiname = fldvalue.Trim
                        Case "賃貸革命10契約分類No"
                            tmp_newkyruino = fldvalue.Trim
                        Case "定期借家として扱う"
                            tmp_newkyruiteisyakukbn = IIf(fldvalue = "True", 1, 0)
                    End Select

                Next

                'ハッシュテーブル格納(契約分類データ)
                If tmp_newkyruino <> "" Then
                    If Hash_Rel_Kyrui.Contains(tmp_oldkyruiname) = False Then
                        Hash_Rel_Kyrui.Add(tmp_oldkyruiname, tmp_newkyruino)
                    End If
                End If

                'ハッシュテーブル格納(契約分類の定期借家区分)
                If tmp_newkyruino <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Kyrui_Teikisyakuyakbn.Add(tmp_oldkyruiname, tmp_newkyruiteisyakukbn)
                    If Hash_Rel_Kyrui_Teikisyakuyakbn.Contains(tmp_oldkyruiname) = False Then
                        Hash_Rel_Kyrui_Teikisyakuyakbn.Add(tmp_oldkyruiname, tmp_newkyruiteisyakukbn)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 取引態様マスタ紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Toritaiyo()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "取引態様マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldtoritaiyoname As String = ""
                Dim tmp_newtoritaiyono As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元取引態様名称"
                            tmp_oldtoritaiyoname = fldvalue.Trim
                        Case "賃貸革命10取引態様No"
                            tmp_newtoritaiyono = fldvalue.Trim
                    End Select

                Next

                'ハッシュテーブル格納
                If tmp_newtoritaiyono <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Toritaiyo.Add(tmp_oldtoritaiyoname, tmp_newtoritaiyono)
                    If Hash_Rel_Toritaiyo.Contains(tmp_oldtoritaiyoname) = False Then
                        Hash_Rel_Toritaiyo.Add(tmp_oldtoritaiyoname, tmp_newtoritaiyono)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 物件分類紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Bkrui()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "物件分類マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldbkruiname As String = ""
                Dim tmp_newbkruino As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元物件分類名称"
                            tmp_oldbkruiname = fldvalue.Trim
                        Case "賃貸革命10物件分類No"
                            tmp_newbkruino = fldvalue.Trim
                    End Select

                Next

                'ハッシュテーブル格納
                If tmp_newbkruino <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Bkrui.Add(tmp_oldbkruiname, tmp_newbkruino)
                    If Hash_Rel_Bkrui.Contains(tmp_oldbkruiname) = False Then
                        Hash_Rel_Bkrui.Add(tmp_oldbkruiname, tmp_newbkruino)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 部屋分類紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Hyrui()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "部屋分類マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '既存データ有無確認 20160829 革命10バージョンアップに伴う修正 -add
            If Hash_Rel_Hyrui.Count <> 0 Then
                Exit Sub
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldhyruiname As String = ""
                Dim tmp_newhyruino As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元部屋分類名称"
                            tmp_oldhyruiname = fldvalue.Trim
                        Case "賃貸革命10部屋分類No"
                            tmp_newhyruino = fldvalue.Trim
                    End Select

                Next

                'ハッシュテーブル格納
                If tmp_newhyruino <> "" Then
                    '20160829 革命10バージョンアップに伴う修正 -chg sta
                    'Hash_Rel_Hyrui.Add(tmp_oldhyruiname, tmp_newhyruino)
                    If Hash_Rel_Hyrui.Contains(tmp_oldhyruiname) = False Then
                        Hash_Rel_Hyrui.Add(tmp_oldhyruiname, tmp_newhyruino)
                    End If
                    '20160829 革命10バージョンアップに伴う修正 -chg end
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 構造マスタ紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Kozo()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "構造マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldbkozoname As String = ""
                Dim tmp_newkozono As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元構造名称"
                            tmp_oldbkozoname = fldvalue.Trim
                        Case "賃貸革命10構造No"
                            tmp_newkozono = fldvalue.Trim
                    End Select

                Next

                'ハッシュテーブル格納
                If tmp_newkozono <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Kozo.Add(tmp_oldbkozoname, tmp_newkozono)
                    If Hash_Rel_Kozo.Contains(tmp_oldbkozoname) = False Then
                        Hash_Rel_Kozo.Add(tmp_oldbkozoname, tmp_newkozono)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 用途地域・都市計画マスタ紐付情報取得 2016.04.06 都市計画・用途地域を紐付データから取得するように修正
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Yototiki()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            '20161014 都市計画用途地域ファイルオープンエラー修正 -chg sta
            'Dim relsheetname As String = "用途地域・都市計画マスタ"
            Dim relsheetname As String = "都市計画・用途地域マスタ"
            '20161014 都市計画用途地域ファイルオープンエラー修正 -chg end
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '初期化
            Hash_Rel_Tosiyotokbn.Clear()
            Hash_Rel_Tosiyotono.Clear()

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldyotoname As String = ""
                Dim tmp_newyototikikbn As String = ""
                Dim tmp_newyototikino As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元用途地域・都市計画名"
                            tmp_oldyotoname = fldvalue.Trim
                        Case "移行先用途地域・都市計画区分"
                            tmp_newyototikikbn = fldvalue.Trim
                        Case "移行先用途地域・都市計画No"
                            tmp_newyototikino = fldvalue.Trim
                    End Select

                Next

                'ハッシュテーブル格納
                If tmp_newyototikikbn <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Tosiyotokbn.Add(tmp_oldyotoname, tmp_newyototikikbn)
                    If Hash_Rel_Tosiyotokbn.Contains(tmp_oldyotoname) = False Then
                        Hash_Rel_Tosiyotokbn.Add(tmp_oldyotoname, tmp_newyototikikbn)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

                If tmp_newyototikino <> "" Then
                    '20160527 指摘事項対応 -chg sta
                    'Hash_Rel_Tosiyotono.Add(tmp_oldyotoname, tmp_newyototikino)
                    If Hash_Rel_Tosiyotono.Contains(tmp_oldyotoname) = False Then
                        Hash_Rel_Tosiyotono.Add(tmp_oldyotoname, tmp_newyototikino)
                    End If
                    '20160527 指摘事項対応 -chg end
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 自社口座紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_Jisyakoza()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "自社口座マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '既存データ有無確認
            If Hash_Rel_JisyaKoza.Count <> 0 Then
                Exit Sub
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                '20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                'Dim tmp_kinyuno As String = ""
                'Dim tmp_kinyutenno As String = ""
                'Dim tmp_kozasyu As String = ""   '20160609 紐付取得用に追加
                'Dim tmp_kozabango As String = ""
                Dim tmp_kozasyutokumoto As String = ""
                Dim tmp_kozasyutokumotono As String = ""
                '20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                Dim tmp_jisyano As String = ""
                Dim tmp_jisyakozano As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        '20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                        'Case "金融機関No"
                        '    tmp_kinyuno = fldvalue.Trim
                        'Case "支店No"
                        '    tmp_kinyutenno = fldvalue.Trim
                        'Case "口座種別" '20160609 紐付取得用に追加
                        '    tmp_kozasyu = fldvalue.Trim
                        'Case "口座番号"
                        '    tmp_kozabango = fldvalue.Trim
                        'Case "自社No"
                        '    tmp_jisyano = fldvalue.Trim
                        'Case "自社口座No"
                        '    tmp_jisyakozano = fldvalue.Trim
                        Case "口座取得元"
                            tmp_kozasyutokumoto = fldvalue.Trim
                        Case "各口座設定No"
                            tmp_kozasyutokumotono = fldvalue.Trim
                        Case "自社No"
                            tmp_jisyano = fldvalue.Trim
                        Case "自社口座No"
                            tmp_jisyakozano = fldvalue.Trim
                            '20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                    End Select

                Next

                '紐付け用に値を成形
                '20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                ''20160609 紐付取得用に追加 -chg sta
                ''Dim tmp_kozainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                'Dim tmp_kozainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                ''20160609 紐付取得用に追加 -chg end
                Dim tmp_kozainfo As String = tmp_kozasyutokumoto & "-" & tmp_kozasyutokumotono
                '20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                Dim tmp_jisyainfo As String = tmp_jisyano & "-" & tmp_jisyakozano

                'ハッシュテーブル格納
                If tmp_jisyano <> "" AndAlso tmp_jisyakozano <> "" Then
                    If Hash_Rel_JisyaKoza.Contains(tmp_kozainfo) = False Then
                        Hash_Rel_JisyaKoza.Add(tmp_kozainfo, tmp_jisyainfo)
                    End If
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' FB関連紐付情報取得 '20160609 紐付設定値取得に伴う修正
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_FBInfo()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "FBフォーマット割付"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'オブジェクト初期化
            Hash_Rel_FBInfo_KozaFurikaeFmt.Clear()
            Hash_Rel_FBInfo_FuriIraiFmt.Clear()
            Hash_Rel_FBInfo_NsSettingFmt.Clear()
            Hash_Rel_FBInfo_NsSettingJisya.Clear()

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_V7fmtsyubetu As String = ""
                Dim tmp_V7kozano As String = ""
                Dim tmp_10fmtno As String = ""
                Dim tmp_jisyano As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "フォーマット種別"
                            tmp_V7fmtsyubetu = fldvalue.Trim
                        Case "各口座No"
                            tmp_V7kozano = fldvalue.Trim
                        Case "フォーマットNo"
                            tmp_10fmtno = fldvalue.Trim
                        Case "自社No"
                            tmp_jisyano = fldvalue.Trim
                    End Select

                Next

                Select Case tmp_V7fmtsyubetu
                    Case "口座振替"
                        If Hash_Rel_FBInfo_KozaFurikaeFmt.Contains(tmp_V7kozano) = False Then
                            Hash_Rel_FBInfo_KozaFurikaeFmt.Add(tmp_V7kozano, tmp_10fmtno)
                        End If
                    Case "総合振込"
                        If Hash_Rel_FBInfo_FuriIraiFmt.Contains(tmp_V7kozano) = False Then
                            Hash_Rel_FBInfo_FuriIraiFmt.Add(tmp_V7kozano, tmp_10fmtno)
                        End If
                    Case "入出金"
                        If Hash_Rel_FBInfo_NsSettingFmt.Contains(tmp_V7kozano) = False Then
                            Hash_Rel_FBInfo_NsSettingFmt.Add(tmp_V7kozano, tmp_10fmtno)
                        End If
                        If Hash_Rel_FBInfo_NsSettingJisya.Contains(tmp_V7kozano) = False Then
                            Hash_Rel_FBInfo_NsSettingJisya.Add(tmp_V7kozano, tmp_jisyano)
                        End If
                End Select

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        ''' <summary>
        ''' 鍵タイトルマスタ情報取得 '20160613 鍵情報の取得処理修正
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Set_RelData_KagiInfo(ByVal sqlcnnv10 As SqlConnection)

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "鍵タイトルマスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'オブジェクト初期化
            Hash_KagiTitle_V7to10.Clear()

            '鍵タイトル情報仮テーブル
            Dim tmp_kagiinfo_tblname As String = "tmp_kagiinfo"
            Dim tmp_rowcnt As Integer = 0

            '鍵タイトル情報仮テーブル初期化
            Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(tmp_kagiinfo_tblname, True)
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmp_rowcnt)

            '鍵タイトル情報仮テーブル作成
            Dim tmp_sql_create As String = ""
            tmp_sql_create = tmp_sql_create & " CREATE TABLE " & tmp_kagiinfo_tblname
            tmp_sql_create = tmp_sql_create & " 	( "
            tmp_sql_create = tmp_sql_create & " 		 oldkagino INT "
            tmp_sql_create = tmp_sql_create & " 		,newkagikbn INT "
            tmp_sql_create = tmp_sql_create & " 	) "
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmp_rowcnt)

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_V7kagino As String = ""
                Dim tmp_10kagisyubetu As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "鍵No"
                            tmp_V7kagino = fldvalue.Trim
                        Case "共用チェック"
                            tmp_10kagisyubetu = IIf(fldvalue.Trim = "True", 1, 2)
                    End Select

                Next

                '仮テーブルへ挿入
                Dim tmp_insertitem As String = tmp_V7kagino & "," & tmp_10kagisyubetu
                Dim tmp_sql_insert As String = " INSERT INTO " & tmp_kagiinfo_tblname & " VALUES(" & tmp_insertitem & ");"
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_rowcnt)

            Next

            '仮テーブルからV7鍵Noと10鍵No(鍵区分を含む)を紐付けたハッシュテーブルを作成する
            Dim tmp_sql_select As String = ""
            tmp_sql_select = tmp_sql_select & " SELECT "
            tmp_sql_select = tmp_sql_select & " 	 oldkagino "
            tmp_sql_select = tmp_sql_select & " 	,CONVERT(VARCHAR,newkagikbn) + '-' + CONVERT(VARCHAR,ROW_NUMBER()OVER(PARTITION BY newkagikbn ORDER BY oldkagino)) AS newkagino "
            tmp_sql_select = tmp_sql_select & " FROM tmp_kagiinfo "
            DBExec.Exec_DataReader_Col_Hash(tmp_sql_select, sqlcnnv10, Hash_KagiTitle_V7to10)

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            '鍵タイトル情報仮テーブル削除
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmp_rowcnt)

        End Sub

        ''' <summary>
        ''' 口座種別紐付情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Get_RelData_KozaSyubetu()

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim relsheetname As String = "口座種別マスタ"
            Dim rtn As Boolean = True

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '既存データ有無確認(口座種別は他の項目でも参照するため)
            If Hash_Rel_Kozasyubetu.Count <> 0 Then
                Exit Sub
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            For cntii = 0 To rowcnt - 1

                '作業用変数作成
                Dim tmp_oldkozasyubetuname As String = ""
                Dim tmp_newkozasyubetuno As String = ""

                '紐付設定値取得
                For cntjj = 1 To columncnt

                    'ヘッダー格納
                    Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                    '移行値格納
                    Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                    Select Case fldname
                        Case "移行元口座種別名称"
                            tmp_oldkozasyubetuname = fldvalue.Trim
                        Case "賃貸革命10口座種別No"
                            tmp_newkozasyubetuno = fldvalue.Trim
                    End Select

                Next

                'ハッシュテーブル格納
                If tmp_newkozasyubetuno <> "" Then
                    If Hash_Rel_Kozasyubetu.Contains(tmp_oldkozasyubetuname) = False Then
                        Hash_Rel_Kozasyubetu.Add(tmp_oldkozasyubetuname, tmp_newkozasyubetuno)
                    End If
                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        End Sub

        '20160609 紐付設定値取得に伴う修正 -del sta
        ' ''' <summary>
        ' ''' FBフォーマット紐付情報取得 '20160517 入出金取得情報の新規作成
        ' ''' </summary>
        ' ''' <remarks></remarks>
        'Public Shared Sub Set_RelData_FBFmt()

        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim relsheetname As String = "FBフォーマットマスタ"
        '    Dim rtn As Boolean = True

        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

        '    'Excelファイル初期設定
        '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

        '    For cntii = 0 To rowcnt - 1

        '        '作業用変数作成
        '        Dim tmp_oldfmtname As String = ""
        '        Dim tmp_newfmtno As String = ""

        '        '紐付設定値取得
        '        For cntjj = 1 To columncnt

        '            'ヘッダー格納
        '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

        '            '移行値格納
        '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

        '            Select Case fldname
        '                Case "移行元フォーマット名称"
        '                    tmp_oldfmtname = fldvalue.Trim
        '                Case "賃貸革命10フォーマットNo"
        '                    tmp_newfmtno = fldvalue.Trim
        '            End Select

        '        Next

        '        'ハッシュテーブル格納
        '        If tmp_newfmtno <> "" Then
        '            '20160527 指摘事項対応 -chg sta
        '            'Hash_Rel_FBFmt.Add(tmp_oldfmtname, tmp_newfmtno)
        '            If Hash_Rel_FBFmt.Contains(tmp_oldfmtname) = False Then
        '                Hash_Rel_FBFmt.Add(tmp_oldfmtname, tmp_newfmtno)
        '            End If
        '            '20160527 指摘事項対応 -chg end
        '        End If

        '    Next

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        'End Sub
        '20160609 紐付設定値取得に伴う修正 -del end

    End Class

#End Region


End Namespace
