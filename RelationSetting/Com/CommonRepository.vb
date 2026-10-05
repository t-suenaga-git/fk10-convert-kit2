Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports System.Text.RegularExpressions
Imports System.IO

Namespace Njc.Common

#Region "共通処理クラス"

    ''' <summary>
    ''' 共通処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CommonRepository

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
        ''' 接続OPEN処理
        ''' </summary>
        ''' <param name="cnninfo">接続情報</param>
        ''' <param name="sqlcnn">接続 [out]</param>
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
                '2015.07.06 sol レビュー後修正_レビュー№14 -chg sta
                ''==================================================
                '' SQL Server認証を利用して接続
                ''==================================================
                'cnnset.ConnectionString = _
                '     "Data Source = " & cnninfo.ServerName & _
                '     ";Initial Catalog = " & cnninfo.InitialCatalog & _
                '     ";User ID = " & cnninfo.User & _
                '     ";Password = " & cnninfo.Pass & _
                '     ";Connection Timeout = " & cnninfo.TimeOut

                ''==================================================
                '' Windows認証を利用して接続
                ''==================================================
                ''cnnset.ConnectionString = _
                ''     "Data Source = " & cnninfo.ServerName & _
                ''     ";Initial Catalog = " & cnninfo.CatalogName & _
                ''     ";Integrated Security = SSPI" & _
                ''      ";Connection Timeout = " & cnninfo.TimeOut
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
                '2015.07.06 sol レビュー後修正_レビュー№14 -chg end

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
        ''' 接続CLOSE処理(リソース解放)
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

        ' ''' <summary>
        ' ''' DBより全レコードを取得して返却 [DataTable]
        ' ''' </summary>
        ' ''' <param name="sqlstr">取得クエリ文</param>
        ' ''' <param name="sqlcnn">接続情報</param>
        ' ''' <param name="readtbl">読込TBL情報 [out]</param>
        ' ''' <returns>レコード件数<br/></returns>
        ' ''' <remarks>
        ' ''' ・sqlcnnv7/v10→sqlcnn<br/>
        ' ''' ・発行クエリより抽出した全データをDataTableへ全格納
        ' ''' </remarks>
        'Public Shared Function SqlRecSet( _
        '                                 ByVal sqlstr As String, _
        '                                 ByRef sqlcnn As System.Data.SqlClient.SqlConnection, _
        '                                 ByRef readtbl As DataTable
        '                                 ) As Integer

        '    Dim adater As New SqlDataAdapter()
        '    Dim time_sta As DateTime

        '    readtbl.Clear()                                                     '蓄積データクリア
        '    adater.SelectCommand = New SqlCommand(sqlstr, sqlcnn)
        '    adater.SelectCommand.CommandTimeout = LimitTimeOut                  'タイムアウト設定
        '    time_sta = Now                                                      '開始時間
        '    adater.Fill(readtbl)                                                'データ取得
        '    Debug.Print(Typ.ToStr(Now - TimeOfDay))                             '経過時間

        '    Return readtbl.Rows.Count

        'End Function

    End Class

#End Region

#Region "ログ出力設定クラス"

    Public Class LogSetting

        ''' <summary>
        ''' ログ出力テーブルCREATEクエリ生成
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LogTblCreateQry() As String

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " CREATE TABLE " & LOG_REL_TABLENAME
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & LOG_HEADER1 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER2 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER3 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER4 & " NCHAR(50), "
            tmp_sql = tmp_sql & LOG_HEADER5 & " NCHAR(200), "
            tmp_sql = tmp_sql & LOG_HEADER6 & " NCHAR(100), "
            tmp_sql = tmp_sql & LOG_HEADER7 & " NCHAR(100), "
            tmp_sql = tmp_sql & LOG_HEADER8 & " NCHAR(200), "
            tmp_sql = tmp_sql & LOG_HEADER9 & " NCHAR(200), "
            tmp_sql = tmp_sql & LOG_HEADER10 & " NCHAR(200), "
            tmp_sql = tmp_sql & LOG_HEADER11 & " NCHAR(200), "
            tmp_sql = tmp_sql & LOG_HEADER12 & " NCHAR(50), "
            tmp_sql = tmp_sql & "  "
            tmp_sql = tmp_sql & " 	); "

            Return tmp_sql

        End Function

        '2016.03.18構築中 -sta

        ''' <summary>
        ''' ログ出力テーブルINSERTクエリ生成
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LogTblInsertQry(ByVal logvalue As String) As String

            Dim tmp_sql As String = ""
            Dim taisyotbl As String = ""

            tmp_sql = " INSERT INTO " & LOG_REL_TABLENAME & " VALUES ("
            tmp_sql = tmp_sql & logvalue
            tmp_sql = tmp_sql & ");"

            Return tmp_sql

        End Function

        ''' <summary>
        ''' ログ出力テーブルINSERTクエリ生成 (テーブルをそのままINSERTする)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LogTblInsertQry() As String

            Dim tmp_sql As String = ""

            tmp_sql = " INSERT INTO " & LOG_MAIN_TABLENAME & " SELECT * FROM " & LOG_REL_TABLENAME & " ORDER BY [出力日時] ;"

            Return tmp_sql

        End Function

        ''' <summary>
        ''' ログ出力テーブルSELECTクエリ作成
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_LogTblSelectQry() As String

            Dim tmp_sql As String = ""
            Dim taisyotbl As String = ""

            taisyotbl = LOG_REL_TABLENAME

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
            tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER12 & ") AS " & LOG_HEADER12
            'If Dev_CVFlg Then                                                                       'kakaka メモ：TBL名は通常非表示にする
            '    tmp_sql = tmp_sql & "    ,RTRIM(" & LOG_HEADER12 & ") AS " & LOG_HEADER12
            'End If
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
            tmp_logitem(6) = taisyodata
            tmp_logitem(7) = hubigein
            tmp_logitem(8) = taisyo
            tmp_logitem(9) = beforechgvalue
            tmp_logitem(10) = afterchgvalue
            tmp_logitem(11) = tblname

            Select Case typeno
                Case 1  '処理開始
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_NAIYO_REL_STA
                Case 9  '処理終了
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_NAIYO_REL_END

                Case 11 '読込開始
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_NAIYO_READSTA
                Case 12 '読込エラー
                    tmp_logitem(1) = LOG_RUI_TYUUI
                    tmp_logitem(2) = LOG_NAIYO_ERR_MISMATCH
                Case 19 '読込終了
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_NAIYO_READEND

                Case 21  '出力開始
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_NAIYO_WRITE_STA
                Case 29  '出力終了
                    tmp_logitem(1) = LOG_RUI_RIREKI
                    tmp_logitem(2) = LOG_NAIYO_WRITE_END


            End Select

            For cntii = 0 To itemcnt - 1

                tmp_str = tmp_str & ",'" & tmp_logitem(cntii) & "'"

            Next

            rtn_logvalue = tmp_str.Remove(0, 1)

            Return rtn_logvalue

        End Function

        Public Shared Sub Set_Log_Value_KomkErr(ByVal relitem As String, ByVal taisyodata As String, ByVal errstr As String, ByRef log_sql As String)

            'エラーメッセージの分割
            Dim tmp_errstr() As String = errstr.Split("-")

            Dim syorikekka As String = tmp_errstr(0)
            Dim hubinaiyo As String = tmp_errstr(1)
            Dim taisyo As String = tmp_errstr(2)

            '最終的なログ文字列群を取得
            Dim log_totalstr As String = LogSetting.Set_LogValue(12, relitem, taisyodata, syorikekka, hubinaiyo, taisyo)

            '挿入用クエリへ加工
            log_sql = Get_LogTblInsertQry(log_totalstr)

        End Sub

        ' ''' <summary>
        ' ''' 移行時に生じたログ内容を各変数へ格納　　　　　　　　　　　　　　　
        ' ''' 別メソッド(同一クラス内)にて成形                                  
        ' ''' 挿入用クエリへ加工
        ' ''' ソートリストへ格納
        ' ''' </summary>
        ' ''' <param name="hash_log"></param>
        ' ''' <param name="hash_beforechk"></param>
        ' ''' <param name="hash_afterchk"></param>
        ' ''' <param name="taisyodata"></param>
        ' ''' <param name="tblname"></param>
        ' ''' <param name="logcnt"></param>
        ' ''' <param name="sortlist"></param>
        ' ''' <remarks></remarks>
        'Public Shared Sub Set_Log_Value_KomkErr(ByVal hash_log As Hashtable, ByVal hash_beforechk As Hashtable, ByVal hash_afterchk As Hashtable, _
        '                                        ByVal taisyodata As String, ByVal tblname As String, ByRef logcnt As Integer, _
        '                                        ByRef sortlist As SortedList(Of Integer, String))

        '    For Each logvalue In hash_log

        '        Dim log_key As String = logvalue.Key
        '        Dim log_value As String = logvalue.Value
        '        Dim log_totalstr As String = ""

        '        '処理項目を分割
        '        Dim tmp_syori() As String = log_key.Split("/")

        '        '日本語テーブル名を取得して「処理項目_大分類」へ格納
        '        Dim tmp_tblfldname As String = tmp_syori(0)                 '※テーブル名は共通なので配列のインデックスは任意の値
        '        Dim tmp_tblname() As String = tmp_tblfldname.Split("-")
        '        Dim syorikomk As String = Hash_TblName_AlphaToJp(tmp_tblname(0))

        '        '日本語フィールド名を取得して「処理項目_小分類」へ格納
        '        Dim tmp_syoriitem As String = ""
        '        For cntii = 0 To UBound(tmp_syori)
        '            Dim tmp_str As String = Hash_FldName_AlphaToJp(tmp_syori(cntii))
        '            tmp_syoriitem = tmp_syoriitem & "、" & tmp_str
        '        Next
        '        tmp_syoriitem = tmp_syoriitem.Remove(0, 1)
        '        Dim syoriitem As String = tmp_syoriitem

        '        '「-」で連結されたログ内容を分割
        '        Dim tmp_log() As String = log_value.Split("-")

        '        'ログの中身を各変数へ格納
        '        Dim syorikekka As String = tmp_log(0)
        '        Dim hubinaiyo As String = tmp_log(1)
        '        Dim taisyo As String = tmp_log(2)

        '        '調整前後の値を変数へ格納
        '        Dim tmp_before As String = hash_beforechk(tmp_tblfldname(1))
        '        Dim tmp_after As String = hash_afterchk(tmp_tblfldname(1))
        '        Dim tmp_empty As String = "-"
        '        Dim beforechgvalue As String = ""
        '        Dim afterchgvalue As String = ""

        '        If hash_afterchk.Count = 0 Then         '移行されなかった場合(チェック後ハッシュテーブルにデータが存在しない)は値を設定しない
        '            beforechgvalue = tmp_empty
        '            afterchgvalue = tmp_empty
        '        ElseIf tmp_before = tmp_after Then      '調整前後で値が同じ場合は値を設定しない
        '            beforechgvalue = tmp_empty
        '            afterchgvalue = tmp_empty
        '        ElseIf tmp_before <> tmp_after Then     '調整前後で値が異なる場合は値を設定する(出力時を考慮しカンマを全角に変換しておく)
        '            beforechgvalue = tmp_before.Replace(",", "，")
        '            afterchgvalue = tmp_after.Replace(",", "，")
        '        End If

        '        '最終的なログ文字列群を取得
        '        log_totalstr = Set_LogValue(42, syorikomk, syoriitem, syorikekka, taisyodata, hubinaiyo, taisyo, beforechgvalue, afterchgvalue, tblname)

        '        '挿入用クエリへ加工
        '        Dim log_sql As String = Get_LogTblInsertQry(log_totalstr)

        '        'ログカウントを更新してリストへ格納
        '        logcnt = logcnt + 1
        '        sortlist.Add(logcnt, log_sql)

        '    Next

        'End Sub

        ''' <summary>
        ''' ログファイル出力処理
        ''' </summary>
        ''' <param name="logfilepath">出力先パス</param>
        ''' <param name="logmsg">出力メッセージ</param>
        ''' <returns>True.正常 False.異常</returns>
        ''' <remarks></remarks>
        Public Function OutPutLog( _
                                  ByVal logfilepath As String, _
                                  ByRef logmsg() As String _
                                  ) As Boolean
            Dim tmpmsg As String
            Dim rtn As Boolean = True

            Dim outputtime As String = IIf(logmsg(1).Trim = "", "-", logmsg(1))
            Dim group As String = IIf(logmsg(1).Trim = "", "-", logmsg(2))
            Dim phase As String = IIf(logmsg(1).Trim = "", "-", logmsg(3))
            Dim result As String = IIf(logmsg(1).Trim = "", "-", logmsg(4))
            Dim category As String = IIf(logmsg(1).Trim = "", "-", logmsg(5))
            Dim data As String = IIf(logmsg(1).Trim = "", "-", logmsg(6))
            Dim cause As String = IIf(logmsg(1).Trim = "", "-", logmsg(7))
            Dim content As String = IIf(logmsg(1).Trim = "", "-", logmsg(8))
            Dim beforechgvalue As String = IIf(logmsg(1).Trim = "", "-", logmsg(9))
            Dim afterchgvalue As String = IIf(logmsg(1).Trim = "", "-", logmsg(10))
            Dim tbl As String = IIf(logmsg(1).Trim = "", "-", logmsg(11))
            Dim logno As String = IIf(logmsg(1).Trim = "", "-", logmsg(12))

            Try
                tmpmsg = outputtime & ","
                tmpmsg = tmpmsg & group & ","
                tmpmsg = tmpmsg & phase & ","
                tmpmsg = tmpmsg & result & ","
                tmpmsg = tmpmsg & category & ","
                tmpmsg = tmpmsg & data & ","
                tmpmsg = tmpmsg & cause & ","
                tmpmsg = tmpmsg & content & ","
                tmpmsg = tmpmsg & beforechgvalue & ","
                tmpmsg = tmpmsg & afterchgvalue & ","
                tmpmsg = tmpmsg & tbl & ","
                tmpmsg = tmpmsg & logno & ","
                Call FileMethod.FileWriteAppend(tmpmsg, logfilepath, True)

            Catch ex As Exception
                rtn = False

            End Try

            Return rtn

        End Function

    End Class

#End Region

#Region "ファイル処理クラス"

    ''' <summary>
    ''' 【ファイル処理クラス】
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FileMethod

        ''' <summary>
        ''' ファイル書込み処理 2015.07.06 sol レビュー後修正_レビュー№35
        ''' </summary>
        ''' <param name="str">出力文字列</param>
        ''' <param name="filepath">ファイルパス(例：C:\Temp\test.txt)</param>
        ''' <param name="append">ファイル追記 True.追記(Def) False.上書き</param>
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
        ''' <param name="logflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TblView_Output_CSV(ByVal sqlcnn As SqlConnection, ByVal csvfilepath As String, ByVal taisyoname As String, ByVal sortstr As String, ByVal errstr As String, _
                                                  Optional ByVal sql As String = "", Optional ByVal logflg As Boolean = False) As Boolean

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
            tmp_cmdstr = """" & tmp_sql & """" & outtype & csvfilepath & servername & catalogname & username & password & strtype

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
                errstr = LOG_HUBI_MAKELOGFILE & vbCrLf & "エラー内容：" & ex.Message
                rtn = False
            End Try

            Return rtn

        End Function

    End Class

#End Region

#Region "Excelファイル設定クラス"

    Public Class ExcelFileManager

        ' ''' <summary>
        ' ''' 書込時のExcelファイルオープン処理
        ' ''' </summary>
        ' ''' <param name="appli"></param>
        ' ''' <param name="wbook"></param>
        ' ''' <param name="wsheet"></param>
        ' ''' <param name="startrow"></param>
        ' ''' <param name="columncnt"></param>
        ' ''' <param name="maxrowcnt"></param>
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
        '                                        ByRef dirpath As String, _
        '                                        ByVal filename As String, _
        '                                        ByVal sheetname As String _
        '                                        ) As Boolean

        '    Dim rtn As Boolean = True
        '    Dim midfilepath As String = dirpath & "\" & filename & ".xlsx"

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
        '        Dim dataarea As String = startrow & ":" & maxrowcnt

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

        ''' <summary>
        ''' 書込時のExcelファイルオープン処理
        ''' </summary>
        ''' <param name="appli"></param>
        ''' <param name="wbook"></param>
        ''' <param name="wsheet"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_ExcelFile_WriteOpen( _
                                                ByRef appli As Excel.Application, _
                                                ByRef wbook As Excel.Workbook, _
                                                ByRef wsheet As Excel.Worksheet, _
                                                ByRef dirpath As String, _
                                                ByVal filename As String, _
                                                ByVal sheetname As String _
                                                ) As Boolean

            Dim rtn As Boolean = True
            Dim midfilepath As String = dirpath & "\" & filename & ".xlsx"

            Try
                'ファイル準備
                appli = CreateObject("Excel.Application")
                appli.Visible = False
                wbook = appli.Workbooks.Open(midfilepath)
                wsheet = wbook.Worksheets(sheetname)

                '書込開始行の設定
                Dim startrow As Integer = MIDFILE_READWRITE_ROW

                '既存データの確認(最大行の取得)
                Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count

                '既存データ範囲を取得
                Dim dataarea As String = startrow & ":" & maxrowcnt

                '初期化
                If maxrowcnt >= startrow Then
                    wsheet.Rows(dataarea).Delete()
                End If

                ''ヘッダの列数取得
                'columncnt = wsheet.UsedRange.Columns.Count

            Catch ex As Exception
                'エラー処理を行うこと
                Debug.WriteLine(ex.Message)
                rtn = False
            End Try

            Return rtn

        End Function

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appli)
                wsheet = Nothing
                wbook = Nothing
                appli = Nothing
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try

        End Sub

        ''' <summary>
        ''' 読込時のExcelファイルオープン処理
        ''' Optional hanyoflg を追加 '20161004 自社口座の口座種別取得処理の修正
        ''' </summary>
        ''' <param name="appli"></param>
        ''' <param name="wbook"></param>
        ''' <param name="wsheet"></param>
        ''' <param name="startrow"></param>
        ''' <param name="columncnt"></param>
        ''' <param name="maxrowcnt"></param>
        ''' <param name="rowcnt"></param>
        ''' <param name="dirpath"></param>
        ''' <param name="filename"></param>
        ''' <param name="sheetname"></param>
        ''' <param name="hanyoflg"></param>
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
                                                ByVal sheetname As String, _
                                                Optional ByVal hanyoflg As Boolean = False
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
                '20161004 自社口座の口座種別取得処理の修正 -chg sta
                'If startrow = 0 Then
                '    startrow = MIDFILE_READWRITE_ROW
                'End If
                If hanyoflg Then
                    startrow = BASEMIDFILE_READWRITE_ROW
                Else
                    startrow = MIDFILE_READWRITE_ROW
                End If
                '20161004 自社口座の口座種別取得処理の修正 -chg end
                '既存データの確認(ヘッダを含めた最大行の取得)
                maxrowcnt = wsheet.UsedRange.Rows.Count

                'ヘッダの列数取得
                columncnt = wsheet.UsedRange.Columns.Count

                '行数取得(全行 - 読込開始行 + 1)
                rowcnt = maxrowcnt - startrow + 1

            Catch ex As Exception
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appli)              'kakaka ロック明示的解放
                wsheet = Nothing
                wbook = Nothing
                appli = Nothing
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try

        End Sub

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
            Dim skipcnt() As String = Nothing                         '分割SKIPカウント値
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

        ' ''' <summary>
        ' ''' 英数字半角変換(一部記号を含む)
        ' ''' </summary>
        ' ''' <param name="str">対象文字列</param>
        ' ''' <returns>変換後文字列</returns>
        ' ''' <remarks>
        ' ''' ・変換文字列：０-９Ａ-Ｚａ-ｚ：，－　<br/>
        ' ''' </remarks>
        'Public Shared Function Chg_NumNarrow(ByVal str As String) As String

        '    Dim reg As Regex = New Regex("[０-９Ａ-Ｚａ-ｚ：，－　]+")
        '    Dim output As String = reg.Replace(str, AddressOf RepNarrow)
        '    Return output

        'End Function

        ' ''' <summary>
        ' ''' 半角文字列変換
        ' ''' </summary>
        ' ''' <param name="mat">一致文字列</param>
        ' ''' <returns>変換後文字列</returns>
        ' ''' <remarks></remarks>
        'Public Shared Function RepNarrow(ByVal mat As Match) As String

        '    Return Strings.StrConv(mat.Value, VbStrConv.Narrow, 0)

        'End Function

        ''' <summary>
        ''' 値が基準値の倍数か判断  2016.02.22 プログレスバーの表示修正
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="base"></param>
        ''' <param name="chgvalue"></param>
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
        ''' 数値範囲チェック
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="min"></param>
        ''' <param name="max"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_NumRange(ByRef value As String, ByVal min As Integer, ByVal max As Integer) As Boolean

            Dim rtn As Boolean = True

            If Not (value >= min And value <= max) Then
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 文字列成形 (照合用に文字列を統一する)
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Str_Shape(ByVal value As String) As String

            Dim rtn As String = ""
            Dim tmp_str As String = value.Trim

            '大文字変換
            tmp_str = StrConv(tmp_str, VbStrConv.Uppercase)

            '全角変換
            tmp_str = StrConv(tmp_str, VbStrConv.Wide)

            'ひらがな変換
            tmp_str = StrConv(tmp_str, VbStrConv.Hiragana)

            '空白除去
            tmp_str = tmp_str.Replace("　", "")

            rtn = tmp_str

            Return rtn

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
        ''' グリッドの列幅を調整 '20160725 列幅修正
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Shared Sub Set_DgvColumnSize(ByVal dgv As DataGridView)

            Dim colcnt As Integer = dgv.ColumnCount - 1
            Dim tmp_colsize(colcnt) As Integer
            Dim colsizesum As Integer = 0
            Dim dgvwidth As Integer = dgv.Width

            '列幅をヘッダーに合わせる
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader

            'ヘッダーに合わせた列幅を取得
            For cntjj = 0 To colcnt
                tmp_colsize(cntjj) = dgv.Columns(cntjj).Width
            Next

            '列幅の自動調整を解除(列幅を可変にするため)
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            '取得しておいた列幅を設定(幅が80(全角文字列5文字分)未満の場合は最低幅の値を設定する)
            Dim mincolsize As Integer = 80
            For cntjj = 0 To colcnt
                If tmp_colsize(cntjj) < mincolsize Then
                    dgv.Columns(cntjj).Width = mincolsize
                Else
                    dgv.Columns(cntjj).Width = tmp_colsize(cntjj)
                End If
                If dgv.Columns(cntjj).Visible Then
                    colsizesum = colsizesum + dgv.Columns(cntjj).Width
                End If
            Next

            '表示領域の幅がグリッド幅より小さい場合は一致するように調整
            If colsizesum < dgvwidth Then
                Dim addcolsize As Integer = Math.Ceiling((dgvwidth - colsizesum) / dgv.ColumnCount)
                For cntjj = 0 To colcnt
                    dgv.Columns(cntjj).Width = dgv.Columns(cntjj).Width + addcolsize
                Next
            End If

        End Sub

        ''' <summary>
        ''' ファイル/フォルダのフルパスを設定 '20161012 紐付ツール速度改善対応 -add
        ''' ※本体からそのまま引用
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

    End Class

#End Region

#Region "フィールド/移行値紐付けハッシュテーブル作成クラス"

    Public Class GetHashFldToValue

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
            min_str = "1900/01/01"
            max_str = "2100/12/31"

            Dim tmp_date_min As DateTime = DateTime.Parse(min_str)
            Dim tmp_date_max As DateTime = DateTime.Parse(max_str)
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
                Return rtn
            End If

            '日付変換チェック
            If Not (DateTime.TryParse(chkstr, chgdate)) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (tmp_date_min <= chgdate AndAlso chgdate <= tmp_date_max) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                rtn = False
                Return rtn
            End If

            '正常値はそのまま格納
            chkstrafter = chkstr

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
        ''' <param name="chkstrafter"></param>
        ''' <param name="min_str"></param>
        ''' <param name="max_str"></param>
        ''' <param name="defvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataNumeric(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim lng As Long = 0
            Dim rtn As Boolean = True
            Dim min As Long = Int64.Parse(min_str)
            Dim max As Long = Int64.Parse(max_str)
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
                Return rtn
            End If

            '数値変換チェック
            If Not (Int64.TryParse(chkstr, lng)) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= lng AndAlso lng <= max) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
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
        ''' <param name="chkstrafter"></param>
        ''' <param name="min_str"></param>
        ''' <param name="max_str"></param>
        ''' <param name="defvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataDecimal(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim dec As Decimal = 0
            Dim rtn As Boolean = True
            Dim min As Decimal = Decimal.Parse(min_str)
            Dim max As Decimal = Decimal.Parse(max_str)
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
                Return rtn
            End If

            '数値変換チェック
            If Not (Decimal.TryParse(chkstr, dec)) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= dec AndAlso dec <= max) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
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
            Dim min As Double = Double.Parse(min_str)
            Dim max As Double = Double.Parse(max_str)
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
                Return rtn
            End If

            '数値変換チェック
            If Not (Double.TryParse(chkstr, dbl)) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= dbl AndAlso dbl <= max) Then
                chkstrafter = def
                errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
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
        ''' <param name="chkstrafter"></param>
        ''' <param name="min_str"></param>
        ''' <param name="max_str"></param>
        ''' <param name="defvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMoney(ByVal chkstr As String, ByRef chkstrafter As String, ByVal min_str As String, ByVal max_str As String, ByVal defvalue As String, ByRef errstr As String) As Boolean

            Dim dec As Decimal = 0
            Dim rtn As Boolean = True
            Dim min As Decimal = Decimal.Parse(min_str)
            Dim max As Decimal = Decimal.Parse(max_str)
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
                Return rtn
            End If

            '半角変換
            tmp_str = EtcMethod.Get_LenB(EtcMethod.Chg_NumNarrow(chkstr))

            'カンマチェック
            If InStr(tmp_str, ",") Then
                tmp_str = Replace(tmp_str, ",", "")
            End If

            '数値変換チェック
            If Not (Decimal.TryParse(tmp_str, dec)) Then
                chkstrafter = def                                               'kakaka moneyのデフォルト値は0円？
                rtn = False
                Return rtn
            End If

            '範囲チェック
            If Not (min <= dec AndAlso dec <= max) Then
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

            '外字チェック
            '外字チェック処理
            '変換処理をして作業用変数へ格納(未実装)                       'kakaka 10の導入先がV7と同じはずなので、外字はクリアされると思うが、汎用を考えて入れておく

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

            Return rtn

        End Function

        ''' <summary>
        ''' 都道府県マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_Todofuken(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'マスタからコードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT ken_no FROM m_ken"
            Dim flg As Boolean = True
            Dim rtn As Boolean = True

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

        ''' <summary>
        ''' 市区町村マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_Sikucyoson(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'マスタからコードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT si_no FROM m_si "
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
        Public Shared Function Chk_DataMstExist_Kinyu(ByVal sqlcnnv10 As SqlConnection, ByVal value_kinyuno As String, ByVal value_kinyutenno As String, _
                                                      ByRef value_kinyuname As String, ByRef value_kinyutenname As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True
            Dim readtbl As New DataTable
            Dim reccnt As Integer
            Dim fldname As String
            Dim fldvalue As String
            Dim tmp_kinyuno As String = ""
            Dim tmp_kinyuname As String = ""
            Dim tmp_tenno As String = ""
            Dim tmp_tenname As String = ""
            Dim hash_kinyu As New Hashtable
            Dim hash_kinyuten As New Hashtable

            '20160517 FB構築に伴う自社口座情報の修正 -add sta
            If value_kinyuno = "" Then
                Return rtn
            End If
            '20160517 FB構築に伴う自社口座情報の修正 -add end

            '金融機関マスタ抽出クエリ
            Dim strsql As String = ""
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 KINYU.kinyu_no "
            strsql = strsql & " 	,KINYU.kinyu_name "
            strsql = strsql & " 	,TEN.kinyu_tenno "
            strsql = strsql & " 	,TEN.kinyu_tenname "
            strsql = strsql & " FROM m_kinyu_ten AS TEN "
            strsql = strsql & " LEFT JOIN m_kinyu AS KINYU ON TEN.kinyu_no = KINYU.kinyu_no "
            strsql = strsql & " ORDER BY KINYU.kinyu_no,TEN.kinyu_tenno "

            '金融機関情報取得
            reccnt = DBExec.Exec_DataTable(strsql, sqlcnnv10, readtbl)

            '金融機関マスタ取得開始
            For cntii As Integer = 0 To readtbl.Rows.Count - 1

                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                    '項目名取得
                    fldname = readtbl.Columns(cntjj).ColumnName

                    '登録値取得
                    fldvalue = readtbl.Rows(cntii).Item(cntjj).ToString.Trim

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
                If Not hash_kinyu.Contains(tmp_kinyuno) Then
                    hash_kinyu.Add(tmp_kinyuno, tmp_kinyuname)
                End If
                hash_kinyuten.Add(tmp_kinyuno & "-" & tmp_tenno, tmp_tenname)

            Next

            '金融機関照合
            If hash_kinyu.Contains(value_kinyuno) Then
                value_kinyuname = hash_kinyu(value_kinyuno)
            Else
                Dim errkinyu As String = "金融機関"
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & errkinyu & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                rtn = False
                Return rtn
            End If

            '金融機関支店照合

            '20160517 FB構築に伴う自社口座情報の修正 -add sta
            Dim tmp_chkkinyuno() As String = value_kinyutenno.Split("-")
            If tmp_chkkinyuno(1) = "" Then
                Return rtn
            End If
            '20160517 FB構築に伴う自社口座情報の修正 -add end

            If hash_kinyuten.Contains(value_kinyutenno) Then
                value_kinyutenname = hash_kinyuten(value_kinyutenno)
            Else
                Dim errkinyuten As String = "金融機関支店"
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & errkinyuten & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                rtn = False
                Return rtn
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 仲介業者マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_GyCyukai(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT gy_fudono FROM gydata_fudo "
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
        ''' 施工業者マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_GySeko(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT gy_sekono FROM gydata_seko "
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
        ''' 保守業者マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_GyHosyu(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT gy_sisetuno FROM gydata_sisetu "
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
        ''' 自社マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_Jisya(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT jisya_no FROM jisyadata "
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
        ''' 家賃入金口座マスタ有無チェック
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_YatinKoza(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            'コードを取得
            Dim tmp_list As New List(Of String)
            Dim tmp_sql As String = " SELECT yatin_kozano FROM m_yatinkoza "
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

        ''' <summary>
        ''' 都市開発/用途地域マスタ有無チェック
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="erritem"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_TosiYoto(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            If value = "99" Then
                errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                chkvalue = "-1"
                rtn = False
            Else
                chkvalue = value
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 紐付け項目有無チェック
        ''' </summary>
        ''' <param name="hash_rel"></param>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="erritem"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_RelItem(ByVal hash_rel As Hashtable, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            'ダミーレコードとして挿入する必要がある場合は「挿入用」として抽出しておいたデータをそのまま移行する
            '(送金ルール入金項目のその他請求等)
            If value = "挿入用" Then
                chkvalue = "0"
                Return rtn
            End If

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
        ''' 間取文字列有無チェック
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="erritem"></param>
        ''' <param name="errstr"></param>
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
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="erritem"></param>
        ''' <param name="errstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_DataMstExist_MadoriUtiwakeStr(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

            Dim rtn As Boolean = True

            If value <> "" Then

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
                        errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                        chkvalue = ""
                        rtn = False
                End Select

            End If

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
        ''' <param name="value"></param>
        ''' <param name="chkvalue"></param>
        ''' <param name="errstr"></param>
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

        ' ''' <summary>
        ' ''' 全体データチェック→チェックメソッド呼出
        ' ''' </summary>
        ' ''' <param name="tblfldname"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <remarks></remarks>
        'Public Shared Sub Call_ChkMethodMidTotal(ByVal tblfldname As String, ByVal value As String, ByRef errstr As String)

        '    Dim tmp_type As String = Hash_FiledTypeJp(tblfldname)
        '    Dim tmp_size As Integer = Int32.Parse(Hash_FiledSizeJp(tblfldname))
        '    Dim min As New Object
        '    Dim max As New Object
        '    Dim datemin As String = ""
        '    Dim datemax As String = ""

        '    Dim chkaftervalue As String = ""
        '    Dim defvalue As String = ""
        '    Dim normalflg As Boolean = True                     'kakaka 以下の条件を表など一覧にして仕様書に記載しておく

        '    Select Case tmp_type                                'kakaka 要確認：資料を見てチェックする(金丸)
        '        Case "int", "smallint", "bigint", "tinyint"
        '            Call Chg_SizeToValue(tmp_size, min, max)
        '            normalflg = Chk_DataNumeric(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
        '        Case "bit"
        '            normalflg = Chk_DataBit(value, chkaftervalue, defvalue, errstr)
        '        Case "char", "varchar", "text"
        '            normalflg = Chk_DataString(value, chkaftervalue, tmp_size, 0, errstr)
        '        Case "nchar", "nvarchar", "ntext", "sysname"
        '            normalflg = Chk_DataString(value, chkaftervalue, tmp_size, 1, errstr)
        '        Case "datetime", "smalldatetime", "date", "datetime2", "datetimeoffset"
        '            normalflg = Chk_DataDate(value, chkaftervalue, datemin, datemax, defvalue, errstr)
        '        Case "time"                                     'kakaka 注意：未実装
        '            normalflg = Chk_DataTime(value, chkaftervalue, errstr)
        '        Case "money"
        '            Call Chg_SizeToValue(tmp_size, min, max)
        '            normalflg = Chk_DataMoney(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
        '        Case "decimal", "numeric"
        '            Call Chg_SizeToValue(tmp_size, min, max)
        '            normalflg = Chk_DataDecimal(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
        '        Case "float", "real"
        '            normalflg = Chk_DataDouble(value, chkaftervalue, min.ToString, max.ToString, defvalue, errstr)
        '        Case "varbinary"

        '        Case "xml", "uniqueidentifier"
        '            'プログラム内で生成しているためチェック不要
        '    End Select

        '    'hash_err.Add(tblfldname, )



        'End Sub

        ' ''' <summary>
        ' ''' 個別データチェック→チェックメソッド呼出
        ' ''' </summary>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="type"></param>
        ' ''' <param name="min"></param>
        ' ''' <param name="max"></param>
        ' ''' <param name="defvalue"></param>
        ' ''' <param name="normalflg"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <remarks></remarks>
        'Public Shared Sub Call_ChkMethodPerItem(ByVal value As String, ByRef chkvalue As String, ByVal type As String, ByVal min As String, ByVal max As String, ByVal defvalue As String, ByRef normalflg As Boolean, ByRef errstr As String)

        '    '初期化                            'kakaka メモ：型チェック(ソースレビューでは全て必要なチェックがあることを前提とするので、確認しません)
        '    chkvalue = ""
        '    errstr = ""
        '    normalflg = True

        '    Select Case type
        '        Case "int", "smallint", "bigint", "tinyint"
        '            normalflg = Chk_DataNumeric(value, chkvalue, min, max, defvalue, errstr)
        '        Case "bit"
        '            normalflg = Chk_DataBit(value, chkvalue, defvalue, errstr)
        '        Case "char", "varchar", "text"
        '            normalflg = Chk_DataString(value, chkvalue, max, 0, errstr)
        '        Case "nchar", "nvarchar", "ntext", "sysname"
        '            normalflg = Chk_DataString(value, chkvalue, max, 1, errstr)
        '        Case "datetime", "smalldatetime", "date", "datetime2", "datetimeoffset"
        '            normalflg = Chk_DataDate(value, chkvalue, min, max, defvalue, errstr)
        '        Case "time"
        '            normalflg = Chk_DataTime(value, chkvalue, errstr)
        '        Case "money"
        '            normalflg = Chk_DataMoney(value, chkvalue, min, max, defvalue, errstr)
        '        Case "decimal", "numeric"
        '            normalflg = Chk_DataDecimal(value, chkvalue, min, max, defvalue, errstr)
        '        Case "float", "real"
        '            normalflg = Chk_DataDouble(value, chkvalue, min, max, defvalue, errstr)
        '        Case "varbinary"

        '        Case "xml", "uniqueidentifier"
        '            'プログラム内で生成しているためチェックせずそのまま返却
        '            chkvalue = value
        '    End Select

        'End Sub

        ' ''' <summary>
        ' ''' 有無参照データチェックメソッド呼出
        ' ''' </summary>
        ' ''' <param name="sqlcnnv10"></param>
        ' ''' <param name="value"></param>
        ' ''' <param name="chkvalue"></param>
        ' ''' <param name="chkkbn"></param>
        ' ''' <param name="normalflg"></param>
        ' ''' <param name="errstr"></param>
        ' ''' <remarks></remarks>
        'Public Shared Sub Call_ChkMethodPerItem_MstExist(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal chkkbn As String, ByRef normalflg As Boolean, ByRef errstr As String)

        '    '初期化
        '    chkvalue = ""
        '    errstr = ""
        '    normalflg = True

        '    Select Case chkkbn
        '        Case "都道府県"
        '            normalflg = Chk_DataMstExist_Todofuken(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "市区町村"
        '            normalflg = Chk_DataMstExist_Sikucyoson(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "金融機関"
        '            '構築中
        '            chkvalue = value
        '        Case "金融機関支店"
        '            '構築中
        '            chkvalue = value
        '        Case "仲介業者"
        '            normalflg = Chk_DataMstExist_GyCyukai(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "ライフライン(電気)", "ライフライン(上水)", "ライフライン(ガス)", "ライフライン(排水)", "ライフライン(灯油)", "ライフライン(その他1)", "ライフライン(その他2)", "ライフライン(その他3)"
        '            normalflg = Chk_DataMstExist_GyLifeline(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "家賃入金口座"
        '            normalflg = Chk_DataMstExist_YatinKoza(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "施工業者"
        '            normalflg = Chk_DataMstExist_GySeko(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "保守業者"
        '            normalflg = Chk_DataMstExist_GyHosyu(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "自社"
        '            normalflg = Chk_DataMstExist_Jisya(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "自社口座"
        '            '構築中
        '            chkvalue = value
        '        Case "家主"
        '            '構築中
        '            chkvalue = value
        '        Case "家主口座"
        '            '構築中
        '            chkvalue = value
        '        Case "契約者"
        '            '構築中
        '            chkvalue = value
        '        Case "都市計画・用途地域"
        '            normalflg = Chk_DataMstExist_TosiYoto(value, chkvalue, chkkbn, errstr)
        '        Case "エリア"
        '            normalflg = Chk_DataMstExist_Area(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "沿線"
        '            normalflg = Chk_DataMstExist_Ensen(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "駅"
        '            normalflg = Chk_DataMstExist_Eki(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "間取"
        '            normalflg = Chk_DataMstExist_MadoriStr(value, chkvalue, chkkbn, errstr)
        '        Case "間取内訳"
        '            normalflg = Chk_DataMstExist_MadoriUtiwakeStr(value, chkvalue, chkkbn, errstr)
        '        Case "契約分類"
        '            '構築中
        '            normalflg = Chk_DataMstExist_Kybrui(sqlcnnv10, value, chkvalue, chkkbn, errstr)
        '        Case "物件分類"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Bkrui, value, chkvalue, chkkbn, errstr)
        '        Case "構造"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Kozo, value, chkvalue, chkkbn, errstr)
        '        Case "部屋分類"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Hyrui, value, chkvalue, chkkbn, errstr)
        '        Case "パス"
        '            normalflg = Chk_DataExist_FilePath(value, chkvalue, errstr)
        '        Case "設備"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Setubi, value, chkvalue, chkkbn, errstr)
        '        Case "取引態様"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Toritaiyo, value, chkvalue, chkkbn, errstr)
        '        Case "口座種別"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Kozasyubetu, value, chkvalue, chkkbn, errstr)
        '        Case "入金項目"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Nkinkomk, value, chkvalue, chkkbn, errstr)
        '        Case "入金区分"
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Nkinkbn, value, chkvalue, chkkbn, errstr)
        '        Case "変動費マスタ"
        '            '構築中(変動費マスタを作成する必要がある)
        '            chkvalue = value
        '        Case "メーター分類"   '要紐付け作成(入金項目名→メーター分類)   ※仮で紐付けファイルだけ用意して検証
        '            normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Hendometer, value, chkvalue, chkkbn, errstr)
        '    End Select

        'End Sub

        ' ''' <summary>
        ' ''' 個別データチェックメイン処理(移行処理からの第一階層)
        ' ''' </summary>
        ' ''' <param name="tblname"></param>
        ' ''' <param name="keyvalue"></param>
        ' ''' <param name="list"></param>
        ' ''' <param name="hash_chkbefore"></param>
        ' ''' <param name="hash_chkafter"></param>
        ' ''' <param name="conditioncnt"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function Chk_DataPerItem(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByVal list_chkduplicate As List(Of String), ByVal list_basekeydata As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer, ByRef hash_log As Hashtable) As Boolean
        '    'kakaka ↑"keyvalue"は読み込んだセル値が渡されるが、どこで使用する？
        '    Dim rtn As Boolean = True
        '    Dim condflg As Boolean = False
        '    Dim tmp_mainkey As String = ""
        '    Dim midkeytblfldname(10) As String
        '    Dim midkeyvalue(10) As String

        '    '------------------------
        '    '各移行値チェック
        '    '------------------------

        '    For Each hashvalue In hash_chkbefore

        '        Dim tmp_key As String = hashvalue.Key                           '移行項目名 (アルファベット)                                              'kakaka 説明を入れてください。
        '        Dim get_key As String = tblname & "-" & tmp_key                 'テーブル名-フィールド名 (アルファベット)
        '        Dim tmp_value As String = hashvalue.Value                       '移行値
        '        Dim tmp_type As String = Hash_FiledTypeAlpha.Item(get_key)      'データ型
        '        Dim tmp_size_min As String = Hash_Min_Code.Item(get_key)        '最小値 (画面上)
        '        Dim tmp_size_max As String = Hash_Max_Code.Item(get_key)        '最大値 (画面上)
        '        Dim tmp_defvalue As String = Hash_DefaultValue.Item(get_key)    'デフォルト値
        '        Dim tmp_chkvalue As String = ""                                 'データチェック後の値の格納用
        '        Dim errstr As String = ""                                       'エラー文字列格納用 (ログ出力)

        '        Dim normalflg As Boolean = True
        '        Dim keyflg As Boolean = False
        '        Dim requiredflg As Boolean = False

        '        'キーフィールドの判別
        '        If Hash_MidKey_FieldAlpha.Contains(get_key) Then

        '            '2016.02.22 型落ちを考慮した処理へ修正 -chg sta
        '            'Dim tmp_keyno As Integer = Int32.Parse(Hash_MidKey_FieldAlpha.Item(get_key))        'kakaka 型落ちは大丈夫？
        '            'midkeytblfldname(tmp_keyno) = get_key
        '            'midkeyvalue(tmp_keyno) = tmp_value
        '            'keyflg = True
        '            Dim tmp_nostr As String = Hash_MidKey_FieldAlpha.Item(get_key)
        '            Dim tmp_keyno As Integer = 0
        '            If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        '                midkeytblfldname(tmp_keyno) = get_key
        '                midkeyvalue(tmp_keyno) = tmp_value
        '                keyflg = True
        '            End If
        '            '2016.02.22 型落ちを考慮した処理へ修正 -chg end

        '        End If

        '        'キーではないが登録する際に必須となる項目の判別(契約情報の契約開始日等)                 'kakaka メモ：実際の革命にて登録を行って確認した結果。型による一定のチェックとは別。
        '        If List_Required_Field.Contains(get_key) Then                                           'kakaka メモ：作業ファイルへ予め用意。それがセットされたもの
        '            requiredflg = True
        '        End If

        '        'データチェック                                                                        'kakaka メモ：型(共通)チェック
        '        Call DataChk.Call_ChkMethodPerItem(tmp_value, tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, normalflg, errstr)

        '        '参照マスタ有無チェック                                                                'kakaka わかり易い文言にして下さい(参照マスタチェック(設定元マスタがあるかどうか))
        '        If Hash_Mst_ReferenceAlpha.Contains(get_key) Then
        '            Dim chkkbn As String = Hash_Mst_ReferenceAlpha.Item(get_key)
        '            Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr)
        '        End If

        '        '2016.02.22 エラー処理をまとめる -chg sta
        '        '↓↓↓旧srcコメントアウト↓↓↓
        '        ''チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける                    'kakaka メモ：上記までの処理の結果で不備の場合、処理をしない
        '        'If (keyflg AndAlso Not normalflg) OrElse (keyflg AndAlso tmp_chkvalue = "") Then                   'kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
        '        '    errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
        '        '    hash_log.Add(get_key, errstr)                                                      'kakaka エラーの場合の処理を纏められたらまとめて下さい(ここと、この下)
        '        '    rtn = False
        '        '    Return rtn
        '        'End If

        '        ''チェックデータが必須項目かつ空の場合処理を抜ける
        '        'If requiredflg AndAlso tmp_chkvalue = "" Then
        '        '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
        '        '    hash_log.Add(get_key, errstr)
        '        '    rtn = False
        '        '    Return rtn
        '        'End If
        '        '↑↑↑旧srcコメントアウト↑↑↑
        '        If (keyflg AndAlso normalflg = False) OrElse (keyflg AndAlso tmp_chkvalue = "") Then            'チェックデータがキーかつエラー、またはキーかつ空の場合は移行不可とする
        '            errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
        '            rtn = False
        '        ElseIf requiredflg AndAlso tmp_chkvalue = "" Then                                           'チェックデータが必須項目かつ空の場合は移行不可とする
        '            errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
        '            rtn = False
        '        End If
        '        If rtn = False AndAlso errstr <> "" Then
        '            hash_log.Add(get_key, errstr)
        '            Return rtn
        '        End If
        '        '2016.02.22 エラー処理をまとめる -chg end

        '        '調整された値が存在する場合の処理
        '        If normalflg = False Then                                                                'kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
        '            '調整フラグをONにする
        '            condflg = True
        '            'エラー内容をオブジェクトへ格納
        '            'hash_log.Add(get_key, errstr)
        '        End If

        '        'チェックした値をチェック済み格納用ハッシュテーブルへ格納
        '        hash_chkafter.Add(tmp_key, tmp_chkvalue)

        '    Next

        '    '------------------------
        '    '重複チェック
        '    '------------------------

        '    Dim tmp_tblfldvalue As String = ""
        '    Dim tmp_keyvalue As String = ""

        '    '退避しておいたキーフィールドと値を取得
        '    For cntii = 1 To UBound(midkeyvalue)
        '        If midkeyvalue(cntii) <> "" Then
        '            tmp_tblfldvalue = tmp_tblfldvalue & "/" & midkeytblfldname(cntii)
        '            tmp_keyvalue = tmp_keyvalue & "-" & midkeyvalue(cntii)
        '        End If
        '    Next
        '    tmp_tblfldvalue = tmp_tblfldvalue.Remove(0, 1)                                          'kakaka メモ：先頭文字除去
        '    tmp_keyvalue = tmp_keyvalue.Remove(0, 1)

        '    '照合
        '    If list_chkduplicate.Contains(tmp_keyvalue) Then                                        'kakaka メモ：先で貯めこんでいた重複データを比較
        '        hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
        '        hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
        '        rtn = False
        '        Return rtn
        '    End If

        '    '------------------------
        '    '親マスタ有無チェック
        '    '------------------------

        '    '照合
        '    If list_basekeydata.Count <> 0 Then
        '        If Not list_basekeydata.Contains(midkeyvalue(1)) Then
        '            hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
        '            hash_log.Add(midkeytblfldname(1), LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE)
        '            rtn = False
        '            Return rtn
        '        End If
        '    End If

        '    '--------------------------------------------
        '    '調整フラグがTrueの場合、調整件数を追加する
        '    '--------------------------------------------

        '    If condflg Then
        '        conditioncnt = conditioncnt + 1
        '    End If

        '    Return rtn

        'End Function

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

    End Class

#End Region

#Region "プログレスバー表示設定クラス"

    Public Class ProgressBarManager

        ''' <summary>
        ''' 読込時プログレスバー初期設定
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbInit(ByVal int As Integer)

            With Njc.Frm.RelationFrm.pgbRelItemRead
                .Minimum = 0
                .Value = 0
                .Maximum = int
            End With

        End Sub

        ''' <summary>
        ''' 読込時プログレスバー更新
        ''' </summary>
        ''' <param name="int"></param>
        ''' <remarks></remarks>
        Public Sub pgbsetting(ByVal int As Integer)

            With Njc.Frm.RelationFrm.pgbRelItemRead
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

    End Class

#End Region

End Namespace
