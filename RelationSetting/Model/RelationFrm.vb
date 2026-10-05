Imports RelationSetting.Njc.N3Lib.Utys
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Repository
Imports System.IO
Imports Microsoft.Win32
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports RelationSetting.Njc.Query

Namespace Njc.Frm

    Public Class RelationFrm

#Region "宣言"

        Public sqlcnnV7 As System.Data.SqlClient.SqlConnection
        Public sqlcnnV10 As System.Data.SqlClient.SqlConnection
        Public fstmodelv7 As New Njc.Model.DefSQLConnection       'V7用Model格納変数
        Public fstmodelv10 As New Njc.Model.DefSQLConnection      'V10用Model格納変数

        Private cnnV7 As New Njc.Common.DBConnection                        'V7用接続
        Private cnnV10 As New Njc.Common.DBConnection                       'V10用接続
        Private tabPageManager_Main As Njc.Common.TabPageManager            'タブページ表示用(メインタブ用)
        Private tabPageManager_Sub As Njc.Common.TabPageManager             'タブページ表示用(設定箇所用)
        Private logset As New Njc.Common.LogSetting                         'ログ出力用
        Private tblname = New Njc.Model.TblName               'テーブル名(日本語)
        Private tmptblname = New Njc.Model.TmpTableName       '仮テーブル名

        Private blnRtn As Boolean
        Private gridchgflg As Boolean = False
        Private tabchgflg As Boolean = False
        Private taihi_nkbnname As String = ""
        Private taihi_nkbnzkseiname As String = ""
        Private taihi_nkinname As String = ""
        Private taihi_nkinzkseiname As String = ""
        Private taihi_setubigrp As String = ""
        Private taihi_setubimst As String = ""
        Private taihi_setubikomk As String = ""

        Private cellchg As Boolean = False

        Private cmdline() As String             '本体からの引数格納
        Private singlesta As Boolean = True    '単体起動判別用 True:単体起動 / False:本体呼び出し        

        Private hash_relitemtogrid As New Hashtable
        Private hash_gridtorelitem As New Hashtable
        Private hash_relitemtomodel As New Hashtable
        Private hash_relitemtorepository As New Hashtable
        Private hash_gridtocolno As New Hashtable

        '2016.04.11 呼出起動の修正 -add sta
        Private list_relitem_singlesta As New List(Of String)       '単体起動時に選択された紐付項目を取得して格納
        Private list_relitem_call As New List(Of String)            '呼出起動時にコマンドラインに含まれる紐付関連項目から紐付項目を取得して格納
        '2016.04.11 呼出起動の修正 -add end

        Private beforecolor_V7 As New Color           '20160525 全体的な動作の修正
        Private beforecolor_10 As New Color           '20160525 全体的な動作の修正

        '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add
        Public chgcolorrow As New List(Of Integer)


        'テスト用引数
        '0 1 True,PC-1KA_SOL\SQL2K8,fk5dtsql,sa,p7s2#c1j3n True,PC-1KA_SOL\SQL2012,fk8db_free,sa,p7s2#c1j3n,15 C:\Converter10\RelationSetting\relationfile C:\Converter10\middlefile bkruichk-hyruichk-nkinkbnchk-toritaiyochk-nkinkomkchk-kozochk-kozasyuchk-setubichk-kagichk-jisyakozachk-tosiyotochk-kyruichk-gazochk-fbfmtchk-sohokenchk

#End Region

#Region "初期設定"

        ''' <summary>
        ''' 起動
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

            '20160926 紐付ツール呼出処理の仕様変更 -add sta
            '本体から呼び出された際の引数を格納
            cmdline = System.Environment.GetCommandLineArgs()

            '呼出起動確認
            If UBound(cmdline) >= 1 Then
                If cmdline(1) = 0 Then
                    singlesta = False
                Else
                    '呼出以外は終了する
                    Me.Close()
                    Exit Sub
                End If
            Else
                '呼出以外は終了する
                Me.Close()
                Exit Sub
            End If
            '20160926 紐付ツール呼出処理の仕様変更 -add end

            '初期設定
            Call Initialize()

        End Sub

        ''' <summary>
        ''' 初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Initialize()

            '2016.04.11 呼出起動の修正 -del sta
            '初期設定のその他へ移動させる
            ''本体から呼び出された際の引数を格納
            'cmdline = System.Environment.GetCommandLineArgs()
            'If UBound(cmdline) >= 1 Then
            '    singlesta = False
            'End If

            ''呼出起動時のコマンドライン取得
            'If singlesta = False Then
            '    Call cnnV7.CnnSession(fstmodelv7, sqlcnnV7, True)
            '    Call cnnV10.CnnSession(fstmodelv10, sqlcnnV10, True)
            '    Call Me.Rel_ReadData()
            '    cellchg = True
            '    '紐付ファイル/ログファイル格納先取得
            '    RelDirPath = cmdline(12)
            '    'LogDirPath = cmdline(13)
            '    'test
            '    'RelationDirPath = "C:\Converter10\File\relationfile"
            '    LogDirPath = "C:\Converter10\File\rellog"
            'End If
            '2016.04.11 呼出起動の修正 -del end

            '画面初期設定
            Me.IniForm()

            '各タブ初期設定
            Me.IniFormTabPage1()            '初期設定タブ
            Me.IniFormTabPage2()            'DB接続タブ
            Me.IniFormTabPage3()            '対象項目選択タブ
            Me.IniFormTabPage4(True)        '紐付設定タブ
            Me.IniFormTabPage5()            '終了タブ

            'その他
            Me.IniFormEtc()

            Call Set_TabOrder_Main(0)
            tabchgflg = True

        End Sub

        ''' <summary>
        ''' メイン画面設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniForm()

            '20160725_改善対応 add -sta
            'タブ隠し(TAB3)                             
            Me.lblHidden3.Width = 944
            Me.lblHidden3.Height = 55
            '20160725_改善対応 add -end

            'フォームコントロールボックス非可視設定
            Me.ControlBox = False

            '2016.03.23 フォームサイズ固定化 -add
            Me.FormBorderStyle = FormBorderStyle.FixedSingle

            '20160830 鍵紐付箇所の削除 -add
            grpKagicntgrp.Visible = False

        End Sub

        ''' <summary>
        ''' 初期設定タブ初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormTabPage1()

            'ラベル
            'Me.lblCaution_old.Visible = False              '20160725 汎用CVK対応 del

            'テキストボックス
            '2016.04.11 呼出起動の修正 -chg sta
            'Me.txtMidDirPath.Text = "C:\Converter10\middlefile"
            'Me.txtRelationDirPath.Text = "C:\Converter10\RelationSetting\relationfile"
            'Me.txtLogDirPath.Text = "C:\Converter10\RelationSetting\log"

            '実行ファイルパス、各関連ファイル格納フォルダパス取得
            Dim exepath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim exedir As String = IO.Path.GetDirectoryName(exepath)

            '各関連フォルダパス取得
            Dim tmp_str As String = exedir.Replace(DIR_RELEXEDIR_NAME, "")
            Dim tmp_defmiddirpath As String = Me.Set_Path(tmp_str, DIR_MID_NAME)    '中間ファイル格納パス
            Dim tmp_defreldirpath As String = Me.Set_Path(exedir, DIR_REL_NAME)    '紐付ファイル格納パス
            Dim tmp_defrellogdirpath As String = Me.Set_Path(exedir, DIR_RELLOG_NAME)    '紐付ファイル格納パス

            Me.txtMidDirPath.Text = tmp_defmiddirpath
            Me.txtRelationDirPath.Text = tmp_defreldirpath
            Me.txtLogDirPath.Text = tmp_defrellogdirpath

            '2016.04.11 呼出起動の修正 -chg end
            Me.txtMidDirPath.Enabled = False
            Me.txtRelationDirPath.Enabled = False
            Me.txtLogDirPath.Enabled = False
            Me.txtMidDirPath.BackColor = Color.White
            Me.txtLogDirPath.BackColor = Color.White
            Me.txtRelationDirPath.BackColor = Color.White

        End Sub

        ''' <summary>
        ''' DB接続タブ初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormTabPage2()

            '接続情報初期設定
            Call Me.Set_V7DBInfo()
            Call Me.Set_10DBInfo()
            Me.Set_DefConInfo_To_Control(fstmodelv7, fstmodelv10)

            'パスワード入力ボックスマスク設定
            Me.txtV7Pass.PasswordChar = "*"c
            Me.txtV10Pass.PasswordChar = "*"c

            'ネットワークライブラリはとりあえず不可視にしておく
            Me.lblNetworklib.Visible = False
            Me.txtV7Networklib.Visible = False
            Me.txtV10Networklib.Visible = False

        End Sub

        ''' <summary>
        ''' 移行項目選択タブ初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormTabPage3()

            Dim container As New Control
            Dim controls As New List(Of Control)

            container = tabPage2

            controls.Add(Me.chkBkrui)
            controls.Add(Me.chkHyrui)
            controls.Add(Me.chkNkinkbn)
            controls.Add(Me.chkToritaiyo)
            controls.Add(Me.chkNkinkomk)
            controls.Add(Me.chkKozo)
            controls.Add(Me.chkKozasyu)
            controls.Add(Me.chkSetubi)
            controls.Add(Me.chkKagi)
            controls.Add(Me.chkJisyakoza)
            controls.Add(Me.chkGazo)        '20160523 画像紐付設定処理の追加 -add
            controls.Add(Me.chkFBFmt)       '20160525 FBフォーマット紐付設定処理の追加 -add

            For i As Integer = 0 To controls.Count - 1
                container.Controls.SetChildIndex(controls(i), i)        'kakaka 20160723
            Next
            '20160905 紐付ツールの未使用グリッド非表示処理 -add
            Me.DataGridView1.Visible = False

        End Sub

        ''' <summary>
        ''' 紐付設定タブ初期設定
        ''' </summary>
        ''' <param name="initflg"></param>
        ''' <remarks></remarks>
        Private Sub IniFormTabPage4(ByVal initflg As Boolean)

            '20160725 紐付設定改善 add -sta
            'If initflg Then
            '    tabPageManager_Sub = New Njc.Common.TabPageManager(tabCtrlRelwork)
            'End If

            ''起動時は全て非表示にしておく (選択された項目のみ表示させるため)
            'tabPageManager_Sub.ChangeTabPageVisible(0, False)
            'tabPageManager_Sub.ChangeTabPageVisible(1, False)
            'tabPageManager_Sub.ChangeTabPageVisible(2, False)
            'tabPageManager_Sub.ChangeTabPageVisible(3, False)
            'tabPageManager_Sub.ChangeTabPageVisible(4, False)
            'tabPageManager_Sub.ChangeTabPageVisible(5, False)
            'tabPageManager_Sub.ChangeTabPageVisible(6, False)
            'tabPageManager_Sub.ChangeTabPageVisible(7, False)
            'tabPageManager_Sub.ChangeTabPageVisible(8, False)
            'tabPageManager_Sub.ChangeTabPageVisible(9, False)
            'tabPageManager_Sub.ChangeTabPageVisible(10, False)  '20160523 画像紐付設定処理の追加 -add
            'tabPageManager_Sub.ChangeTabPageVisible(11, False)  '20160525 FBフォーマット紐付設定処理の追加 -add
            'tabPageManager_Sub.ChangeTabPageVisible(12, False)  '20160526 契約分類マスタの追加 -add
            'tabPageManager_Sub.ChangeTabPageVisible(13, False)  '20160531 都市計画用途地域の追加 -add
            'タブ隠し(TAB3)                             
            Me.lblHidden4.Width = 903
            Me.lblHidden4.Height = 26
            '非表示設定
            Call Chg_TabPageSub(initflg)
            '20160725 紐付設定改善 add -end

            '20160622 入金項目No任意コード一括設定 -add
            Me.txtNkinnoSta.MaxLength = 3

            '20160701 指摘事項まとめファイルの対応 ラベル非表示対応 -add
            Me.lblCaution.Visible = False
            Me.Label40.Visible = False  '20160825 ラベル表示制御対応 -add
            Me.Label37.Visible = False  '20160825 ラベル表示制御対応 -add

        End Sub

        ''' <summary>
        ''' 終了タブ初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormTabPage5()



        End Sub

        ''' <summary>
        ''' その他初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub IniFormEtc()

            '2016.04.11 呼出起動の修正 -chg sta
            ''ボタンの活性制御
            'If singlesta = False Then
            '    '呼出起動時
            '    'ボタンの活性制御
            '    Call Chg_BtnStatus(3)
            'Else
            '    '単体起動時
            '    'ボタンの活性制御
            '    Call Chg_BtnStatus(0)
            'End If

            ''TabPageManagerオブジェクト生成
            'tabPageManager_Main = New Njc.Common.TabPageManager(tabCtrlRelMain)

            ''TAB非表示設定
            'If singlesta = False Then
            '    '呼出起動時
            '    tabPageManager_Main.ChangeTabPageVisible(0, False)
            '    tabPageManager_Main.ChangeTabPageVisible(1, False)
            '    tabPageManager_Main.ChangeTabPageVisible(3, False)
            'Else
            '    '単体起動時
            '    tabPageManager_Main.ChangeTabPageVisible(1, False)
            '    tabPageManager_Main.ChangeTabPageVisible(2, False)
            '    tabPageManager_Main.ChangeTabPageVisible(3, False)
            '    tabPageManager_Main.ChangeTabPageVisible(4, False)
            'End If

            'TabPageManagerオブジェクト生成
            tabPageManager_Main = New Njc.Common.TabPageManager(tabCtrlRelMain)

            '20160926 紐付ツール呼出処理の仕様変更 -del sta
            ''本体から呼び出された際の引数を格納
            'cmdline = System.Environment.GetCommandLineArgs()

            ''20160725 紐付設定改善 chg -sta
            ''If UBound(cmdline) >= 1 Then
            ''    singlesta = False
            ''End If
            ''単独起動選択(0.呼出起動, 0以外.単体起動)
            'If cmdline(1) = 0 Then
            '    singlesta = False
            'End If
            ''20160725 紐付設定改善 chg -end
            '20160926 紐付ツール呼出処理の仕様変更 -del end

            If singlesta = False Then

                '受け取ったコマンドラインを各オブジェクトへ格納
                Call Me.Set_CmdlineInfoToObj()

                'タブ制御
                tabPageManager_Main.ChangeTabPageVisible(0, False)
                tabPageManager_Main.ChangeTabPageVisible(1, False)
                tabPageManager_Main.ChangeTabPageVisible(2, False)
                tabPageManager_Main.ChangeTabPageVisible(4, False)
                '20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_紐付 -add sta
                '画面制御
                If CNVNO = ConvertTypes._汎用 Then
                    Call Me.Set_HanyoGamen()
                End If
                '20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_紐付 -add end
                'ボタン遷移
                Call Chg_BtnStatus(35)

                'データ読込
                Call Me.Rel_ReadData()

                'ボタン遷移
                Call Chg_BtnStatus(36)

            Else
                '単体起動時

                'タブ制御
                tabPageManager_Main.ChangeTabPageVisible(1, False)
                tabPageManager_Main.ChangeTabPageVisible(2, False)
                tabPageManager_Main.ChangeTabPageVisible(3, False)
                tabPageManager_Main.ChangeTabPageVisible(4, False)

                'ボタン遷移
                Call Chg_BtnStatus(0)
            End If

            '2016.04.11 呼出起動の修正 -chg end

        End Sub

#End Region

#Region "イベント処理"

        ''' <summary>
        ''' イベント処理：戻るボタンクリック
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

            Dim seltab As String = Me.tabCtrlRelMain.SelectedTab.Name

            Select Case seltab
                Case "tabPage0"  'なし

                Case "tabPage1"  '戻る
                    tabCtrlRelMain.SelectedTab = tabPage0
                    tabPageManager_Main.ChangeTabPageVisible(0, True)
                    tabPageManager_Main.ChangeTabPageVisible(1, False)
                    Call Chg_BtnStatus(0)
                Case "tabPage2"  '戻る
                    tabCtrlRelMain.SelectedTab = tabPage1
                    tabPageManager_Main.ChangeTabPageVisible(1, True)
                    tabPageManager_Main.ChangeTabPageVisible(2, False)
                    Call Chg_BtnStatus(1)
                Case "tabPage3"  '戻る
                    tabCtrlRelMain.SelectedTab = tabPage2
                    tabPageManager_Main.ChangeTabPageVisible(2, True)
                    tabPageManager_Main.ChangeTabPageVisible(3, False)
                    cellchg = False
                    Call Chg_BtnStatus(10)
                Case "tabPage4"  'なし

                Case Else

            End Select
        End Sub

        ''' <summary>
        ''' イベント処理：次へボタンクリック
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

            Dim seltab As String = Me.tabCtrlRelMain.SelectedTab.Name

            Select Case seltab
                Case "tabPage0"  '次へ
                    tabCtrlRelMain.SelectedTab = tabPage1
                    Call Set_TabOrder_Main(1)
                    tabPageManager_Main.ChangeTabPageVisible(1, True)
                    tabPageManager_Main.ChangeTabPageVisible(0, False)
                    Call Chg_BtnStatus(1)
                Case "tabPage1"  '次へ
                    tabCtrlRelMain.SelectedTab = tabPage2
                    Call Set_TabOrder_Main(2)
                    tabPageManager_Main.ChangeTabPageVisible(2, True)
                    tabPageManager_Main.ChangeTabPageVisible(1, False)
                    Call Chg_BtnStatus(10)
                Case "tabPage2"  '読込

                    'データ読込前処理
                    If Me.Rel_ReadDataBefore() = False Then
                        Exit Sub
                    End If

                    'ボタン遷移
                    '2016.04.11 呼出起動の修正 -del
                    'Call Chg_BtnStatus(32)

                    '画面/ボタン遷移
                    tabCtrlRelMain.SelectedTab = tabPage3
                    Call Set_TabOrder_Main(3)
                    tabPageManager_Main.ChangeTabPageVisible(3, True)
                    tabPageManager_Main.ChangeTabPageVisible(2, False)
                    '2016.04.11 呼出起動の修正 -add
                    Call Chg_BtnStatus(35)

                    'データ読込処理
                    Call Me.Rel_ReadData()

                    '画面/ボタン遷移
                    '2016.04.11 呼出起動の修正 -chg sta
                    'Call Chg_BtnStatus(3)
                    If CancelFlg = False Then
                        '中断されていない場合はボタン状態変更 (画面は遷移済み)
                        Call Chg_BtnStatus(36)
                    Else
                        '中断されている場合は終了タブへ遷移
                        Call Set_TabOrder_Main(3)
                        tabCtrlRelMain.SelectedTab = tabPage4
                        tabPageManager_Main.ChangeTabPageVisible(4, True)
                        tabPageManager_Main.ChangeTabPageVisible(3, False)
                        Call Chg_BtnStatus(4)
                    End If
                    '2016.04.11 呼出起動の修正 -chg end

                Case "tabPage3"  '次へ
                    Select Case Replace(btnNext.Text, " ", "")
                        Case "実行"

                            'データ書込前処理
                            If Me.Rel_WriteDataBefore() = False Then
                                Exit Sub
                            End If

                            'ボタン遷移
                            '2016.04.11 呼出起動の修正 -chg sta
                            'Call Chg_BtnStatus(32)
                            Call Chg_BtnStatus(37)
                            '2016.04.11 呼出起動の修正 -chg end

                            'データ書込処理
                            Call Me.Rel_WriteData()

                            '2016.04.11 呼出起動の修正 -add sta
                            '単体/呼出起動の処理の分岐があるため上部へ移動
                            'ログ出力
                            Dim logdropflg As Boolean = Me.chkLogTblDrop.Checked        'ログテーブル削除フラグ(開発用)
                            Call Me.Set_LogFile(logdropflg)
                            '2016.04.11 呼出起動の修正 -add end

                            '画面/ボタン遷移
                            '2016.04.11 呼出起動の修正 -chg sta
                            'Call Chg_BtnStatus(32)
                            'Call Set_TabOrder_Main(3)
                            'tabCtrlRelMain.SelectedTab = tabPage4
                            'tabPageManager_Main.ChangeTabPageVisible(4, True)
                            'tabPageManager_Main.ChangeTabPageVisible(3, False)
                            'Call Chg_BtnStatus(4)
                            If singlesta Then
                                Call Chg_BtnStatus(32)
                                Call Set_TabOrder_Main(3)
                                tabCtrlRelMain.SelectedTab = tabPage4
                                tabPageManager_Main.ChangeTabPageVisible(4, True)
                                tabPageManager_Main.ChangeTabPageVisible(3, False)
                                Call Chg_BtnStatus(4)
                            Else
                                '20160905 紐付設定終了時の完了メッセージ除去 -del sta
                                ''20160825 完了メッセージを自動で閉じる処理を追加 -chg sta
                                ''MsgResult = MessageBox.Show(MSG_END_CALL, "確認", _
                                ''                     MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                                'Dim result As DialogResult
                                'result = CreateObject("WScript.Shell").Popup(MSG_END_CALL, 15, "確認", MessageBoxButtons.OK + MessageBoxIcon.Asterisk)
                                ''20160825 完了メッセージを自動で閉じる処理を追加 -chg end
                                '20160905 紐付設定終了時の完了メッセージ除去 -del end
                                '20160707 本体と紐付ツールの中断を同期させる処理の追加 -add
                                Environment.ExitCode = 0

                                Me.Close()
                            End If
                            '2016.04.11 呼出起動の修正 -chg end

                            '2016.04.11 呼出起動の修正 -del sta
                            '単体/呼出起動の処理の分岐があるため上部へ移動
                            ''ログ出力
                            'Dim logdropflg As Boolean = Me.chkLogTblDrop.Checked        'ログテーブル削除フラグ(開発用)
                            'Call Me.Set_LogFile(logdropflg)
                            '2016.04.11 呼出起動の修正 -del end

                    End Select
                Case "tabPage4"
                    System.Diagnostics.Process.Start(LogFilePath)
                Case Else

            End Select
        End Sub

        ''' <summary>
        ''' イベント処理：終了ボタンクリック
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click

            Dim closeflg As Boolean = False

            '終了・中断処理
            Select Case Replace(btnEnd.Text, " ", "")
                Case "中止"
                    MsgResult = MessageBox.Show(MSG_STOP_A, "確認", _
                                            MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.Yes Then
                        tabCtrlRelMain.SelectedTab = tabPage3
                        lblFinaltxt.Text = MSG_STOP_B
                        closeflg = True

                        '20160707 本体と紐付ツールの中断を同期させる処理の追加 -add
                        Environment.ExitCode = 1

                    End If

                Case "終了"
                    MsgResult = MessageBox.Show(MSG_SYURYO_B, "確認", _
                                                 MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    closeflg = True
                Case "キャンセル"
                    '処理中断
                    MsgResult = MessageBox.Show(MSG_CANCEL_A, "確認", _
                                               MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                    If MsgResult = DialogResult.Yes Then
                        '2016.04.11 呼出起動の修正 -chg sta
                        'tabCtrlRelMain.SelectedTab = tabPage3
                        'lblFinaltxt.Text = MSG_CANCEL_B
                        'tabCtrlRelMain.SelectedTab = tabPage3
                        'tabPageManager_Main.ChangeTabPageVisible(4, True)
                        'tabPageManager_Main.ChangeTabPageVisible(3, False)
                        'Call Chg_BtnStatus(4)
                        CancelFlg = True
                        lblFinaltxt.Text = MSG_STOP_B
                        If singlesta = False Then
                            MsgResult = MessageBox.Show(MSG_CANCEL_B, "確認", _
                                                 MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                        End If
                        '2016.04.11 呼出起動の修正 -chg end

                        '20160707 本体と紐付ツールの中断を同期させる処理の追加 -add
                        Environment.ExitCode = 1

                    Else
                        '処理再開
                    End If
                Case Else

            End Select

            '終了・中断処理
            If closeflg Then
                cnnV7.CnnClose(sqlcnnV7)
                cnnV10.CnnClose(sqlcnnV10)
                Me.Close()
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：参照ボタン(紐付ファイル出力先フォルダ指定)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub btnDirSeach_Click(sender As Object, e As EventArgs) Handles btnMidDirSeach.Click, btnRelDirSeach.Click, btnLogDirSeach.Click

            Dim dirdialog As New FolderBrowserDialog

            '上部に表示する説明テキストを指定する
            dirdialog.Description = "フォルダを指定してください。"

            'ルートフォルダ指定(デフォルトでDesktop)
            dirdialog.RootFolder = Environment.SpecialFolder.Desktop

            '最初に選択するフォルダを指定する
            Select Case True
                Case sender Is Me.btnMidDirSeach
                    If Me.txtMidDirPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtMidDirPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnRelDirSeach
                    If Me.txtRelationDirPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtRelationDirPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                Case sender Is Me.btnLogDirSeach
                    If Me.txtLogDirPath.Text <> "" Then
                        dirdialog.SelectedPath = Me.txtLogDirPath.Text
                    Else
                        dirdialog.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If

            End Select

            'ユーザーが新しいフォルダを作成できるようにする
            dirdialog.ShowNewFolderButton = True

            'ダイアログを表示する
            If dirdialog.ShowDialog(Me) = DialogResult.OK Then
                Select Case True
                    Case sender Is Me.btnMidDirSeach        '20160525 全体的な動作の修正 -chg
                        Me.txtMidDirPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnLogDirSeach
                        Me.txtLogDirPath.Text = dirdialog.SelectedPath
                    Case sender Is Me.btnRelDirSeach
                        Me.txtRelationDirPath.Text = dirdialog.SelectedPath
                End Select
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：接続テスト
        ''' </summary>
        ''' <remarks>
        ''' ・接続テストが成功した場合のみ、次の処理へ進めるようにする
        ''' </remarks>
        Private Sub btnConnectTest_Click(sender As Object, e As EventArgs) Handles btnConnectTest.Click

            Dim optAuthent As Boolean = optV7Authent1.Checked             '認証方法選択
            Dim condb As String = ""

            Dim flg_authentV7 As Boolean = optV7Authent1.Checked
            Dim flg_authentV10 As Boolean = optV10Authent1.Checked

            'コントロール設定値を接続情報へ格納
            Me.Set_Control_To_ConModel(fstmodelv7)
            Me.Set_Control_To_ConModel(fstmodelv10, 1)

            blnRtn = True

            'V7接続処理
            If cnnV7.CnnSession(fstmodelv7, sqlcnnV7, flg_authentV7) = False Then
                condb = REL_FROM_NAME & "DB"
                blnRtn = False
            End If

            'V10接続処理
            If cnnV10.CnnSession(fstmodelv10, sqlcnnV10, flg_authentV10) = False Then
                If condb = "" Then
                    condb = REL_TO_NAME & "DB"
                Else
                    condb = condb & "、" & REL_TO_NAME & "DB"
                End If
                blnRtn = False
            End If

            '接続結果
            If blnRtn Then

                '成功時
                MsgResult = MessageBox.Show(MSG_CNN_SUCCESS_B, "成功", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

                '接続確認ボタン非活性化
                Me.btnConnectTest.Enabled = False
                'kakaka6 test chg s --------------------------------------------------
                'Me.grpConnectInfo.Enabled = False
                Me.grpV7ConnectInfo.Enabled = False
                Me.grp10ConnectInfo.Enabled = False
                'kakaka6 test chg e --------------------------------------------------

                '次へボタン非活性化
                btnNext.Enabled = True

                '2016.02.29 テスト -del sta
                'データ取得
                'Call Rel_GetData()
                'cellchg = True
                '2016.02.29 テスト -del end


            Else

                '失敗時
                MsgResult = MessageBox.Show(MSG_CNN_FAILURE_A & vbCrLf & "接続先：" & condb, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                cnnV7.CnnClose(sqlcnnV7)
                cnnV10.CnnClose(sqlcnnV10)

            End If

        End Sub

        ''' <summary>
        ''' イベント処理：タブ選択
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub tab_Selected(sender As Object, e As TabControlEventArgs)

            '紐付設定タブのタブオーダー設定
            For Each item_main As Control In sender.Controls
                If item_main.GetType().Equals(GetType(TabPage)) Then
                    If DirectCast(item_main, TabPage).Name = sender.SelectedTab.Name Then
                        For Each item_sub As Control In item_main.Controls
                            If item_sub.GetType().Equals(GetType(DataGridView)) Then
                                Call Set_TabOrder_Sub(item_sub)
                            End If
                        Next
                    End If
                End If
            Next

        End Sub

        ''' <summary>
        ''' イベント処理：認証方法のオプションボタン 2015.07.06 sol レビュー後修正_レビュー№14
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub opt_CheckedChanged(sender As Object, e As EventArgs) _
                Handles optV7Authent1.CheckedChanged, optV7Authent2.CheckedChanged, optV10Authent1.CheckedChanged, optV10Authent2.CheckedChanged

            'Windows認証の場合はユーザー名、パスワードを編集不可にする
            'V7
            If optV7Authent2.Checked Then
                Me.txtV7User.Enabled = False
                Me.txtV7Pass.Enabled = False
            Else
                Me.txtV7User.Enabled = True
                Me.txtV7Pass.Enabled = True
            End If

            'V10
            If optV10Authent2.Checked Then
                Me.txtV10User.Enabled = False
                Me.txtV10Pass.Enabled = False
            Else
                Me.txtV10User.Enabled = True
                Me.txtV10Pass.Enabled = True
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：'20160525 全体的な動作の修正
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub dataGridView_CellLeave(sender As Object, e As DataGridViewCellEventArgs) _
            Handles dgvBkrui.CellLeave, dgvHyrui.CellLeave, dgvNkinKbn.CellLeave, dgvToritaiyo.CellLeave, dgvNkinkomk.CellLeave, _
            dgvKozo.CellLeave, dgvKozasyu.CellLeave, dgvSetubi.CellLeave, dgvJisyakoza.CellLeave, dgvGazo.CellLeave, dgvFBInfo.CellLeave, dgvKyrui.CellLeave, _
            dgvTosiYoto.CellLeave, dgvKagi.CellLeave

            Dim dgv As DataGridView = DirectCast(sender, DataGridView)
            Dim rowcurrent As Integer = sender.CurrentCell.RowIndex

            '20160701 指摘事項まとめファイルの対応 選択行の着色修正 -chg sta
            'Dim stacolno As Integer = 0
            'If hash_gridtocolno.Contains(dgv) Then
            '    stacolno = hash_gridtocolno(dgv)
            '    For cntcol = stacolno To dgv.ColumnCount - 1
            '        dgv(cntcol, rowcurrent).Style.BackColor = Color.Beige
            '    Next
            'End If

            Dim bordercol As Integer = 0
            If hash_gridtocolno.Contains(dgv) Then
                bordercol = hash_gridtocolno(dgv)
            End If
            If dgv.Name = "dgvNkinkomk" Then
                Call Me.Set_Color_Nkinkomk_CellNotChg(dgv, rowcurrent, bordercol)
            Else
                Call Me.Set_Color_Leave_All(dgv, rowcurrent, bordercol)
            End If
            '20160701 指摘事項まとめファイルの対応 選択行の着色修正 -chg end

            '20160701 指摘事項まとめファイルの対応 画像種別が物件時のセル着色修正 -del sta
            ''20160602 画像紐付け修正 -add sta
            'If dgv.Name = "dgvGazo" Then
            '    For cntkk = 1 To dgv.RowCount
            '        Dim tmp_10_syubetu As String = dgv(3, cntkk - 1).Value
            '        If tmp_10_syubetu = "部屋" Then
            '            dgv(3, cntkk - 1).Style.BackColor = Color.LightGray
            '            dgv(5, cntkk - 1).Style.BackColor = Color.LightGray
            '        End If
            '    Next
            'End If
            ''20160602 画像紐付け修正 -add end
            '20160701 指摘事項まとめファイルの対応 画像種別が物件時のセル着色修正 -del end

            '20160701 指摘事項まとめファイルの対応 随時変動費以外の変動費メーター分類列のグレイアウト対応 -add sta
            If dgv.Name = "dgvNkinkomk" Then
                Call Me.Set_Color_Nkinkomk_Zuiji(dgv)
            End If
            '20160701 指摘事項まとめファイルの対応 随時変動費以外の変動費メーター分類列のグレイアウト対応 -add end

            'FB割付でフォーマット種別が入出金以外の場合、自社情報設定箇所をグレイアウトさせる
            If dgv.Name = "dgvFBInfo" Then
                '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -chg sta
                'For cntkk = 1 To dgv.RowCount
                '    Dim tmp_fmtkbn As String = dgv(0, cntkk - 1).Value
                '    If tmp_fmtkbn <> "入出金" Then
                '        dgv(12, cntkk - 1).Style.BackColor = Color.LightGray
                '        dgv(13, cntkk - 1).Style.BackColor = Color.LightGray
                '    End If
                'Next
                Call Me.Set_Color_FBFmt(dgv)
                '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -chg end
            End If

            '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -add sta
            If dgv.Name = "dgvGazo" Then
                Call Me.Set_Color_Gazo(dgv)
            End If
            '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -add end
            '20160829 エレベーター移行対応 -add sta
            If dgv.Name = "dgvSetubi" Then
                Call Me.Set_Color_Setubi(dgv)
            End If
            '20160829 エレベーター移行対応 -add end
            '20160829 入金区分紐付不可行制御処理を追加 -add sta
            If dgv.Name = "dgvNkinKbn" Then
                Call Me.Set_Color_NkbnReadOnly(dgv)
            End If
            '20160829 入金区分紐付不可行制御処理を追加 -add end
        End Sub

        ''' <summary>
        ''' イベント処理：セル値チェンジ
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub dataGridView_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
            Handles dgvBkrui.CellValueChanged, dgvHyrui.CellValueChanged, dgvNkinKbn.CellValueChanged, dgvToritaiyo.CellValueChanged, dgvNkinkomk.CellValueChanged, _
                    dgvKozo.CellValueChanged, dgvKozasyu.CellValueChanged, dgvSetubi.CellValueChanged, dgvJisyakoza.CellValueChanged, dgvGazo.CellValueChanged, _
                    dgvFBInfo.CellValueChanged, dgvKyrui.CellValueChanged, dgvTosiYoto.CellValueChanged

            '画面起動時は処理を抜ける
            If cellchg = False Then
                Exit Sub
            End If

            Dim dgv As DataGridView = DirectCast(sender, DataGridView)
            Dim currentvalue As String = dgv.CurrentCell.Value
            Dim colcurrent As Integer = dgv.CurrentCell.ColumnIndex
            Dim rowcurrent As Integer = dgv.CurrentCell.RowIndex
            Dim flg As Boolean = True
            Dim relitem_model As New Object
            Dim relitem_rep As New Object
            Dim relitem As String = hash_gridtorelitem(dgv)

            relitem_model = hash_relitemtomodel(relitem)
            relitem_rep = hash_relitemtorepository(relitem)

            'セル値チェンジイベント内でセル値を変更するためイベントハンドラを削除しておく
            RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged

            'セル変更時に対応する値を出力
            flg = relitem_rep.Set_DgvMatchValue(dgv, colcurrent, rowcurrent, currentvalue)

            '重複チェック
            Select Case relitem
                '2016.04.15 設備の重複不可制御をコメントアウト -del sta 
                'Case "設備マスタ"
                '    Call relitem_rep.Chk_Duplicate(rowcurrent, False)
                '2016.04.15 設備の重複不可制御をコメントアウト -del end
                Case "入金項目マスタ", "自社口座マスタ", "画像割付"           '20160523 画像紐付設定処理の追加 -add (画像割付を追加)
                    If Me.chkDuplicate.Checked = False Then
                        Call relitem_rep.Chk_Duplicate(rowcurrent, False)
                    End If
            End Select

            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add sta
            If relitem = "入金項目マスタ" Then
                Call Me.Set_Color_Nkinkomk_CellChg(dgv, rowcurrent)
            End If
            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add end

            '20160725 親項目変更時の子項目制御 -add sta
            Select Case relitem
                Case "入金区分マスタ", "自社口座マスタ", "入金項目マスタ", "契約分類マスタ", "都市計画・用途地域マスタ", "FBフォーマット割付", "設備マスタ"
                    Call relitem_rep.Chg_ChildValue(dgv, colcurrent, rowcurrent)
            End Select
            '20160725 親項目変更時の子項目制御 -add end

            '20160825 紐付画面件数表示処理対応 -add sta
            '件数表示
            Call relitem_rep.Set_RelCntHyoji()
            '20160825 紐付画面件数表示処理対応 -add end

            '20160829 エレベーター移行対応 -add sta
            If dgv.Name = "dgvSetubi" Then
                Call Me.Set_Color_Setubi(dgv)
            End If
            '20160829 エレベーター移行対応 -add end

            '20160525 全体的な動作の修正 -del sta
            '選択行を着色するように修正
            ''2016.03.23 未設定箇所への着色機能の追加 -add sta
            ''未設定箇所へ着色
            'Dim stacolno As Integer = 0
            'If hash_gridtocolno.Contains(dgv) Then
            '    stacolno = hash_gridtocolno(dgv)
            '    For cntcol = stacolno To dgv.ColumnCount - 1
            '        Dim value As String = IIf(dgv(cntcol, rowcurrent).Value Is Nothing, "", dgv(cntcol, rowcurrent).Value)
            '        If value = "" Then
            '            dgv(cntcol, rowcurrent).Style.BackColor = Color.Pink
            '        Else
            '            dgv(cntcol, rowcurrent).Style.BackColor = Color.Beige
            '        End If
            '    Next
            'End If
            ''2016.03.23 未設定箇所への着色機能の追加 -add end
            '20160525 全体的な動作の修正 -del end

            '削除したイベントハンドラを再度追加
            AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged

        End Sub

        ''' <summary>
        ''' イベント処理：セル編集終了処理 2015.07.06 sol レビュー後修正_レビュー№23
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub dgv_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)


            Dim dgvname As String = sender.Name
            Dim currentvalue As String = sender.CurrentCell.Value
            Dim colcurrent As Integer = sender.CurrentCell.ColumnIndex
            Dim rowcurrent As Integer = sender.CurrentCell.RowIndex

            Dim relitem_model As New Object
            Dim relitem_rep As New Object
            Dim relitem As String = hash_gridtorelitem(sender)

            relitem_model = hash_relitemtomodel(relitem)
            relitem_rep = hash_relitemtorepository(relitem)

            'カレントセルが空文字の場合は処理を抜ける
            If currentvalue = "" Then
                Exit Sub
            End If



            ''グリッド毎に処理
            ''コンボボックスが設定されている列に対し、前回の値と比較して異なっている場合は紐付く項目を削除
            'Select Case dgvname
            '    Case "dgvNkinKbn"
            '        Select Case colcurrent
            '            Case 3
            '                If taihi_nkbnname <> "" And taihi_nkbnname <> currentvalue Then
            '                    Call Del_CellValue(sender, colcurrent, rowcurrent)
            '                End If
            '                taihi_nkbnname = currentvalue
            '        End Select
            '    Case "dgvNkinkomk"
            '        Select Case colcurrent
            '            Case 4
            '                If taihi_nkinname <> "" And taihi_nkinname <> currentvalue Then
            '                    Call Del_CellValue(sender, colcurrent, rowcurrent)
            '                End If
            '                taihi_nkinname = currentvalue
            '            Case 6
            '                If taihi_nkinzkseiname <> "" And taihi_nkinzkseiname <> currentvalue Then
            '                    Call Del_CellValue(sender, colcurrent, rowcurrent)
            '                End If
            '                taihi_nkinzkseiname = currentvalue
            '        End Select
            '    Case "dgvSetubi"
            '        Select Case colcurrent
            '            Case 6
            '                If taihi_setubigrp <> "" And taihi_setubigrp <> currentvalue Then
            '                    Call Del_CellValue(sender, colcurrent, rowcurrent)
            '                End If
            '                taihi_setubigrp = currentvalue
            '            Case 10
            '                If taihi_setubimst <> "" And taihi_setubimst <> currentvalue Then
            '                    Call Del_CellValue(sender, colcurrent, rowcurrent)
            '                End If
            '                taihi_setubimst = currentvalue
            '            Case 14
            '                If taihi_setubikomk <> "" And taihi_setubikomk <> currentvalue Then
            '                    Call Del_CellValue(sender, colcurrent, rowcurrent)
            '                End If
            '                taihi_setubikomk = currentvalue
            '        End Select
            'End Select

        End Sub

        ''' <summary>
        ''' イベント処理：セルクリック
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub dataGridView_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
            Handles dgvBkrui.CellClick, dgvHyrui.CellClick, dgvNkinKbn.CellClick, dgvToritaiyo.CellClick, dgvNkinkomk.CellClick, _
                    dgvKozo.CellClick, dgvKozasyu.CellClick, dgvSetubi.CellClick, dgvJisyakoza.CellClick, dgvGazo.CellClick, dgvFBInfo.CellClick, dgvKyrui.CellClick, _
                    dgvTosiYoto.CellClick, dgvKagi.CellClick

            Dim dgv As DataGridView = DirectCast(sender, DataGridView)
            Dim colcurrent As Integer = sender.CurrentCell.ColumnIndex
            Dim rowcurrent As Integer = sender.CurrentCell.RowIndex
            Dim currentvalue As String = sender(colcurrent, rowcurrent).Value
            Dim relitem_model As New Object
            Dim relitem_rep As New Object
            Dim cellcombo As New DataGridViewComboBoxCell

            Dim relitem As String = hash_gridtorelitem(sender)
            relitem_model = hash_relitemtomodel(relitem)
            relitem_rep = hash_relitemtorepository(relitem)

            '20160523 画像紐付設定処理の追加 -add sta
            'セルの読み取り専用制御
            Select Case relitem     '20160701 指摘事項まとめファイルの対応 入金区分を追加 -add
                Case "画像割付", "設備マスタ", "自社口座マスタ", "契約分類マスタ", "都市計画・用途地域マスタ", "FBフォーマット割付", "入金区分マスタ"   '20160525 全体的な動作の修正 "設備マスタ"、"設備マスタ"、"契約分類マスタ" を追加
                    Call relitem_rep.Set_CellReadOnly(colcurrent, rowcurrent)
                Case "入金項目マスタ"  '20160525 入金項目紐付設定の修正
                    Call relitem_rep.Set_CellReadOnly(colcurrent, rowcurrent)
                    Call relitem_rep.Set_CellReadOnly_Hendo(colcurrent, rowcurrent)
            End Select
            '20160523 画像紐付設定処理の追加 -add end

            Select Case relitem
                Case "設備マスタ", "入金項目マスタ", "画像割付", "都市計画・用途地域マスタ", "FBフォーマット割付"     '20160523 画像紐付設定処理の追加(画像割付を追加)
                    Call relitem_rep.Set_DgvCmbList(colcurrent, rowcurrent, currentvalue, cellcombo)
                Case Else
                    Call relitem_rep.Set_DgvCmbList(colcurrent, currentvalue, cellcombo)
            End Select

            If cellcombo.Items.Count <> 0 Then
                RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                sender(colcurrent, rowcurrent) = cellcombo
                cellcombo.DisplayStyleForCurrentCellOnly = True
                sender(colcurrent, rowcurrent).Value = ""
                sender(colcurrent, rowcurrent).Value = currentvalue
                '20161227 コンボボックスプルダウンの表示幅修正 -add
                Call Me.Set_CmbListWidth(cellcombo)
                AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
            End If

            '20160701 指摘事項まとめファイルの対応 選択行の着色修正 -chg sta
            ''20160525 全体的な動作の修正 -add sta
            'Dim stacolno As Integer = 0
            'If hash_gridtocolno.Contains(dgv) Then
            '    stacolno = hash_gridtocolno(dgv)
            '    For cntcol = stacolno To dgv.ColumnCount - 1
            '        dgv(cntcol, rowcurrent).Style.BackColor = Color.Pink
            '    Next
            'End If
            ''20160525 全体的な動作の修正 -add end

            '選択行の着色
            Call Me.Set_Color_Select_All(dgv, rowcurrent)
            '20160701 指摘事項まとめファイルの対応 選択行の着色修正 -chg end

            '20160602 画像紐付け修正 -add sta
            If dgv.Name = "dgvGazo" Then
                '20160701 指摘事項まとめファイルの対応 画像種別が物件時のセル着色修正 -chg sta
                'For cntkk = 1 To dgv.RowCount
                '    Dim tmp_10_syubetu As String = dgv(3, cntkk - 1).Value
                '    If tmp_10_syubetu = "部屋" Then
                '        dgv(3, cntkk - 1).Style.BackColor = Color.LightGray
                '        dgv(5, cntkk - 1).Style.BackColor = Color.LightGray
                '    End If
                'Next
                Call Me.Set_Color_Gazo(dgv)
                '20160701 指摘事項まとめファイルの対応 画像種別が物件時のセル着色修正 -chg end
            End If
            '20160602 画像紐付け修正 -add end

            '20160829 エレベーター移行対応 -add sta
            If dgv.Name = "dgvSetubi" Then
                Call Me.Set_Color_Setubi(dgv)
            End If
            '20160829 エレベーター移行対応 -add end

            '20160701 指摘事項まとめファイルの対応 随時変動費以外の変動費メーター分類列のグレイアウト対応 -add sta
            If dgv.Name = "dgvNkinkomk" Then
                Call Me.Set_Color_Nkinkomk_Zuiji(dgv)
            End If
            '20160701 指摘事項まとめファイルの対応 随時変動費以外の変動費メーター分類列のグレイアウト対応 -add end

            'FB割付でフォーマット種別が入出金以外の場合、自社情報設定箇所をグレイアウトさせる
            If dgv.Name = "dgvFBInfo" Then
                '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -chg sta
                'For cntkk = 1 To dgv.RowCount
                '    Dim tmp_fmtkbn As String = dgv(0, cntkk - 1).Value
                '    If tmp_fmtkbn <> "入出金" Then
                '        dgv(12, cntkk - 1).Style.BackColor = Color.LightGray
                '        dgv(13, cntkk - 1).Style.BackColor = Color.LightGray
                '    End If
                'Next
                Call Me.Set_Color_FBFmt(dgv)
                '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -chg end
            End If
            '20160725 コンボボックス背景色変更 -add
            '20160829 入金区分紐付不可行制御処理を追加 -add sta
            If dgv.Name = "dgvNkinKbn" Then
                Call Me.Set_Color_NkbnReadOnly(dgv)
            End If
            '20160829 入金区分紐付不可行制御処理を追加 -add end
            cellcombo.Style.BackColor = Nothing

            'Select Case dgvname
            '    Case "dgvBkrui"
            '        If colcurrent = 2 Then
            '            'Call Set_DgvCmb(sender, model_m_bk_rui, colcurrent, rowcurrent, currentvalue)
            '            'Call Set_DgvCmb(sender, model_relitem, colcurrent, rowcurrent, currentvalue)

            '        End If
            '    Case "dgvHyrui"
            '        If colcurrent = 3 Then
            '            Call Set_DgvCmb(sender, model_m_hy_rui, colcurrent, rowcurrent, currentvalue)
            '        End If
            '    Case "dgvNkinKbn"
            '        Select Case colcurrent
            '            Case 3
            '                Call Set_DgvCmb(sender, model_nkbn_Rel, colcurrent, rowcurrent, currentvalue)
            '            Case 5
            '                If sender(2, rowcurrent).Value = "" Then
            '                    sender(4, rowcurrent).ReadOnly = True
            '                    sender(5, rowcurrent).ReadOnly = True
            '                Else
            '                    sender(4, rowcurrent).ReadOnly = False
            '                    sender(5, rowcurrent).ReadOnly = False
            '                    Call Set_DgvCmb(sender, model_nkbn_z_Rel, colcurrent, rowcurrent, currentvalue)
            '                End If
            '        End Select
            '    Case "dgvToritaiyo"
            '        If colcurrent = 3 Then
            '            Call Set_DgvCmb(sender, model_toritaiyo_Rel, colcurrent, rowcurrent, currentvalue)
            '        End If
            '    Case "dgvKozo"
            '        If colcurrent = 3 Then
            '            Call Set_DgvCmb(sender, model_kozo_Rel, colcurrent, rowcurrent, currentvalue)
            '        End If
            '    Case "dgvKozasyu"
            '        If colcurrent = 3 Then
            '            Call Set_DgvCmb(sender, model_kozasyu_Rel, colcurrent, rowcurrent, currentvalue)
            '        End If
            '    Case "dgvNkinkomk"
            '        Select Case colcurrent
            '            Case 4
            '                Call Set_DgvCmb(sender, model_m_nkin_Rel, colcurrent, rowcurrent, currentvalue, sender(2, rowcurrent).Value)
            '            Case 6
            '                If sender(3, rowcurrent).Value = "" Then
            '                    sender(5, rowcurrent).ReadOnly = True
            '                    sender(6, rowcurrent).ReadOnly = True
            '                Else
            '                    sender(5, rowcurrent).ReadOnly = False
            '                    sender(6, rowcurrent).ReadOnly = False
            '                    Call Set_DgvCmb(sender, model_m_nkin_z_Rel, colcurrent, rowcurrent, currentvalue)
            '                End If
            '        End Select
            '    Case "dgvSetubi"
            '        Dim tmp_guid As String = ""
            '        Select Case colcurrent
            '            Case 6
            '                Call Set_DgvCmb(sender, model_setubi_grp_Rel, colcurrent, rowcurrent, currentvalue)
            '            Case 10
            '                If sender(4, rowcurrent).Value = "" Then
            '                    sender(9, rowcurrent).ReadOnly = True
            '                    sender(10, rowcurrent).ReadOnly = True
            '                Else
            '                    sender(9, rowcurrent).ReadOnly = False
            '                    sender(10, rowcurrent).ReadOnly = False
            '                    tmp_guid = sender(4, rowcurrent).Value
            '                    Call Set_DgvCmb(sender, model_setubi_ms_Rel, colcurrent, rowcurrent, currentvalue, sender(4, rowcurrent).Value)
            '                End If
            '            Case 14
            '                If sender(7, rowcurrent).Value = "" Then
            '                    sender(13, rowcurrent).ReadOnly = True
            '                    sender(14, rowcurrent).ReadOnly = True
            '                Else
            '                    sender(13, rowcurrent).ReadOnly = False
            '                    sender(14, rowcurrent).ReadOnly = False
            '                    tmp_guid = sender(7, rowcurrent).Value
            '                    Call Set_DgvCmb(sender, model_setubi_komk_Rel, colcurrent, rowcurrent, currentvalue, tmp_guid)
            '                End If
            '        End Select
            'End Select

        End Sub

        ''' <summary>
        ''' イベント処理：セルエディットコントロール取得
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub dgv_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) _
            Handles dgvBkrui.EditingControlShowing, dgvHyrui.EditingControlShowing, dgvNkinKbn.EditingControlShowing, dgvToritaiyo.EditingControlShowing, dgvNkinkomk.EditingControlShowing, _
                    dgvKozo.EditingControlShowing, dgvKozasyu.EditingControlShowing, dgvSetubi.EditingControlShowing, dgvJisyakoza.EditingControlShowing, dgvGazo.EditingControlShowing, _
                    dgvFBInfo.EditingControlShowing, dgvKyrui.EditingControlShowing, dgvTosiYoto.EditingControlShowing

            Dim dgvname As String = sender.Name

            If TypeOf e.Control Is TextBox Then

                Dim colcurrent As Integer = sender.CurrentCell.ColumnIndex
                Dim dgvtexteditctrl As DataGridViewTextBoxEditingControl = DirectCast(e.Control, DataGridViewTextBoxEditingControl)

                AddHandler dgvtexteditctrl.KeyDown, AddressOf dgv_KeyDown

                Select Case dgvname
                    Case "dgvBkrui", "dgvHyrui", "dgvToritaiyo", "dgvKozo", "dgvKozasyu", "dgvKyrui"        '20160525 全体的な動作の修正 -add(dgvBkrui以外を追加)
                        If colcurrent = 1 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                        '20160525 全体的な動作の修正 -del sta
                        'Case "dgvHyrui", "dgvToritaiyo", "dgvKozo", "dgvKozasyu"
                        '    If colcurrent = 2 Then
                        '        'dgvtextedit = DirectCast(e.Control, TextBox)
                        '    End If
                        '20160525 全体的な動作の修正 -del end
                    Case "dgvJisyakoza"     '20160525 全体的な動作の修正 -add
                        If colcurrent = 15 Or colcurrent = 17 Then          '20160526 自社口座本体側の修正反映 12、14 → 14、16列に修正 '20160829 自社口座にデフォルト値を設定する処理を追加 14、16列→15、17列
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                    Case "dgvTosiYoto"     '20160525 全体的な動作の修正 -add
                        If colcurrent = 2 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                    Case "dgvNkinKbn"
                        '20160525 全体的な動作の修正 -chg sta
                        'If colcurrent = 2 Or colcurrent = 4 Then
                        '    'dgvtextedit = DirectCast(e.Control, TextBox)
                        'End If
                        If colcurrent = 1 Or colcurrent = 3 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                        '20160525 全体的な動作の修正 -chg end
                    Case "dgvNkinkomk"
                        '20160525 全体的な動作の修正 -chg sta
                        'If colcurrent = 3 Or colcurrent = 5 Then
                        '    'dgvtextedit = DirectCast(e.Control, TextBox)
                        'End If
                        If colcurrent = 2 Or colcurrent = 4 Or colcurrent = 7 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                        '20160525 全体的な動作の修正 -chg end
                    Case "dgvSetubi"
                        '20160525 全体的な動作の修正 -chg sta
                        'If colcurrent = 5 Or colcurrent = 9 Or colcurrent = 13 Then
                        '    'dgvtextedit = DirectCast(e.Control, TextBox)
                        'End If
                        If colcurrent = 2 Or colcurrent = 4 Or colcurrent = 6 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                        '20160525 全体的な動作の修正 -chg end
                    Case "dgvGazo"  '20160523 画像紐付設定処理の追加 -add
                        If colcurrent = 4 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                    Case "dgvFBInfo"  '20160525 FBフォーマット紐付設定処理の追加 -add
                        If colcurrent = 10 Or colcurrent = 12 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                    Case "dgvKyrui"     '20160526 契約分類マスタの追加 -add
                        If colcurrent = 1 Then
                            AddHandler dgvtexteditctrl.KeyPress, AddressOf dgvtext_KeyPress
                        End If
                End Select

            End If

        End Sub

        ''' <summary>
        ''' イベント処理：セルテキストボックスキープレス
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub dgvtext_KeyPress(sender As Object, e As KeyPressEventArgs)

            'セルへの入力制限(数値のみ)
            If (e.KeyChar < "0"c Or e.KeyChar > "9"c) AndAlso e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：セルキーダウン
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub dgv_KeyDown(sender As Object, e As KeyEventArgs)

            If e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then
                sender = ""
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：重複許可チェックチェンジ
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub chkDuplicate_CheckedChanged(sender As Object, e As EventArgs) Handles chkDuplicate.CheckedChanged

            Dim dgv As DataGridView = Me.dgvNkinkomk
            Dim currentvalue As String = dgv.CurrentCell.Value
            Dim colcurrent As Integer = dgv.CurrentCell.ColumnIndex
            Dim rowcurrent As Integer = dgv.CurrentCell.RowIndex
            Dim flg As Boolean = True
            Dim relitem_model As New Object
            Dim relitem_rep As New Object
            Dim relitem As String = hash_gridtorelitem(dgv)

            relitem_model = hash_relitemtomodel(relitem)
            relitem_rep = hash_relitemtorepository(relitem)

            'セル値を変更するためセル値チェンジイベントハンドラを削除しておく
            RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged

            If Me.chkDuplicate.Checked = False Then
                MsgResult = MessageBox.Show(MSG_CHK_DUPLICATE, "確認", _
                                MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                If MsgResult = DialogResult.Yes Then

                    For cntrow = dgv.RowCount - 1 To 0 Step -1
                        Call relitem_rep.Chk_SetubiDuplicate(cntrow, True)
                    Next

                Else
                    Me.chkDuplicate.Checked = True
                End If

            End If

            '削除したイベントハンドラを再度追加
            AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged

        End Sub

        ''' <summary>
        ''' イベント処理： '20160525 全体的な動作の修正 -add
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub dgvcell_KeyDown(sender As Object, e As KeyEventArgs) _
            Handles dgvBkrui.KeyDown, dgvHyrui.KeyDown, dgvNkinKbn.KeyDown, dgvToritaiyo.KeyDown, dgvNkinkomk.KeyDown, dgvKozo.KeyDown, dgvKozasyu.KeyDown, dgvSetubi.KeyDown, _
                    dgvKagi.KeyDown, dgvJisyakoza.KeyDown, dgvGazo.KeyDown, dgvFBInfo.KeyDown, dgvKyrui.KeyDown, dgvTosiYoto.KeyDown

            Dim dgv As DataGridView = DirectCast(sender, DataGridView)            '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -add
            Dim colcurrent As Integer = sender.CurrentCell.ColumnIndex
            Dim rowcurrent As Integer = sender.CurrentCell.RowIndex
            Dim relitem As String = hash_gridtorelitem(sender)
            Dim relitem_rep As New Object
            relitem_rep = hash_relitemtorepository(relitem)

            If e.KeyCode = Keys.Up Or e.KeyCode = Keys.Left Or e.KeyCode = Keys.Right Or e.KeyCode = Keys.Down Then

                'セルの読み取り専用制御
                Select Case relitem     '20160701 指摘事項まとめファイルの対応 入金区分を追加 -add
                    Case "画像割付", "設備マスタ", "自社口座マスタ", "契約分類マスタ", "都市計画・用途地域マスタ", "FBフォーマット割付", "入金区分マスタ"
                        Call relitem_rep.Set_CellReadOnly(colcurrent, rowcurrent)
                    Case "入金項目マスタ"
                        Call relitem_rep.Set_CellReadOnly(colcurrent, rowcurrent)
                        Call relitem_rep.Set_CellReadOnly_Hendo(colcurrent, rowcurrent)
                End Select

                '20160701 指摘事項まとめファイルの対応 DELETEKEY対応 -add sta
            ElseIf e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then
                'カレントセルが読み取り専用でない場合
                If sender(colcurrent, rowcurrent).ReadOnly = False Then
                    sender(colcurrent, rowcurrent).Value = Nothing
                End If
                '20160701 指摘事項まとめファイルの対応 DELETEKEY対応 -add end
            End If

        End Sub

        ''' <summary>
        ''' イベント処理：入金項目No任意コード一括設定 '20160622 入金項目No任意コード一括設定
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnSetNkinnoBulk_Click(sender As Object, e As EventArgs) Handles btnSetNkinnoBulk.Click

            Dim relitem_rep As New Object
            Dim cnt As Integer = 0

            '開始No設定確認
            If Me.txtNkinnoSta.Text.Trim = "" Then
                MsgResult = MessageBox.Show(MSG_ERR_SET_NKINNO, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            '実行確認
            MsgResult = MessageBox.Show(MSG_STA_SET_NKINNO, "確認", _
                                                        MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                Exit Sub
            End If

            'オブジェクト取得
            relitem_rep = hash_relitemtorepository("入金項目マスタ")

            'メイン処理
            cnt = Int32.Parse(Me.txtNkinnoSta.Text)

            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add sta
            'セル値チェンジイベント内でセル値を変更するためイベントハンドラを削除しておく
            Dim dgv As New DataGridView
            dgv = Me.dgvNkinkomk
            RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add end

            Call relitem_rep.Set_NkinnoBulk(cnt, chgcolorrow)

            '20160825 紐付画面件数表示処理対応 -add
            Call relitem_rep.Set_RelCntHyoji()

            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add sta
            '削除したイベントハンドラを再度追加
            AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add end

            '終了確認
            MsgResult = MessageBox.Show(MSG_END_SET_NKINNO, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        End Sub

        ''' <summary>
        ''' イベント処理：入金項目No任意コード開始No設定制御 '20160622 入金項目No任意コード一括設定
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub txtNkinnoSta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNkinnoSta.KeyPress
            '0～9と、バックスペース以外の時は、イベントをキャンセルする
            If (e.KeyChar < "0"c OrElse "9"c < e.KeyChar) AndAlso _
                    e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If
        End Sub

        ''' <summary>
        ''' イベント処理：紐付選択画面遷移
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' ・画面左の紐付設定項目をクリックすると、関連タブへ遷移
        ''' ・選択されたラベルの背景色を変更
        ''' </remarks>
        Private Sub lblRelSelectKomok_Click(sender As Object, e As EventArgs) Handles _
                                            lblRelSelectBkBunrui.Click, lblRelSelectHyBunrui.Click, lblRelSelectNkKbn.Click, _
                                            lblRelSelectTaiyo.Click, lblRelSelectNkin.Click, lblRelSelectKozo.Click, _
                                            lblRelSelectKozaSyubetu.Click, lblRelSelectHySetubi.Click, lblRelSelectKyoyoKagi.Click, _
                                            lblRelSelectJisyaKoza.Click, lblRelSelectGazo.Click, lblRelSelectFBFmt.Click, _
                                            lblRelSelectKyBunrui.Click, lblRelSelectTosiYoto.Click

            Dim bfnum As Integer = tabCtrlRelwork.SelectedIndex
            Dim afnum As Integer = Nothing
            Dim aftab As String = Nothing

            If sender.Enabled Then
                Select Case sender.Name
                    Case "lblRelSelectBkBunrui" : afnum = 0 : aftab = "tabPage10"
                    Case "lblRelSelectHyBunrui" : afnum = 1 : aftab = "tabPage11"
                    Case "lblRelSelectNkKbn" : afnum = 2 : aftab = "tabPage12"
                    Case "lblRelSelectTaiyo" : afnum = 3 : aftab = "tabPage13"
                    Case "lblRelSelectNkin" : afnum = 4 : aftab = "tabPage14"
                    Case "lblRelSelectKozo" : afnum = 5 : aftab = "tabPage15"
                    Case "lblRelSelectKozaSyubetu" : afnum = 6 : aftab = "tabPage16"
                    Case "lblRelSelectHySetubi" : afnum = 7 : aftab = "tabPage17"
                    Case "lblRelSelectKyoyoKagi" : afnum = 8 : aftab = "tabPage18"
                    Case "lblRelSelectJisyaKoza" : afnum = 9 : aftab = "tabPage19"
                    Case "lblRelSelectGazo" : afnum = 10 : aftab = "tabPage20"
                    Case "lblRelSelectFBFmt" : afnum = 11 : aftab = "tabPage21"
                    Case "lblRelSelectKyBunrui" : afnum = 12 : aftab = "tabPage22"
                    Case "lblRelSelectTosiYoto" : afnum = 13 : aftab = "tabPage23"
                    Case Else
                        Exit Sub
                End Select
                If sender.ForeColor <> escFColor Then
                    Call Chg_IniBackColor()
                    tabCtrlRelwork.SelectedTab = tabCtrlRelwork.TabPages(aftab)
                    sender.Visible = True
                    sender.ForeColor = Color.Black
                    sender.BackColor = Color.Pink
                End If
            End If
        End Sub

#End Region

#Region "読込処理"

        ''' <summary>
        ''' 読込前処理
        ''' </summary>
        ''' <remarks></remarks>
        Private Function Rel_ReadDataBefore() As Boolean

            Dim rtn As Boolean = True

            '------------------------------------
            '各ファイル/フォルダ有無チェック
            '------------------------------------
            'ファイルが存在しない場合は処理を抜ける
            Dim errdirname As String = ""
            If Me.Chk_SettingDir(errdirname) = False Then
                MsgResult = MessageBox.Show("", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                rtn = False
                Return rtn
            End If

            '------------------------------------
            '紐付項目取得
            '------------------------------------
            Dim list_relitem As New List(Of String)
            Call Me.Set_RelItem(list_relitem)
            '選択されていない場合は処理を抜ける
            If list_relitem.Count = 0 Then
                MsgResult = MessageBox.Show(MSG_RELITEMERR, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                rtn = False
                Return rtn
            End If

            '------------------------------------
            '実行確認
            '------------------------------------
            MsgResult = MessageBox.Show(MSG_READ_ITEM, "確認", _
                                                        MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2)
            If MsgResult = DialogResult.No Then
                rtn = False
                Return rtn
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 読込処理
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Rel_ReadData()

            Dim normalflg As Boolean = True     '正常終了フラグ (True:正常終了 False:異常終了)
            Dim tmpcnt As Integer = 0
            Dim msgflg As Boolean = True                            '20160622 前回設定値復元処理修正(前回設定値の出力確認メッセージ表示フラグ)
            Dim existdatareadflg As Boolean = False                 '20160622 前回設定値復元処理修正(前回設定値の出力フラグ)
            Dim list_existdata_readerr As New List(Of String)       '20160622 前回設定値復元処理修正(前回設定値を復元できなかった項目格納用)

            '2016.04.11 呼出起動の修正 -add sta
            Dim obj_pgb As New ProgressBarManager
            Dim pgbform As New Form
            Dim pgbtotalcnt As Integer = 0
            Dim pgbcnt As Integer = 0
            '2016.04.11 呼出起動の修正 -add end

            '2016.04.11 呼出起動の修正 -add sta
            Me.lblRelItemRead.Text = SITUATION_READ
            Me.lblRelItemRead.Visible = True
            Me.pgbRelItemRead.Visible = True
            Me.tabCtrlRelwork.Visible = False
            '2016.04.11 呼出起動の修正 -add end
            '20160725 紐付設定改善 add 
            Dim fstselectflg As Boolean = True                     'TRUE.紐付設定項目初期選択

            '------------------------------------
            '紐付項目取得
            '------------------------------------
            Dim list_relitem As New List(Of String)
            Call Me.Set_RelItem(list_relitem)

            '------------------------------------
            'ログ初期設定
            '------------------------------------
            'ログテーブル作成
            Call Me.Set_LogInit()

            '------------------------------------
            '初期化
            '------------------------------------
            hash_relitemtogrid.Clear()
            hash_gridtorelitem.Clear()
            hash_relitemtomodel.Clear()
            hash_relitemtorepository.Clear()
            hash_gridtocolno.Clear()
            Call Me.IniFormTabPage4(False)
            Call Me.Set_RelItemObjInit()        '20160525 全体的な動作の修正 -add

            '2016.04.11 呼出起動の修正 -add sta
            '------------------------------------
            'プログレスバー設定
            '------------------------------------
            pgbtotalcnt = list_relitem.Count
            Call obj_pgb.pgbInit(pgbtotalcnt)
            '2016.04.11 呼出起動の修正 -add end

            '20160622 前回設定値復元処理修正 -del sta
            '前回設定値復元処理の修正に伴い移動する
            ''20160527 前回設定値復元判別処理の追加 -add sta
            'Dim existdataflg As Boolean = False
            'Dim existdatareadflg As Boolean = False
            'Call ReadRel_ExistData(list_relitem, existdataflg)
            'If existdataflg Then
            '    MsgResult = MessageBox.Show(MSG_EXISTDATA_READ & vbCrLf & MSG_EXISTDATA_READ_DEV, "確認", _
            '                                            MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            '    If MsgResult = DialogResult.Yes Then
            '        existdatareadflg = True
            '    End If
            'End If
            ''20160527 前回設定値復元判別処理の追加 -add end
            '20160622 前回設定値復元処理修正 -del end
            '20160825 前回設定値読込確認メッセージを最初に表示するように修正 -add sta
            Dim existdataflg As Boolean = False
            '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -chg sta
            'Call ReadRel_ExistData(list_relitem, existdataflg)
            If Me.ReadRel_ExistData(list_relitem, existdataflg) = False Then
                Environment.ExitCode = 99
                Me.Close()
                Exit Sub
            End If
            '20161104 紐付設定値保存用ファイルの読込エラー時の対応 -chg end
            If existdataflg Then
                '20160905 汎用紐付対応 メッセージ文言修正 -chg sta
                'MsgResult = MessageBox.Show(MSG_EXISTDATA_READ & vbCrLf & MSG_EXISTDATA_READ_DEV, "確認", _
                '                                        MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                Dim tmp_msg As String = ""
                Select Case CNVNO
                    Case ConvertTypes._既存ユーザ用
                        tmp_msg = MSG_EXISTDATA_READ & vbCrLf & MSG_EXISTDATA_READ_DEV
                    Case ConvertTypes._汎用
                        tmp_msg = MSG_EXISTDATA_READ
                End Select
                MsgResult = MessageBox.Show(tmp_msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                '20160905 汎用紐付対応 メッセージ文言修正 -chg end
                If MsgResult = DialogResult.Yes Then
                    existdatareadflg = True
                End If
            End If
            '20160825 前回設定値読込確認メッセージを最初に表示するように修正 -add end
            '------------------------------------
            '読込開始
            '------------------------------------
            For Each relitem In list_relitem

                '2016.04.11 呼出起動の修正 -add sta
                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Exit For
                End If
                '2016.04.11 呼出起動の修正 -add end

                Dim reldataexistflg As Boolean = False      '既存データ有無フラグ (True:既存データあり False:既存データ無し)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム_メインフォーム -chg sta
                'Dim relexistdata As New Object
                Dim relexistdata As New DataTable
                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム_メインフォーム -chg end
                Dim rowcnt As Integer = 0
                Dim colcnt As Integer = 0

                Dim obj_relrepV7 As New Object
                Dim obj_relrep10 As New Object
                Dim dgv_relitem As New DataGridView
                Dim readDBflg As Boolean = True     'DB読込フラグ (True:DB読込 False:ハードコーディング等、DB以外のデータ読込)

                '項目別読込開始ログ出力
                Dim tmp_sql_readsta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(11, relitem))
                DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_readsta, tmpcnt)

                '紐付項目別初期設定 
                '20160725 紐付設定改善 chg -sta
                'Call Me.Set_RelItemInit(relitem, obj_relrepV7, obj_relrep10, readDBflg)
                Call Me.Set_RelItemInit(relitem, obj_relrepV7, obj_relrep10, readDBflg, fstselectflg)
                '20160725 紐付設定改善 chg -end

                '取得したグリッド格納
                dgv_relitem = hash_relitemtogrid(relitem)

                'グリッド共通初期設定
                Call Me.Initialize_Dgv_Common(dgv_relitem)

                'グリッド個別初期設定
                Call obj_relrep10.Initialize_Dgv(dgv_relitem)

                '20160622 前回設定値復元処理修正 -chg sta
                ''初期表示、コンボリスト用データ読込
                'If readDBflg Then
                '    normalflg = obj_relrep10.ReadDB(Me.sqlcnnV10)
                'Else
                '    normalflg = obj_relrep10.ReadRelItem()
                'End If

                ''紐付ファイル読込
                'normalflg = Me.ReadRel(relitem, relexistdata, rowcnt, colcnt, reldataexistflg)

                'If reldataexistflg And existdatareadflg Then    '20160527 前回設定値復元判別処理の追加(条件にexistdatareadflgを追加)

                '    '紐付データが存在する場合はそのまま出力
                '    Call Set_Dgv_Exist(dgv_relitem, relexistdata, rowcnt, colcnt)

                'Else

                '    '紐付データが存在しない場合はV7中間ファイルから読み込み
                '    Select Case relitem
                '        Case "鍵タイトルマスタ", "自社口座マスタ"   '20160525 全体的な動作の修正 -chg   "入金項目マスタ"を削除
                '            normalflg = obj_relrepV7.ReadMid(Me.sqlcnnV10)
                '        Case "入金項目マスタ", "設備マスタ", "画像割付"    '20160523 画像紐付設定処理の追加(画像タイトルはV7DBから直接読み込む) -add sta    '20160525 全体的な動作の修正 -add "入金項目マスタ"、"設備マスタ"を追加
                '            normalflg = obj_relrepV7.ReadMid(Me.sqlcnnV7)
                '        Case Else
                '            normalflg = obj_relrepV7.ReadMid()
                '    End Select

                '    '出力
                '    Call obj_relrep10.Set_Dgv(dgv_relitem)

                'End If

                '*******************************************
                'データの読込
                '*******************************************
                '①革命10側の初期表示用、設定の際のコンボリスト用に10データを読み込む
                If readDBflg Then
                    normalflg = obj_relrep10.ReadDB(Me.sqlcnnV10)
                Else
                    normalflg = obj_relrep10.ReadRelItem()
                End If

                '②中間ファイル読込
                Select Case relitem
                    Case "鍵タイトルマスタ", "自社口座マスタ", "FBフォーマット割付"       '20160701 指摘事項まとめファイルの対応 金融機関名称取得修正 「FBフォーマット割付」を追加 -add
                        normalflg = obj_relrepV7.ReadMid(Me.sqlcnnV10)
                    Case "設備マスタ", "画像割付"                                         '20160829 V7入金区分を革命マスタから取得するように変更 "入金区分マスタ" を追加  '20160915 汎用の紐付時は入金区分および入金項目を中間ファイルから読み取る処理に修正 "入金項目マスタ", "入金区分マスタ" を既存/汎用の条件分岐 -chg 
                        normalflg = obj_relrepV7.ReadMid(Me.sqlcnnV7)
                    Case "入金項目マスタ"                                                  '20160915 汎用の紐付時は入金区分および入金項目を中間ファイルから読み取る処理に修正 -add
                        Select Case CNVNO
                            Case ConvertTypes._既存ユーザ用
                                normalflg = obj_relrepV7.ReadMid(Me.sqlcnnV7)
                            Case ConvertTypes._汎用
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                'normalflg = obj_relrepV7.ReadMid_Base(Me.sqlcnnV7)
                                normalflg = obj_relrepV7.ReadMid_Base(Me.sqlcnnV10)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                        End Select
                    Case "入金区分マスタ"                                                  '20160915 汎用の紐付時は入金区分および入金項目を中間ファイルから読み取る処理に修正 -add
                        Select Case CNVNO
                            Case ConvertTypes._既存ユーザ用
                                normalflg = obj_relrepV7.ReadMid(Me.sqlcnnV7)
                            Case ConvertTypes._汎用
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                'normalflg = obj_relrepV7.ReadMid()
                                normalflg = obj_relrepV7.ReadMid_Base(Me.sqlcnnV10)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                        End Select
                    Case Else
                        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                        'normalflg = obj_relrepV7.ReadMid()
                        normalflg = obj_relrepV7.ReadMid(Me.sqlcnnV10)
                        '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                End Select

                '③紐付ファイル読込
                normalflg = Me.ReadRel(relitem, relexistdata, rowcnt, colcnt, reldataexistflg)

                '20160825 前回設定値読込確認メッセージを最初に表示するように修正 -del sta
                ''④前回設定値がある場合、復元確認を行う
                'If reldataexistflg Then
                '    If msgflg Then

                '        '20160725 汎用CVK対応 chg -sta
                '        'MsgResult = MessageBox.Show(MSG_EXISTDATA_READ & vbCrLf & MSG_EXISTDATA_READ_DEV, "確認", _
                '        '                                   MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                '        Dim tmpmsg As String = ""
                '        If CNVNO = 0 Then
                '            tmpmsg = MSG_H_EXISTDATA_READ_DEV
                '        Else
                '            tmpmsg = MSG_EXISTDATA_READ_DEV
                '        End If
                '        MsgResult = MessageBox.Show(MSG_EXISTDATA_READ & vbCrLf & tmpmsg, "確認", _
                '                                            MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                '        '20160725 汎用CVK対応 chg -end

                '        If MsgResult = DialogResult.Yes Then
                '            existdatareadflg = True
                '        End If
                '        msgflg = False
                '    End If
                'End If
                '20160825 前回設定値読込確認メッセージを最初に表示するように修正 -del end

                '⑤データの出力
                If reldataexistflg = False OrElse existdatareadflg = False Then     '前回設定値が存在しない、または前回設定値を表示しない場合は中間ファイルから読み込んだデータを表示する
                    Call obj_relrep10.Set_Dgv(dgv_relitem)
                Else                                                                '前回設定値が存在し、かつ前回設定値を表示する場合は中間ファイルと前回設定値を比較する
                    Dim matchtype As Integer = Chk_MidAndRelMatching(relitem, relexistdata, rowcnt)
                    Select Case matchtype
                        Case 0    '完全一致
                            '前回設定値出力
                            Call Set_Dgv_Exist(dgv_relitem, relexistdata, rowcnt, colcnt)
                        Case 1    '不一致
                            ''メッセージ
                            'MsgResult = MessageBox.Show(relitem & MSG_ERR_SET_EXISTRELDATA, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            'メッセージは後でまとめて出力する
                            list_existdata_readerr.Add(relitem)
                            '中間ファイル出力
                            Call obj_relrep10.Set_Dgv(dgv_relitem)
                        Case 2    '一致しているが中間ファイルに追加されている分が存在する
                            ''メッセージ
                            'MsgResult = MessageBox.Show(relitem & MSG_ADD_SET_EXISTRELDATA, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            'メッセージは後でまとめて出力する
                            list_existdata_readerr.Add(relitem)
                            '差分表示は構築中のため中間ファイルを表示
                            Call obj_relrep10.Set_Dgv(dgv_relitem)
                    End Select
                End If
                '20160622 前回設定値復元処理修正 -chg end

                '紐付項目をキーにしてハッシュテーブルへ格納 (コンボボックスリストや番号一致時の文字列表示に使用する)
                hash_relitemtorepository.Add(relitem, obj_relrep10)
                '20160729 紐付データが存在しない場合のクリックイベントエラー対応 -add sta
                If dgv_relitem.RowCount = 0 Then
                    dgv_relitem.Enabled = False
                End If
                '20160729 紐付データが存在しない場合のクリックイベントエラー対応 -add end
                '項目別読込終了ログ出力
                Dim tmp_str As String = Chg_FlgToStr(normalflg)
                Dim tmp_sql_readend As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(19, relitem, "-", tmp_str))
                DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_readend, tmpcnt)

                '20160825 紐付画面件数表示処理対応 -add sta
                '件数表示
                Call obj_relrep10.Set_RelCntHyoji()
                '20160825 紐付画面件数表示処理対応 -add end

                '2016.04.11 呼出起動の修正 -add sta
                'プログレスバー設定
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsetting(pgbcnt)
                Me.tabPage3.Refresh()
                '2016.04.11 呼出起動の修正 -add end

                'セルの着色設定
                '20161209 前回設定値復元時にグレイアウトされていないセルが生じるエラーの修正 -chg sta
                'Call Me.Set_Color_NkbnReadOnly(dgv_relitem)  '20160829 入金区分紐付不可行制御処理を追加 -add
                Select Case relitem
                    Case "入金区分マスタ"
                        Call Me.Set_Color_NkbnReadOnly(dgv_relitem)
                    Case "入金項目マスタ"
                        Call Me.Set_Color_Nkinkomk_Zuiji(dgv_relitem)
                    Case "設備マスタ"
                        Call Me.Set_Color_Setubi(dgv_relitem)
                    Case "FBフォーマット割付"
                        Call Me.Set_Color_FBFmt(dgv_relitem)
                End Select
                '20161209 前回設定値復元時にグレイアウトされていないセルが生じるエラーの修正 -chg end

            Next

            '------------------------------------
            '前回設定値の復元不可項目を表示
            '------------------------------------
            If list_existdata_readerr.Count <> 0 Then
                Dim tmp_errmsg As String = ""
                For Each errkomk In list_existdata_readerr
                    tmp_errmsg = tmp_errmsg & vbCrLf & "    " & "・" & errkomk
                Next
                MsgResult = MessageBox.Show(MSG_ERR_EXISTRELDATAREAD & tmp_errmsg, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

            '------------------------------------
            '終了処理
            '------------------------------------

            cellchg = True
            gridchgflg = True

            '2016.04.11 呼出起動の修正 -add sta
            If CancelFlg Then
                If singlesta = False Then
                    Me.Close()
                End If
            Else
                Me.tabPage3.Refresh()
                Me.lblRelItemRead.Visible = False
                Me.pgbRelItemRead.Visible = False
                Me.tabCtrlRelwork.Visible = True
            End If
            '2016.04.11 呼出起動の修正 -add end

            '20160701 指摘事項まとめファイルの対応 ラベル非表示対応 -add
            Me.lblCaution.Visible = True
            Me.Label40.Visible = True  '20160825 ラベル表示制御対応 -add
            Me.Label37.Visible = True  '20160825 ラベル表示制御対応 -add

        End Sub

        ''' <summary>
        ''' グリッドの共通初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Initialize_Dgv_Common(ByVal dgv As DataGridView)

            With dgv

                '初期設定
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                .AllowUserToResizeRows = False
                .AllowUserToResizeColumns = True       '20160701 指摘事項まとめファイルの対応 幅の可変 -chg False → True へ変更
                .AllowUserToAddRows = False
                .RowHeadersVisible = False
                .ColumnHeadersHeight = 20
                .RowTemplate.Height = 20
                .EnableHeadersVisualStyles = False

                '20160610 グリッドのちらつき防止処理追加 -add sta
                Dim myType As System.Type = GetType(DataGridView)
                Dim myPropertyInfo As System.Reflection.PropertyInfo = myType.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)
                myPropertyInfo.SetValue(dgv, True, Nothing)
                '20160610 グリッドのちらつき防止処理追加 -add end

                'kakaka6 chg s -------------------------------------------------- 0420
                'サイズ設定
                '.Width = 750
                '.Height = 350
                '20160725 スクロールバー表示修正 -del sta
                '.Width = 865
                '.Height = 400
                '20160725 スクロールバー表示修正 -del end
                '中央へ配置()
                '.Left = (Me.tabPage10.Width - .Width) / 2
                '.Top = (Me.tabPage10.Height - .Height) / 2
                'kakaka6 del e --------------------------------------------------

                'ソート不可設定
                For Each dgvcol As DataGridViewColumn In dgv.Columns
                    dgvcol.SortMode = DataGridViewColumnSortMode.NotSortable
                Next

            End With

        End Sub

        ''' <summary>
        ''' 既存データのグリッド表示処理
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="existdata"></param>
        ''' <param name="rowcnt"></param>
        ''' <param name="colcnt"></param>
        ''' <remarks></remarks>
        Private Sub Set_Dgv_Exist(ByVal dgv As DataGridView, ByVal existdata As DataTable, ByVal rowcnt As Integer, ByVal colcnt As Integer)

            'グリッドの行設定
            dgv.RowCount = rowcnt

            '表示処理
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            'For cntii = 1 To rowcnt
            '    For cntjj = 1 To colcnt
            '        dgv(cntjj - 1, cntii - 1).Value = existdata(cntii, cntjj)
            '    Next
            'Next
            For cntii = 0 To rowcnt - 1
                For cntjj = 0 To colcnt - 1

                    Dim value As String = Typ.ToStr(existdata.Rows(cntii).Item(cntjj)).Trim

                    '--------------------------------------------------------------------------
                    'チェックボックス列の場合、チェックOFFのセルには何も設定しないようにする
                    'チェックボックス列が存在する項目と列番号
                    '→契約分類マスタ 4列目(インデックス3)
                    '--------------------------------------------------------------------------
                    If dgv Is Me.dgvKyrui Then
                        If cntjj = 3 Then
                            If value <> "True" Then
                                value = Nothing
                            End If
                        End If
                    End If

                    '値をセット
                    dgv(cntjj, cntii).Value = value

                Next
            Next
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
        End Sub

        ''' <summary>
        ''' 紐付ファイルから紐付データを読み込む
        ''' </summary>
        ''' <param name="sheetname"></param>
        ''' <param name="reldata"></param>
        ''' <param name="rowcnt"></param>
        ''' <param name="columncnt"></param>
        ''' <param name="existflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ReadRel(ByVal sheetname As String, ByRef reldata As DataTable, ByRef rowcnt As Integer, ByRef columncnt As Integer, ByRef existflg As Boolean) As Boolean
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            'Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            'Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            'Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            'Dim startrow As Integer                                         '書込開始行
            'Dim maxrowcnt As Integer                                        '既存データの行数
            'Dim rtn As Boolean = True
            'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            'Dim relfldno As Integer = 0

            ''************************
            ''作業準備
            ''************************
            ''20160608 画像CVテストに伴う修正 -add sta
            'If sheetname = "画像割付" Then
            '    Return rtn
            'End If
            ''20160608 画像CVテストに伴う修正 -add end

            ''Excelファイル初期設定
            'rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelDirPath, REL_FILENAME, sheetname)

            ''Excelファイル設定時にエラーが生じた際は処理を抜ける
            'If rtn = False Then
            '    Return rtn
            'End If

            ''************************
            ''処理開始
            ''************************
            ''既存データ件数確認
            'If rowcnt = 0 Then
            '    existflg = False
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            '    Return rtn
            'Else
            '    existflg = True
            'End If

            ''データの部分を取得
            'Dim existdata As Object = wsheet.Range(wsheet.Cells(startrow, 1), wsheet.Cells(rowcnt + 1, columncnt)).Value

            ''データの引渡し
            'reldata = existdata

            ''************************
            ''終了処理
            ''************************
            ''Excelファイル終了設定
            'Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'Return rtn

            '20261005 ACE.OLEDB(Access Database Engine)を使わない読み書きに変更(64ビット対応) -chg sta
            Dim rtn As Boolean = True

            If sheetname = "画像割付" Then
                Return rtn
            End If

            '----- 紐付ファイルからシートの内容を読込 -----
            Dim relfilepath As String = EtcMethod.Set_Path(RelDirPath, REL_FILENAME & ".xlsx")
            Dim dt As DataTable = RelationExcelFile.ReadSheet(relfilepath, sheetname)

            '既存データ件数確認
            rowcnt = dt.Rows.Count
            columncnt = dt.Columns.Count
            If rowcnt = 0 Then
                existflg = False
                Return rtn
            Else
                existflg = True
            End If

            'データの部分を取得
            reldata = dt

            Return rtn
            '20261005 ACE.OLEDB(Access Database Engine)を使わない読み書きに変更(64ビット対応) -chg end
        End Function

        ''' <summary>
        ''' 中間ファイルから紐付データを読み込む
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="sheetname"></param>
        ''' <param name="relfldname"></param>
        ''' <param name="list_relitem"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ReadMid(ByVal filename As String, ByVal sheetname As String, ByVal relfldname As String, ByRef list_relitem As List(Of String)) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim rtn As Boolean = True
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim relfldno As Integer = 0

            '************************
            '作業準備
            '************************

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MidDirPath, filename, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Return rtn
            End If

            '************************
            '処理開始
            '************************

            Dim headerrow As New Object
            Dim relcolvalue As New Object

            'ヘッダー行取得
            headerrow = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '対象フィールドを検索し、ヒットした際にその列の値を取得
            For cntcol = 1 To columncnt

                Dim fldname As String = headerrow(startrow - 1, cntcol).ToString.Trim

                If fldname = relfldname Then
                    relfldno = cntcol
                    relcolvalue = wsheet.Range(wsheet.Cells(startrow, relfldno), wsheet.Cells(rowcnt + 1, relfldno)).Value
                    Exit For
                End If

            Next

            '取得した列のデータをリストへ格納
            For cntii = 1 To rowcnt

                Dim fldvalue As String = ""

                If relcolvalue(cntii, 1) IsNot Nothing Then
                    fldvalue = relcolvalue(cntii, 1).ToString.Trim
                End If

                If list_relitem.Contains(fldvalue) = False Then
                    list_relitem.Add(fldvalue)
                End If

            Next

            '************************
            '終了処理
            '************************

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

        ''' <summary>
        ''' 紐付項目別初期設定 (追加箇所)
        ''' </summary>
        ''' <param name="relitem"></param>
        ''' <param name="objrepV7"></param>
        ''' <param name="objrep10"></param>
        ''' <param name="readDBflg"></param>
        ''' <param name="fstselectflg">初回選択FLG…紐付設定項目初期着色の為に用意</param>
        ''' <remarks>
        ''' 20160725 紐付設定改善 … 引数に[fstselectflg]を追加
        ''' </remarks>
        Private Sub Set_RelItemInit(ByVal relitem As String, ByRef objrepV7 As Object, ByRef objrep10 As Object, ByRef readDBflg As Boolean, ByRef fstselectflg As Boolean)

            Select Case relitem
                Case "物件分類マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(0, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvBkrui)
                    hash_gridtorelitem.Add(Me.dgvBkrui, relitem)
                    hash_gridtocolno.Add(Me.dgvBkrui, 1)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_brui_Repository
                    objrep10 = New Njc.Repository.M_bk_rui_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectBkBunrui.Visible = True
                    Me.lblRelSelectBkBunrui.ForeColor = Color.Black
                    Me.lblRelSelectBkBunrui.BackColor = Color.White
                    Me.lblRelSelectBkBunrui.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectBkBunrui.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "部屋分類マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(1, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvHyrui)
                    hash_gridtorelitem.Add(Me.dgvHyrui, relitem)
                    hash_gridtocolno.Add(Me.dgvHyrui, 1)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_crui_Repository
                    objrep10 = New Njc.Repository.M_hy_rui_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectHyBunrui.Visible = True
                    Me.lblRelSelectHyBunrui.ForeColor = Color.Black
                    Me.lblRelSelectHyBunrui.BackColor = Color.White
                    Me.lblRelSelectHyBunrui.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectHyBunrui.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "入金区分マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(2, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvNkinKbn)
                    hash_gridtorelitem.Add(Me.dgvNkinKbn, relitem)
                    hash_gridtocolno.Add(Me.dgvNkinKbn, 1)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_nkbn_Rev7_Repository
                    objrep10 = New Njc.Repository.M_nkbn_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectNkKbn.Visible = True
                    Me.lblRelSelectNkKbn.ForeColor = Color.Black
                    Me.lblRelSelectNkKbn.BackColor = Color.White
                    Me.lblRelSelectNkKbn.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectNkKbn.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "取引態様マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(3, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvToritaiyo)
                    hash_gridtorelitem.Add(Me.dgvToritaiyo, relitem)
                    hash_gridtocolno.Add(Me.dgvToritaiyo, 1)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_toritaiyo_Rev7_Repository
                    objrep10 = New Njc.Repository.M_toritaiyo_Repository
                    readDBflg = False
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectTaiyo.Visible = True
                    Me.lblRelSelectTaiyo.ForeColor = Color.Black
                    Me.lblRelSelectTaiyo.BackColor = Color.White
                    Me.lblRelSelectTaiyo.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectTaiyo.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "入金項目マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(4, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvNkinkomk)
                    hash_gridtorelitem.Add(Me.dgvNkinkomk, relitem)
                    hash_gridtocolno.Add(Me.dgvNkinkomk, 2)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_nkin_Rev7_Repository
                    objrep10 = New Njc.Repository.M_nkin_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectNkin.Visible = True
                    Me.lblRelSelectNkin.ForeColor = Color.Black
                    Me.lblRelSelectNkin.BackColor = Color.White
                    Me.lblRelSelectNkin.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectNkin.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "構造マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(5, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvKozo)
                    hash_gridtorelitem.Add(Me.dgvKozo, relitem)
                    hash_gridtocolno.Add(Me.dgvKozo, 1)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_kozo_Rev7_Repository
                    objrep10 = New Njc.Repository.M_kozo_Repository
                    readDBflg = False
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectKozo.Visible = True
                    Me.lblRelSelectKozo.ForeColor = Color.Black
                    Me.lblRelSelectKozo.BackColor = Color.White
                    Me.lblRelSelectKozo.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectKozo.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "口座種別マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(6, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvKozasyu)
                    hash_gridtorelitem.Add(Me.dgvKozasyu, relitem)
                    hash_gridtocolno.Add(Me.dgvKozasyu, 1)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_kozasyu_Rev7_Repository
                    objrep10 = New Njc.Repository.M_kozasyu_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectKozaSyubetu.Visible = True
                    Me.lblRelSelectKozaSyubetu.ForeColor = Color.Black
                    Me.lblRelSelectKozaSyubetu.BackColor = Color.White
                    Me.lblRelSelectKozaSyubetu.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectKozaSyubetu.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "設備マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(7, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvSetubi)
                    hash_gridtorelitem.Add(Me.dgvSetubi, relitem)
                    hash_gridtocolno.Add(Me.dgvSetubi, 2)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_setubi_Rev7_Repository
                    objrep10 = New Njc.Repository.M_setubi_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectHySetubi.Visible = True
                    Me.lblRelSelectHySetubi.ForeColor = Color.Black
                    Me.lblRelSelectHySetubi.BackColor = Color.White
                    Me.lblRelSelectHySetubi.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectHySetubi.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "鍵タイトルマスタ"
                    '20160830 鍵紐付箇所の削除 -del
                    'tabPageManager_Sub.ChangeTabPageVisible(8, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvKagi)
                    hash_gridtorelitem.Add(Me.dgvKagi, relitem)
                    hash_gridtocolno.Add(Me.dgvKagi, 2)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_kagi_Rev7_Repository
                    objrep10 = New Njc.Repository.M_kagi_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectKyoyoKagi.Visible = True
                    Me.lblRelSelectKyoyoKagi.ForeColor = Color.Black
                    Me.lblRelSelectKyoyoKagi.BackColor = Color.White
                    Me.lblRelSelectKyoyoKagi.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectKyoyoKagi.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "自社口座マスタ"
                    tabPageManager_Sub.ChangeTabPageVisible(9, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvJisyakoza)
                    hash_gridtorelitem.Add(Me.dgvJisyakoza, relitem)
                    hash_gridtocolno.Add(Me.dgvJisyakoza, 15)            '2016.03.23 未設定箇所への着色機能の追加 -add     '20160526 自社口座本体側の修正反映  7列→14列 '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
                    objrepV7 = New Njc.Repository.M_jisyakoza_Rev7_Repository
                    objrep10 = New Njc.Repository.M_jisyakoza_Repository
                    readDBflg = False
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectJisyaKoza.Visible = True
                    Me.lblRelSelectJisyaKoza.ForeColor = Color.Black
                    Me.lblRelSelectJisyaKoza.BackColor = Color.White
                    Me.lblRelSelectJisyaKoza.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectJisyaKoza.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "画像割付"    '20160523 画像紐付設定処理の追加 -add
                    tabPageManager_Sub.ChangeTabPageVisible(10, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvGazo)
                    hash_gridtorelitem.Add(Me.dgvGazo, relitem)
                    hash_gridtocolno.Add(Me.dgvGazo, 3)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_gazo_title_Rev7_Repository
                    objrep10 = New Njc.Repository.M_gazo_title_Repository
                    readDBflg = False
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectGazo.Visible = True
                    Me.lblRelSelectGazo.ForeColor = Color.Black
                    Me.lblRelSelectGazo.BackColor = Color.White
                    Me.lblRelSelectGazo.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectGazo.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "FBフォーマット割付"    '20160525 FBフォーマット紐付設定処理の追加 -add
                    tabPageManager_Sub.ChangeTabPageVisible(11, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvFBInfo)
                    hash_gridtorelitem.Add(Me.dgvFBInfo, relitem)
                    hash_gridtocolno.Add(Me.dgvFBInfo, 10)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_FBInfo_Rev7_Repository
                    objrep10 = New Njc.Repository.M_FBInfo_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectFBFmt.Visible = True
                    Me.lblRelSelectFBFmt.ForeColor = Color.Black
                    Me.lblRelSelectFBFmt.BackColor = Color.White
                    Me.lblRelSelectFBFmt.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectFBFmt.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "契約分類マスタ"          '20160526 契約分類マスタの追加 -add
                    tabPageManager_Sub.ChangeTabPageVisible(12, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvKyrui)
                    hash_gridtorelitem.Add(Me.dgvKyrui, relitem)
                    hash_gridtocolno.Add(Me.dgvKyrui, 1)            '2016.03.23 未設定箇所への着色機能の追加 -add
                    objrepV7 = New Njc.Repository.M_keirui_Repository
                    objrep10 = New Njc.Repository.M_ky_rui_Repository
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectKyBunrui.Visible = True
                    Me.lblRelSelectKyBunrui.ForeColor = Color.Black
                    Me.lblRelSelectKyBunrui.BackColor = Color.White
                    Me.lblRelSelectKyBunrui.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectKyBunrui.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
                Case "都市計画・用途地域マスタ"     '20160531 都市計画用途地域の追加
                    tabPageManager_Sub.ChangeTabPageVisible(13, True)
                    hash_relitemtogrid.Add(relitem, Me.dgvTosiYoto)
                    hash_gridtorelitem.Add(Me.dgvTosiYoto, relitem)
                    hash_gridtocolno.Add(Me.dgvTosiYoto, 1)
                    objrepV7 = New Njc.Repository.M_TosiYoto_Rev7_Repository
                    objrep10 = New Njc.Repository.M_TosiYoto_Repository
                    readDBflg = False
                    '20160725 紐付設定改善 add -sta
                    Me.lblRelSelectTosiYoto.Visible = True
                    Me.lblRelSelectTosiYoto.ForeColor = Color.Black
                    Me.lblRelSelectTosiYoto.BackColor = Color.White
                    Me.lblRelSelectTosiYoto.Font = New Font("メイリオ", 9, FontStyle.Bold)
                    If fstselectflg Then
                        Me.lblRelSelectTosiYoto.BackColor = Color.Pink
                        fstselectflg = False
                    End If
                    '20160725 紐付設定改善 add -end
            End Select

        End Sub

        ''' <summary>
        ''' 前回設定値が存在するか判別する '20160527 前回設定値復元判別処理の追加
        ''' </summary>
        ''' <param name="list_relitem"></param>
        ''' <param name="existflg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ReadRel_ExistData(ByVal list_relitem As List(Of String), ByRef existflg As Boolean) As Boolean
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            'Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            'Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            'Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            'Dim startrow As Integer                                         '書込開始行
            'Dim maxrowcnt As Integer                                        '既存データの行数
            'Dim rtn As Boolean = True
            'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            'Dim columncnt As Integer = 0
            'Dim rowcnt As Integer = 0

            ''************************
            ''作業準備
            ''************************

            'For Each sheetname In list_relitem

            '    '20160608 画像CVテストに伴う修正 -add sta
            '    If sheetname = "画像割付" Then
            '        Return rtn
            '    End If
            '    '20160608 画像CVテストに伴う修正 -add end

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelDirPath, REL_FILENAME, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If rtn = False Then
            '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            '        Return rtn
            '    End If

            '    '既存データ有無フラグ初期化
            '    existflg = False

            '    '既存データ件数確認
            '    If rowcnt > 0 Then
            '        existflg = True
            '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            '        Return rtn
            '    End If

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'Next

            'Return rtn

            '20261005 ACE.OLEDB(Access Database Engine)を使わない読み書きに変更(64ビット対応) -chg sta
            Dim rtn As Boolean = True
            Dim relfilepath As String = EtcMethod.Set_Path(RelDirPath, REL_FILENAME & ".xlsx")

            If File.Exists(relfilepath) = False Then
                MsgResult = MessageBox.Show("紐付設定ファイルが存在しないためコンバートを続行できません。" & vbCrLf & _
                                            "再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                rtn = False
                Return rtn
            End If

            '----- 紐付ファイルのオープン処理 -----
            Dim wbook As ClosedXML.Excel.XLWorkbook = Nothing
            Try
                wbook = New ClosedXML.Excel.XLWorkbook(relfilepath)

                'アクセス拒否
            Catch ex As UnauthorizedAccessException
                MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                                            "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                rtn = False
                Return rtn
                'その他の例外
            Catch ex As Exception
                MsgResult = MessageBox.Show("紐付設定ファイルの読み込み時にエラーが発生しました。" & vbCrLf & _
                                            "再度セットアップを行ってから、本プログラムを実行して下さい。", _
                                            "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                rtn = False
                Return rtn
            End Try

            '----- いずれかのシートにデータがあるか確認 -----
            Try
                For Each sheetname In list_relitem
                    If sheetname <> "画像割付" Then
                        If RelationExcelFile.HasAnyData(wbook, sheetname) Then
                            existflg = True
                            Return rtn
                        End If
                    End If
                Next
            Finally
                'クローズ処理
                wbook.Dispose()
            End Try

            Return rtn
            '20261005 ACE.OLEDB(Access Database Engine)を使わない読み書きに変更(64ビット対応) -chg end
        End Function

        ''' <summary>
        ''' 前回設定値と中間ファイルの照合
        ''' </summary>
        ''' <param name="relitem"></param>
        ''' <param name="relexistdata"></param>
        ''' <param name="relexistdatarowcnt"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_MidAndRelMatching(ByVal relitem As String, ByVal relexistdata As DataTable, ByVal relexistdatarowcnt As Integer) As Integer '20160622 前回設定値復元処理修正 -add

            Dim rtn_int As Integer = 0      '0…完全一致 1…不一致 2…一致しているが中間ファイルに新たに追加されている
            Dim list_mid As New List(Of String)
            Dim list_rel As New List(Of String)

            '*********************
            '比較する値の取得
            '*********************
            Select Case relitem

                Case "物件分類マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_brui Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_brui.ItemCnt - 1
                        list_mid.Add(RelItem_M_brui.Brui_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "部屋分類マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_crui Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_crui.ItemCnt - 1
                        list_mid.Add(RelItem_M_crui.Crui_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "入金区分マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_nkbn_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_nkbn_Rev7.ItemCnt - 1
                        list_mid.Add(RelItem_M_nkbn_Rev7.Nkbn_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "取引態様マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_toritaiyo_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_toritaiyo_Rev7.ItemCnt - 1
                        list_mid.Add(RelItem_M_toritaiyo_Rev7.Taiyo_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "入金項目マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_nkin_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_nkin_Rev7.ItemCnt - 1
                        Dim syogo1_mid As String = RelItem_M_nkin_Rev7.Nkin_name(cntii)
                        Dim syogo2 As String = RelItem_M_nkin_Rev7.Nkin_kbn(cntii)
                        list_mid.Add(syogo1_mid & "-" & syogo2)
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    Dim syogo1_rel As String = relexistdata(cntjj, 1)
                    '    Dim syogo2_rel As String = relexistdata(cntjj, 2)
                    '    list_rel.Add(syogo1_rel & "-" & syogo2_rel)
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        Dim syogo1_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim
                        Dim syogo2_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(1)).Trim
                        list_rel.Add(syogo1_rel & "-" & syogo2_rel)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "構造マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_kozo_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_kozo_Rev7.ItemCnt - 1
                        list_mid.Add(RelItem_M_kozo_Rev7.Kozo_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "口座種別マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_kozasyu_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_kozasyu_Rev7.ItemCnt - 1
                        list_mid.Add(RelItem_M_kozasyu_Rev7.Kosyu_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "設備マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_setubi_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_setubi_Rev7.ItemCnt - 1
                        Dim syogo1_mid As String = RelItem_M_setubi_Rev7.Setubi_name(cntii)
                        Dim syogo2_mid As String = RelItem_M_setubi_Rev7.Setubi_lstname(cntii)
                        list_mid.Add(syogo1_mid & "-" & syogo2_mid)
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    Dim syogo1_rel As String = relexistdata(cntjj, 1)
                    '    Dim syogo2_rel As String = relexistdata(cntjj, 2)
                    '    list_rel.Add(syogo1_rel & "-" & syogo2_rel)
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        Dim syogo1_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim
                        Dim syogo2_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(1)).Trim
                        list_rel.Add(syogo1_rel & "-" & syogo2_rel)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "鍵タイトルマスタ"
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                    ''中間ファイルデータ取得
                    ''20160823 前回設定値復元時エラー修正 -add sta
                    'If RelItem_M_kagi_Rev7 Is Nothing Then
                    '    rtn_int = 1 : Return rtn_int
                    'End If
                    ''20160823 前回設定値復元時エラー修正 -add end
                    'For cntii = 1 To RelItem_M_kagi_Rev7.ItemCnt - 1
                    '    Dim syogo1_mid As String = RelItem_M_kagi_Rev7.Kagi_no(cntii)
                    '    Dim syogo2_mid As String = RelItem_M_kagi_Rev7.Kagi_name(cntii)
                    '    list_mid.Add(syogo1_mid & "-" & syogo2_mid)
                    'Next
                    ''紐付ファイルデータ(前回設定値)取得
                    'For cntjj = 1 To relexistdatarowcnt
                    '    Dim syogo1_rel As String = relexistdata(cntjj, 1)
                    '    Dim syogo2_rel As String = relexistdata(cntjj, 2)
                    '    list_rel.Add(syogo1_rel & "-" & syogo2_rel)
                    'Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                Case "自社口座マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_jisyakoza_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_jisyakoza_Rev7.ItemCnt - 1
                        Dim syogo1_mid As String = RelItem_M_jisyakoza_Rev7.Kozasyutokumoto(cntii)
                        Dim syogo2_mid As String = RelItem_M_jisyakoza_Rev7.Kozaname(cntii)
                        Dim syogo3_mid As String = RelItem_M_jisyakoza_Rev7.Kinyu_no(cntii)
                        Dim syogo4_mid As String = RelItem_M_jisyakoza_Rev7.Kinyu_name(cntii)
                        Dim syogo5_mid As String = RelItem_M_jisyakoza_Rev7.Ten_no(cntii)
                        Dim syogo6_mid As String = RelItem_M_jisyakoza_Rev7.Ten_name(cntii)
                        Dim syogo7_mid As String = RelItem_M_jisyakoza_Rev7.Kosyu_name(cntii)
                        Dim syogo8_mid As String = RelItem_M_jisyakoza_Rev7.Koza_no(cntii)
                        Dim syogo9_mid As String = RelItem_M_jisyakoza_Rev7.Koza_meigi(cntii)
                        Dim syogo10_mid As String = RelItem_M_jisyakoza_Rev7.Koza_kana(cntii)
                        Dim syogo11_mid As String = RelItem_M_jisyakoza_Rev7.Yucyokigo1(cntii)
                        Dim syogo12_mid As String = RelItem_M_jisyakoza_Rev7.Yucyokigo2(cntii)
                        Dim syogo13_mid As String = RelItem_M_jisyakoza_Rev7.Yucyokozano(cntii)
                        Dim syogo14_mid As String = RelItem_M_jisyakoza_Rev7.biko(cntii)
                        list_mid.Add(syogo1_mid & "-" & syogo2_mid & "-" & syogo3_mid & "-" & syogo4_mid & "-" & syogo5_mid & "-" & syogo6_mid & "-" & syogo7_mid & "-" & _
                                     syogo8_mid & "-" & syogo9_mid & "-" & syogo10_mid & "-" & syogo11_mid & "-" & syogo12_mid & "-" & syogo13_mid & "-" & syogo14_mid)
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    Dim syogo1_rel As String = relexistdata(cntjj, 1)
                    '    Dim syogo2_rel As String = relexistdata(cntjj, 2)
                    '    Dim syogo3_rel As String = relexistdata(cntjj, 3)
                    '    Dim syogo4_rel As String = relexistdata(cntjj, 4)
                    '    Dim syogo5_rel As String = relexistdata(cntjj, 5)
                    '    Dim syogo6_rel As String = relexistdata(cntjj, 6)
                    '    Dim syogo7_rel As String = relexistdata(cntjj, 7)
                    '    Dim syogo8_rel As String = relexistdata(cntjj, 8)
                    '    Dim syogo9_rel As String = relexistdata(cntjj, 9)
                    '    Dim syogo10_rel As String = relexistdata(cntjj, 10)
                    '    Dim syogo11_rel As String = relexistdata(cntjj, 11)
                    '    Dim syogo12_rel As String = relexistdata(cntjj, 12)
                    '    Dim syogo13_rel As String = relexistdata(cntjj, 13)
                    '    Dim syogo14_rel As String = relexistdata(cntjj, 14)
                    '    list_rel.Add(syogo1_rel & "-" & syogo2_rel & "-" & syogo3_rel & "-" & syogo4_rel & "-" & syogo5_rel & "-" & syogo6_rel & "-" & syogo7_rel & "-" & _
                    '                 syogo8_rel & "-" & syogo9_rel & "-" & syogo10_rel & "-" & syogo11_rel & "-" & syogo12_rel & "-" & syogo13_rel & "-" & syogo14_rel)
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        Dim syogo1_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim
                        Dim syogo2_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(1)).Trim
                        Dim syogo3_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(2)).Trim
                        Dim syogo4_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(3)).Trim
                        Dim syogo5_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(4)).Trim
                        Dim syogo6_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(5)).Trim
                        Dim syogo7_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(6)).Trim
                        Dim syogo8_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(7)).Trim
                        Dim syogo9_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(8)).Trim
                        Dim syogo10_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(9)).Trim
                        Dim syogo11_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(10)).Trim
                        Dim syogo12_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(11)).Trim
                        Dim syogo13_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(12)).Trim
                        Dim syogo14_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(13)).Trim
                        list_rel.Add(syogo1_rel & "-" & syogo2_rel & "-" & syogo3_rel & "-" & syogo4_rel & "-" & syogo5_rel & "-" & syogo6_rel & "-" & syogo7_rel & "-" & _
                                     syogo8_rel & "-" & syogo9_rel & "-" & syogo10_rel & "-" & syogo11_rel & "-" & syogo12_rel & "-" & syogo13_rel & "-" & syogo14_rel)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "画像割付"
                    'For cntii = 1 To RelItem_M_brui.ItemCnt - 1
                    '    list_mid.Add(RelItem_M_brui.Brui_name(cntii))
                    'Next
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                Case "FBフォーマット割付"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_FBInfo_Furiirai_Rev7 Is Nothing Or RelItem_M_FBInfo_Kozafurikae_Rev7 Is Nothing Or RelItem_M_FBInfo_Nssetting_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_FBInfo_Furiirai_Rev7.ItemCnt - 1
                        Dim syogo1_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Fmt_syubetu(cntii)
                        Dim syogo2_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Fkom_no(cntii)
                        Dim syogo3_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Fkom_name(cntii)
                        Dim syogo4_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Kinyu_no(cntii)
                        Dim syogo5_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Kinyu_name(cntii)
                        Dim syogo6_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Ten_no(cntii)
                        Dim syogo7_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Ten_name(cntii)
                        Dim syogo8_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Kosyu_no(cntii)
                        Dim syogo9_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Koza_no(cntii)
                        Dim syogo10_mid As String = RelItem_M_FBInfo_Furiirai_Rev7.Koza_meigi(cntii)
                        list_mid.Add(syogo1_mid & "-" & syogo2_mid & "-" & syogo3_mid & "-" & syogo4_mid & "-" & syogo5_mid & "-" & _
                                     syogo6_mid & "-" & syogo7_mid & "-" & syogo8_mid & "-" & syogo9_mid & "-" & syogo10_mid)
                    Next
                    For cntii = 1 To RelItem_M_FBInfo_Kozafurikae_Rev7.ItemCnt - 1
                        Dim syogo1_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Fmt_syubetu(cntii)
                        Dim syogo2_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Fkae_no(cntii)
                        Dim syogo3_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Fkae_name(cntii)
                        Dim syogo4_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Kinyu_no(cntii)
                        Dim syogo5_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Kinyu_name(cntii)
                        Dim syogo6_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Ten_no(cntii)
                        Dim syogo7_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Ten_name(cntii)
                        Dim syogo8_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Kosyu_no(cntii)
                        Dim syogo9_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Koza_no(cntii)
                        Dim syogo10_mid As String = RelItem_M_FBInfo_Kozafurikae_Rev7.Koza_meigi(cntii)
                        list_mid.Add(syogo1_mid & "-" & syogo2_mid & "-" & syogo3_mid & "-" & syogo4_mid & "-" & syogo5_mid & "-" & _
                                     syogo6_mid & "-" & syogo7_mid & "-" & syogo8_mid & "-" & syogo9_mid & "-" & syogo10_mid)
                    Next
                    For cntii = 1 To RelItem_M_FBInfo_Nssetting_Rev7.ItemCnt - 1
                        Dim syogo1_mid As String = RelItem_M_FBInfo_Nssetting_Rev7.Fmt_syubetu(cntii)
                        Dim syogo2_mid As String = RelItem_M_FBInfo_Nssetting_Rev7.Ns_no(cntii)
                        Dim syogo3_mid As String = RelItem_M_FBInfo_Nssetting_Rev7.Ns_name(cntii)
                        Dim syogo4_mid As String = ""
                        Dim syogo5_mid As String = ""
                        Dim syogo6_mid As String = ""
                        Dim syogo7_mid As String = ""
                        Dim syogo8_mid As String = ""
                        Dim syogo9_mid As String = ""
                        Dim syogo10_mid As String = ""
                        list_mid.Add(syogo1_mid & "-" & syogo2_mid & "-" & syogo3_mid & "-" & syogo4_mid & "-" & syogo5_mid & "-" & _
                                     syogo6_mid & "-" & syogo7_mid & "-" & syogo8_mid & "-" & syogo9_mid & "-" & syogo10_mid)
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    Dim syogo1_rel As String = relexistdata(cntjj, 1)
                    '    Dim syogo2_rel As String = relexistdata(cntjj, 2)
                    '    Dim syogo3_rel As String = relexistdata(cntjj, 3)
                    '    Dim syogo4_rel As String = relexistdata(cntjj, 4)
                    '    Dim syogo5_rel As String = relexistdata(cntjj, 5)
                    '    Dim syogo6_rel As String = relexistdata(cntjj, 6)
                    '    Dim syogo7_rel As String = relexistdata(cntjj, 7)
                    '    Dim syogo8_rel As String = relexistdata(cntjj, 8)
                    '    Dim syogo9_rel As String = relexistdata(cntjj, 9)
                    '    Dim syogo10_rel As String = relexistdata(cntjj, 10)
                    '    list_rel.Add(syogo1_rel & "-" & syogo2_rel & "-" & syogo3_rel & "-" & syogo4_rel & "-" & syogo5_rel & "-" & _
                    '                 syogo6_rel & "-" & syogo7_rel & "-" & syogo8_rel & "-" & syogo9_rel & "-" & syogo10_rel)
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        Dim syogo1_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim
                        Dim syogo2_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(1)).Trim
                        Dim syogo3_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(2)).Trim
                        Dim syogo4_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(3)).Trim
                        Dim syogo5_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(4)).Trim
                        Dim syogo6_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(5)).Trim
                        Dim syogo7_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(6)).Trim
                        Dim syogo8_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(7)).Trim
                        Dim syogo9_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(8)).Trim
                        Dim syogo10_rel As String = Typ.ToStr(relexistdata.Rows(cntjj).Item(9)).Trim
                        list_rel.Add(syogo1_rel & "-" & syogo2_rel & "-" & syogo3_rel & "-" & syogo4_rel & "-" & syogo5_rel & "-" & _
                                     syogo6_rel & "-" & syogo7_rel & "-" & syogo8_rel & "-" & syogo9_rel & "-" & syogo10_rel)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "契約分類マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_Keirui Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_Keirui.ItemCnt - 1
                        list_mid.Add(RelItem_M_Keirui.Keirui_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                Case "都市計画・用途地域マスタ"
                    '中間ファイルデータ取得
                    '20160823 前回設定値復元時エラー修正 -add sta
                    If RelItem_M_TosiYoto_Rev7 Is Nothing Then
                        rtn_int = 1 : Return rtn_int
                    End If
                    '20160823 前回設定値復元時エラー修正 -add end
                    For cntii = 1 To RelItem_M_TosiYoto_Rev7.ItemCnt - 1
                        list_mid.Add(RelItem_M_TosiYoto_Rev7.TosiYoto_name(cntii))
                    Next
                    '紐付ファイルデータ(前回設定値)取得
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    'For cntjj = 1 To relexistdatarowcnt
                    '    list_rel.Add(relexistdata(cntjj, 1))
                    'Next
                    For cntjj = 0 To relexistdatarowcnt - 1
                        list_rel.Add(Typ.ToStr(relexistdata.Rows(cntjj).Item(0)).Trim)
                    Next
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
            End Select

            '*********************
            '比較処理
            '*********************
            Dim matchflg As Boolean = True
            Select Case True

                Case list_mid.Count < list_rel.Count        '中間ファイル件数 ＜ 紐付ファイル件数の場合は復元不可とする

                    rtn_int = 1
                    Return rtn_int

                Case list_mid.Count = list_rel.Count        '中間ファイル件数 = 紐付ファイル件数の場合は

                    For Each existdata In list_rel
                        If list_mid.Contains(existdata) = False Then
                            matchflg = False
                            Exit For
                        End If
                    Next

                    If matchflg Then
                        rtn_int = 0     '完全一致
                    Else
                        rtn_int = 1     '不一致
                    End If

                Case list_mid.Count > list_rel.Count

                    For Each existdata In list_rel
                        If list_mid.Contains(existdata) = False Then
                            matchflg = False
                            Exit For
                        End If
                    Next

                    If matchflg Then
                        '前回設定値が全て新しいデータに含まれているので差分を出す
                        '構築中
                        rtn_int = 2
                    Else
                        rtn_int = 1     '不一致
                    End If

            End Select

            Return rtn_int

        End Function

#End Region

#Region "Excelファイルへの出力処理"

        ''' <summary>
        ''' 出力前処理
        ''' </summary>
        ''' <remarks></remarks>
        Private Function Rel_WriteDataBefore()

            Dim rtn As Boolean = True

            '20160705 未設定項目に関するメッセージ表示機能の追加 -chg sta
            ''20160610 紐付未設定項目の紐付実行処理制御の追加 -add sta
            ''紐付項目取得
            'Dim dgv As New DataGridView
            'Dim chkcol As Integer = 0
            'Dim list_relitem As New List(Of String)
            'Call Me.Set_RelItem(list_relitem)

            'For Each relitem In list_relitem

            '    Select Case relitem
            '        Case "画像割付"
            '            dgv = Me.dgvGazo
            '            chkcol = 4
            '    End Select

            '    rtn = Me.Chk_CellBlank(dgv, chkcol)
            '    If rtn = False Then
            '        Dim notrelkomk As String = "未設定項目：" & relitem
            '        Dim notrelcol As String = "未設定列：" & dgv.Columns(chkcol).HeaderText
            '        MsgResult = MessageBox.Show(MSG_CHK_BLANK & vbCrLf & notrelkomk & vbCrLf & notrelcol, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '        Return rtn
            '    End If

            'Next
            ''20160610 紐付未設定項目の紐付実行処理制御の追加 -add end

            'MsgResult = MessageBox.Show(MSG_RELRUN_B, "確認", _
            '                                            MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            'If MsgResult = DialogResult.No Then
            '    rtn = False
            '    Return rtn
            'End If

            ''2016.03.08 紐付設定されていない項目に関して対応中のためコメントアウト -del sta
            ''If Not Chk_CellBlank() Then
            ''    Exit Sub
            ''End If
            ''2016.03.08 紐付設定されていない項目に関して対応中のためコメントアウト -del sta

            Dim cellblankflg As Boolean = False                 'True：未設定項目有り False：未設定項目無し
            Dim requiredflg As Boolean = False                  'True：紐付未設定不可(必須項目) False：紐付未設定可
            Dim list_blankkomkgrp As New List(Of String)        '未設定箇所が存在する項目を格納
            Dim list_blankkomkgrp_req As New List(Of String)    '紐付必須で未設定箇所が存在する項目を格納

            '紐付設定項目を取得
            Dim list_relitem As New List(Of String)
            Call Me.Set_RelItem(list_relitem)

            '各項目毎に未設定箇所を確認
            For Each relitem In list_relitem

                Select Case relitem

                    Case "物件分類マスタ"
                        cellblankflg = Me.Chk_CellBlank_Bkrui()
                        'requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add   '20160905 未設定項目がある場合のダイアログ修正 -del
                    Case "部屋分類マスタ"
                        cellblankflg = Me.Chk_CellBlank_Hyrui()
                        'requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add   '20160905 未設定項目がある場合のダイアログ修正 -del
                    Case "入金区分マスタ"
                        cellblankflg = Me.Chk_CellBlank_NkinKbn()
                        'requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add   '20160905 未設定項目がある場合のダイアログ修正 -del
                    Case "取引態様マスタ"
                        cellblankflg = Me.Chk_CellBlank_Toritaiyo()
                        requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add   
                    Case "入金項目マスタ"
                        cellblankflg = Me.Chk_CellBlank_Nkinkomk()
                        requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add   '20160905 未設定項目がある場合のダイアログ修正(入金項目マスタは自動設定ボタンがあるのでそちらで対応_コメントのみ)
                    Case "構造マスタ"
                        cellblankflg = Me.Chk_CellBlank_Kozo()
                        requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add   
                    Case "口座種別マスタ"
                        cellblankflg = Me.Chk_CellBlank_Kozasyu()
                        requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add
                    Case "設備マスタ"
                        cellblankflg = Me.Chk_CellBlank_Setubi()
                        'requiredflg = True      '20160829 紐付を必須にする仕様変更対応(設備は新規に移行するので必須は除外しておく) -add
                    Case "鍵タイトルマスタ"
                        cellblankflg = Me.Chk_CellBlank_Kagi()
                        'requiredflg = True      '20160829 紐付を必須にする仕様変更対応(鍵は未設定→専用鍵になるので必須は除外) -add
                    Case "自社口座マスタ"
                        cellblankflg = Me.Chk_CellBlank_Jisyakoza()
                        requiredflg = True      '20160722 口座情報紐付修正 -add
                    Case "画像割付"
                        cellblankflg = Me.Chk_CellBlank_Gazo()
                        requiredflg = True
                    Case "FBフォーマット割付"
                        cellblankflg = Me.Chk_CellBlank_FBInfo()
                        requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add
                    Case "契約分類マスタ"
                        cellblankflg = Me.Chk_CellBlank_Kyrui()
                        'requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add   '20160905 未設定項目がある場合のダイアログ修正 -del
                    Case "都市計画・用途地域マスタ"
                        cellblankflg = Me.Chk_CellBlank_TosiYoto()
                        requiredflg = True      '20160829 紐付を必須にする仕様変更対応 -add
                End Select

                '未設定箇所が存在する項目を格納
                If cellblankflg And requiredflg Then
                    list_blankkomkgrp_req.Add(relitem)              '紐付必須
                ElseIf cellblankflg And requiredflg = False Then
                    list_blankkomkgrp.Add(relitem)                  '紐付必須以外
                End If

                '紐付必須項目フラグ初期化
                requiredflg = False

            Next

            Dim tmp_msg As String = ""

            '紐付必須で未設定箇所が存在する項目が存在する場合
            If list_blankkomkgrp_req.Count <> 0 Then

                Dim relitem As String = Me.Get_JoinStr(list_blankkomkgrp_req, "、")
                tmp_msg = MSG_CHK_BLANK_REQ & vbCrLf & vbCrLf & relitem
                MsgResult = MessageBox.Show(tmp_msg, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                rtn = False
                Return rtn

            End If

            '未設定箇所が存在する項目が存在する場合
            If list_blankkomkgrp.Count <> 0 Then
                Dim relitem As String = Me.Get_JoinStr(list_blankkomkgrp, "、")
                tmp_msg = MSG_CHK_BLANK_NOTREQ & vbCrLf & vbCrLf & relitem
            Else
                tmp_msg = MSG_RELRUN_B
            End If

            MsgResult = MessageBox.Show(tmp_msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)

            If MsgResult = DialogResult.No Then
                rtn = False
                Return rtn
            End If
            '20160705 未設定項目に関するメッセージ表示機能の追加 -chg end

            Return rtn

        End Function

        ''' <summary>
        ''' 出力処理
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Rel_WriteData()

            Dim rtn As Boolean = True
            Dim normalflg As Boolean = True
            Dim dgv_relitem As New DataGridView
            Dim tmpcnt As Integer = 0

            '2016.04.11 呼出起動の修正 -add sta
            Dim obj_pgb As New ProgressBarManager
            Dim pgbform As New Form
            Dim pgbtotalcnt As Integer = 0
            Dim pgbcnt As Integer = 0
            '2016.04.11 呼出起動の修正 -add end

            '紐付項目取得
            Dim list_relitem As New List(Of String)
            Call Me.Set_RelItem(list_relitem)

            '2016.04.11 呼出起動の修正 -add sta
            Me.lblRelItemRead.Text = SITUATION_WRITE
            Me.lblRelItemRead.Visible = True
            Me.pgbRelItemRead.Visible = True
            '2016.04.11 呼出起動の修正 -add end

            '2016.04.11 呼出起動の修正 -add sta
            'プログレスバー設定
            pgbtotalcnt = list_relitem.Count
            Call obj_pgb.pgbInit(pgbtotalcnt)
            '2016.04.11 呼出起動の修正 -add end
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
            '紐付ファイル初期化
            rtn = Me.Init_Excel()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end
            'Excelファイルへ出力
            For Each relitem In list_relitem

                '2016.04.11 呼出起動の修正 -add sta
                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Exit For
                End If
                '2016.04.11 呼出起動の修正 -add end

                '項目別データ出力開始ログ出力
                Dim tmp_sql_sta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(21, relitem))
                DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_sta, tmpcnt)

                dgv_relitem = hash_relitemtogrid(relitem)
                '20160523 画像紐付設定処理の追加 -chg sta
                '画像割付はExcelファイルではなく10DBに仮テーブルを作成する
                'normalflg = Me.Write_Excel(relitem, dgv_relitem)
                Select Case relitem
                    Case "画像割付"
                        normalflg = Me.Write_DB(dgv_relitem)
                    Case Else
                        normalflg = Me.Write_Excel(relitem, dgv_relitem)
                End Select
                '20160523 画像紐付設定処理の追加 -chg end

                '項目別データ出力終了ログ出力
                Dim tmp_str As String = Chg_FlgToStr(normalflg)
                Dim tmp_sql_end As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(29, relitem, "-", tmp_str))
                DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_end, tmpcnt)

                '2016.04.11 呼出起動の修正 -add sta
                'プログレスバー設定
                pgbcnt = pgbcnt + 1
                Call obj_pgb.pgbsetting(pgbcnt)
                Me.tabPage3.Refresh()
                '2016.04.11 呼出起動の修正 -add end

            Next

            '2016.04.11 呼出起動の修正 -add sta
            If CancelFlg Then
                If singlesta = False Then
                    Me.Close()
                End If
            End If
            '2016.04.11 呼出起動の修正 -add end

        End Sub

        ''' <summary>
        ''' グリッドで設定された値をExcelへ出力
        ''' </summary>
        ''' <param name="relitem"></param>
        ''' <param name="dgv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Write_Excel(ByVal relitem As String, ByVal dgv As DataGridView) As Boolean
            '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            'Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            'Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            'Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            'Dim rtn As Boolean = True
            'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            'Dim relfldno As Integer = 0

            ''************************
            ''作業準備
            ''************************

            ''Excelファイル初期設定
            'rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, RelDirPath, REL_FILENAME, relitem)

            ''Excelファイル設定時にエラーが生じた際は処理を抜ける
            'If rtn = False Then
            '    Return rtn
            'End If

            ''************************
            ''処理開始
            ''************************

            ''DataGridViewのセルのデータ取得
            'Dim relvalue As String(,) = New String(dgv.Rows.Count - 1, dgv.Columns.Count - 1) {}

            'For cntii = 0 To dgv.Rows.Count - 1
            '    For cntjj = 0 To dgv.Columns.Count - 1

            '        Dim cellvalue As String = ""

            '        If dgv.Rows(cntii).Cells(cntjj).Value Is Nothing = False Then
            '            cellvalue = dgv.Rows(cntii).Cells(cntjj).Value.ToString()
            '        End If

            '        relvalue(cntii, cntjj) = cellvalue

            '    Next
            'Next

            ''Excelへ出力
            'Dim writestartcell As String = MIDFILE_READWRITE_COL & MIDFILE_READWRITE_ROW.ToString
            'Dim ran As String = writestartcell & ":" & Chr(Asc("A") + dgv.Columns.Count - 1) & dgv.Rows.Count + 1

            'wsheet.Range(ran).Value = relvalue

            ''************************
            ''終了処理
            ''************************

            ''Excelファイル終了設定
            'Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

            'Return rtn

            '20261005 ACE.OLEDB(Access Database Engine)を使わない読み書きに変更(64ビット対応) -chg sta
            Dim rtn As Boolean = True
            Dim relfilepath As String = EtcMethod.Set_Path(RelDirPath, REL_FILENAME & ".xlsx")

            '画面上の値を取得
            Dim rows As New List(Of String())
            For cntii = 0 To dgv.Rows.Count - 1

                Dim rowvalues(dgv.Columns.Count - 1) As String

                For cntjj = 0 To dgv.Columns.Count - 1
                    'セルの値を取得
                    If dgv.Rows(cntii).Cells(cntjj).Value Is Nothing Then
                        rowvalues(cntjj) = ""
                    Else
                        rowvalues(cntjj) = dgv.Rows(cntii).Cells(cntjj).Value.ToString
                    End If
                Next

                rows.Add(rowvalues)

            Next

            '紐付ファイルへ追加
            If rows.Count > 0 Then
                RelationExcelFile.AppendRows(relfilepath, relitem, rows)
            End If

            Return rtn
            '20261005 ACE.OLEDB(Access Database Engine)を使わない読み書きに変更(64ビット対応) -chg end
        End Function

        ''' <summary>
        ''' '20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Init_Excel() As Boolean

            Dim rtn As Boolean = True
            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'ファイルパス設定
            Dim relfilepath As String = EtcMethod.Set_Path(RelDirPath, REL_FILENAME & ".xlsx")

            '初期化処理
            Try
                'ファイル準備
                appli = CreateObject("Excel.Application")
                appli.Visible = False
                wbook = appli.Workbooks.Open(relfilepath)

                For cntii = 1 To wbook.Worksheets.Count
                    'シートを取得
                    wsheet = wbook.Worksheets(cntii)
                    '最大行を取得
                    Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count

                    If maxrowcnt > 1 Then
                        '既存データ範囲を取得
                        Dim dataarea As String = MIDFILE_READWRITE_ROW.ToString & ":" & maxrowcnt.ToString
                        wsheet.Rows(dataarea).Delete()
                    End If
                Next

            Catch ex As Exception
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return False
            End Try

            '終了処理
            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

            Return rtn

        End Function

        ''' <summary>
        ''' 画像タイトル紐付設定内容を仮テーブルへ出力 '20160523 画像紐付設定処理の追加
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Write_DB(ByVal dgv As DataGridView) As Boolean

            Dim rtn As Boolean = True
            '20160608 画像CVテストに伴う修正 -chg sta
            'Dim tmptablename As String = "tmp_gazo_rel"
            Dim tmptablename As String = "tmp_gazo_relation"
            '20160608 画像CVテストに伴う修正 -chg end
            Dim rowcnt As Integer = 0

            '仮テーブル初期化(DROP)
            Dim tmp_tmptabledropsql As String = CommonQuery.Qry_TmpTbl_Drop(tmptablename)
            DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_tmptabledropsql, rowcnt)

            '仮テーブル作成
            Dim tmp_tmptablecreatesql As String = Me.Get_CreateTblQry(tmptablename)
            DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_tmptablecreatesql, rowcnt)

            For cntii = 0 To dgv.Rows.Count - 1

                '挿入用文字列
                Dim tmp_insertitem As String = ""

                For cntjj = 0 To dgv.Columns.Count - 1

                    Dim cellvalue As String = IIf(dgv.Rows(cntii).Cells(cntjj).Value Is Nothing, "''", "'" & dgv.Rows(cntii).Cells(cntjj).Value & "'")

                    '種別の文字列を数値へ変換
                    If cntjj = 0 OrElse cntjj = 3 Then
                        cellvalue = Me.Get_GazoSyubetuToNo(cellvalue)
                    End If

                    '挿入用に連結
                    tmp_insertitem = tmp_insertitem & "," & cellvalue

                Next

                '挿入処理
                tmp_insertitem = tmp_insertitem.Remove(0, 1)
                Dim tmp_insertsql As String = " INSERT INTO " & tmptablename & " VALUES(" & tmp_insertitem & ");"
                DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_insertsql, rowcnt)

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 画像紐付設定格納用仮テーブル作成クエリ '20160523 画像紐付設定処理の追加
        ''' </summary>
        ''' <param name="tablename"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_CreateTblQry(ByVal tablename As String) As String

            Dim tmp_sql As String = ""

            '20160608 画像CVテストに伴う修正 -chg sta
            'tmp_sql = tmp_sql & " CREATE TABLE " & tablename
            'tmp_sql = tmp_sql & " 	( "
            'tmp_sql = tmp_sql & " 		[V7_gazo_syubetu] INT, "
            'tmp_sql = tmp_sql & " 		[V7_gazo_no] INT, "
            'tmp_sql = tmp_sql & " 		[V7_gazo_title] VARCHAR(100), "
            'tmp_sql = tmp_sql & " 		[10_gazo_syubetu] INT, "
            'tmp_sql = tmp_sql & " 		[10_gazo_no] INT, "
            'tmp_sql = tmp_sql & " 		[10_gazo_bunrui] VARCHAR(100)		 "
            'tmp_sql = tmp_sql & " 	); "
            tmp_sql = tmp_sql & " CREATE TABLE " & tablename
            tmp_sql = tmp_sql & " 	( "
            tmp_sql = tmp_sql & " 		[v7_gazo_syubetu] INT, "
            tmp_sql = tmp_sql & " 		[v7_gazo_no] INT, "
            tmp_sql = tmp_sql & " 		[v7_gazo_title] VARCHAR(100), "
            tmp_sql = tmp_sql & " 		[v8_gazo_syubetu] INT, "
            tmp_sql = tmp_sql & " 		[v8_gazo_no] INT, "
            tmp_sql = tmp_sql & " 		[v8_gazo_bunrui] VARCHAR(100)		 "
            tmp_sql = tmp_sql & " 	); "
            '20160608 画像CVテストに伴う修正 -chg end

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 画像種別の文字列を数値へ変換 '20160523 画像紐付設定処理の追加
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_GazoSyubetuToNo(ByVal value As String) As String

            Dim rtn_str As String = ""
            Dim tmp_str As String = value

            tmp_str = tmp_str.Replace("'", "")

            If tmp_str = "" Then
                rtn_str = "''"
                Return rtn_str
            End If

            Select Case tmp_str
                Case "物件"
                    rtn_str = "'1'"
                Case "部屋"
                    rtn_str = "'2'"
                Case "周辺"
                    rtn_str = "'3'"
            End Select

            Return rtn_str

        End Function

#End Region

#Region "設定値チェック"

        '20160610 紐付未設定項目の紐付実行処理制御の追加 -del sta
        ' ''' <summary>
        ' ''' 設定値の空白チェック(2015.07.06 sol レビュー後修正_レビュー№32)
        ' ''' </summary>
        ' ''' <remarks></remarks>
        'Private Function Chk_CellBlank()

        '    Dim rtn As Boolean = False

        '    '紐付先が固定テーブルのためデフォルト値が設定される項目
        '    Dim tmp_msgkotei As String = ""
        '    '部屋分類
        '    For cntii As Integer = 0 To dgvHyrui.RowCount - 1
        '        Dim chkcolindex As Integer = 3
        '        If dgvHyrui(chkcolindex, cntii).Value = Nothing Then
        '            tmp_msgkotei = "  ・" & tblname.Hyrui & vbCrLf
        '            Exit For
        '        End If
        '    Next
        '    '口座種別
        '    For cntii As Integer = 0 To dgvKozasyu.RowCount - 1
        '        Dim chkcolindex As Integer = 3
        '        If dgvKozasyu(chkcolindex, cntii).Value = Nothing Then
        '            tmp_msgkotei = tmp_msgkotei & "  ・" & tblname.Kozasyu & vbCrLf
        '            Exit For
        '        End If
        '    Next
        '    '物件構造
        '    For cntii As Integer = 0 To dgvKozo.RowCount - 1
        '        Dim chkcolindex As Integer = 3
        '        If dgvKozo(chkcolindex, cntii).Value = Nothing Then
        '            tmp_msgkotei = tmp_msgkotei & "  ・" & tblname.Kozo & vbCrLf
        '            Exit For
        '        End If
        '    Next
        '    '取引態様
        '    For cntii As Integer = 0 To dgvToritaiyo.RowCount - 1
        '        Dim chkcolindex As Integer = 3
        '        If dgvToritaiyo(chkcolindex, cntii).Value = Nothing Then
        '            tmp_msgkotei = tmp_msgkotei & "  ・" & tblname.Toritaiyo & vbCrLf
        '            Exit For
        '        End If
        '    Next

        '    If tmp_msgkotei <> "" Then
        '        tmp_msgkotei = MSG_CHK_BLANK_A & vbCrLf & tmp_msgkotei
        '    End If

        '    '未設定項目が新規に挿入される項目
        '    Dim tmp_msgnew As String = ""
        '    '物件分類
        '    For cntii As Integer = 0 To dgvBkrui.RowCount - 1
        '        Dim chkcolindex As Integer = 4
        '        If dgvBkrui(chkcolindex, cntii).Value = Nothing Then
        '            tmp_msgnew = tmp_msgnew & "  ・" & tblname.Bkrui & vbCrLf
        '            Exit For
        '        End If
        '    Next
        '    '入金区分
        '    For cntii As Integer = 0 To dgvNkinKbn.RowCount - 1
        '        Dim chkcolindex As Integer = 3
        '        If dgvNkinKbn(chkcolindex, cntii).Value = Nothing Then
        '            tmp_msgnew = tmp_msgnew & "  ・" & tblname.Nkinkbn & vbCrLf
        '            Exit For
        '        End If
        '    Next

        '    If tmp_msgnew <> "" Then
        '        tmp_msgnew = MSG_CHK_BLANK_B & vbCrLf & tmp_msgnew
        '    End If

        '    '未設定の場合、何も設定されない項目
        '    Dim tmp_msgnothing As String = ""
        '    '設備
        '    For cntii As Integer = 0 To dgvSetubi.RowCount - 1
        '        Dim chkcolindex As Integer = 6
        '        If dgvSetubi(chkcolindex, cntii).Value = Nothing Then
        '            tmp_msgnothing = tmp_msgnothing & "  ・" & tblname.Setubi & vbCrLf
        '            Exit For
        '        End If
        '    Next

        '    If tmp_msgnothing <> "" Then
        '        tmp_msgnothing = MSG_CHK_BLANK_C & vbCrLf & tmp_msgnothing
        '    End If

        '    If tmp_msgkotei <> "" Or tmp_msgnew <> "" Or tmp_msgnothing <> "" Then
        '        Dim tmp_msg As String = tmp_msgkotei & vbCrLf & tmp_msgnew & vbCrLf & tmp_msgnothing
        '        MsgResult = MessageBox.Show(MSG_CHK_BLANK & vbCrLf & vbCrLf & tmp_msg, "確認", _
        '                        MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
        '        If MsgResult = DialogResult.Yes Then
        '            rtn = True
        '            Return rtn
        '        End If
        '    End If

        '    Return rtn

        'End Function
        '20160610 紐付未設定項目の紐付実行処理制御の追加 -del end

        '20160705 未設定項目に関するメッセージ表示機能の追加 -del sta
        ' ''' <summary>
        ' ''' 紐付未設定値のチェック '20160610 紐付未設定項目の紐付実行処理制御の追加
        ' ''' </summary>
        ' ''' <param name="dgv"></param>
        ' ''' <param name="chkcol"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Private Function Chk_CellBlank(ByVal dgv As DataGridView, ByVal chkcol As Integer) As Boolean

        '    Dim rtn As Boolean = True

        '    Dim rowcnt As Integer = dgv.RowCount - 1

        '    For cntii = 0 To rowcnt

        '        Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
        '        If tmp_value = "" Then
        '            rtn = False
        '            Return rtn
        '        End If

        '    Next

        '    Return rtn

        'End Function
        '20160705 未設定項目に関するメッセージ表示機能の追加 -del end

        ''' <summary>
        ''' 物件分類マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Bkrui() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvBkrui
            Dim chkcol As Integer = 1
            '20160829 必須項目未設定箇所の回避処理を追加 -add sta
            Dim obj_bkrui As New Object
            obj_bkrui = New Njc.Repository.M_bk_rui_Repository
            '20160829 必須項目未設定箇所の回避処理を追加 -add end

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                If tmp_value = "" Then
                    '20160829 必須項目未設定箇所の回避処理を追加 -chg sta
                    'rtn = True
                    'Return rtn
                    'セル値チェンジイベント内でセル値を変更するためイベントハンドラを削除しておく
                    RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    Dim tmp_no As Integer = obj_bkrui.Get_Maxno(dgv)
                    Dim tmp_name As String = dgv(chkcol - 1, cntii).Value
                    dgv(chkcol, cntii).Value = tmp_no.ToString
                    dgv(chkcol + 1, cntii).Value = tmp_name
                    '削除したイベントハンドラを再度追加
                    AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    '20160829 必須項目未設定箇所の回避処理を追加 -chg end
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 部屋分類マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Hyrui() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvHyrui
            Dim chkcol As Integer = 1
            '20160829 必須項目未設定箇所の回避処理を追加 -add sta
            Dim obj_hyrui As New Object
            obj_hyrui = New Njc.Repository.M_hy_rui_Repository
            '20160829 必須項目未設定箇所の回避処理を追加 -add end

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                If tmp_value = "" Then
                    '20160829 必須項目未設定箇所の回避処理を追加 -chg sta
                    'rtn = True
                    'Return rtn
                    'セル値チェンジイベント内でセル値を変更するためイベントハンドラを削除しておく
                    RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    Dim tmp_no As Integer = obj_hyrui.Get_Maxno(dgv)
                    Dim tmp_name As String = dgv(chkcol - 1, cntii).Value
                    dgv(chkcol, cntii).Value = tmp_no.ToString
                    dgv(chkcol + 1, cntii).Value = tmp_name
                    '削除したイベントハンドラを再度追加
                    AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    '20160829 必須項目未設定箇所の回避処理を追加 -chg end
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 入金区分マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_NkinKbn() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvNkinKbn
            Dim chkcolmain As Integer = 1
            Dim chkcolsub As Integer = 3
            '20160829 必須項目未設定箇所の回避処理を追加 -add sta
            Dim obj_nkbn As New Object
            obj_nkbn = New Njc.Repository.M_nkbn_Repository
            '20160829 必須項目未設定箇所の回避処理を追加 -add end
            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value_main As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_value_sub As String = IIf(dgv(chkcolsub, cntii).Value Is Nothing, "", dgv(chkcolsub, cntii).Value)
                '20160829 必須項目未設定箇所の回避処理を追加 -chg sta
                'If tmp_value_main = "" Or tmp_value_sub = "" Then
                '    rtn = True
                '    Return rtn
                'End If
                'セル値チェンジイベント内でセル値を変更するためイベントハンドラを削除しておく
                RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                If tmp_value_main = "" Then
                    Dim tmp_no As Integer = obj_nkbn.Get_Maxno(dgv)
                    Dim tmp_name As String = dgv(chkcolmain - 1, cntii).Value
                    dgv(chkcolmain, cntii).Value = tmp_no.ToString
                    dgv(chkcolmain + 1, cntii).Value = tmp_name
                    dgv(chkcolsub, cntii).Value = "4"
                    dgv(chkcolsub + 1, cntii).Value = "その他"
                ElseIf tmp_value_sub = "" Then
                    dgv(chkcolsub, cntii).Value = "4"
                    dgv(chkcolsub + 1, cntii).Value = "その他"
                End If
                '削除したイベントハンドラを再度追加
                AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                '20160829 必須項目未設定箇所の回避処理を追加 -chg end
            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 取引態様マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Toritaiyo() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvToritaiyo
            Dim chkcol As Integer = 1

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                If tmp_value = "" Then
                    rtn = True
                    Return rtn
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 入金項目マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Nkinkomk() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvNkinkomk
            Dim nkinkbncol As Integer = 1
            Dim chkcolmain As Integer = 2
            Dim chkcolsub1 As Integer = 4
            Dim chkcolsub2 As Integer = 7

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim nkinkbn As String = IIf(dgv(nkinkbncol, cntii).Value Is Nothing, "", dgv(nkinkbncol, cntii).Value)
                Dim tmp_value_main As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_value_sub1 As String = IIf(dgv(chkcolsub1, cntii).Value Is Nothing, "", dgv(chkcolsub1, cntii).Value)
                Dim tmp_value_sub2 As String = IIf(dgv(chkcolsub2, cntii).Value Is Nothing, "", dgv(chkcolsub2, cntii).Value)

                If tmp_value_main = "" Or tmp_value_sub1 = "" Then
                    rtn = True
                    Return rtn
                ElseIf nkinkbn = "5.随時変動" And tmp_value_sub2 = "" Then
                    rtn = True
                    Return rtn
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 物件構造マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Kozo() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvKozo
            Dim chkcol As Integer = 1

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                If tmp_value = "" Then
                    rtn = True
                    Return rtn
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 口座種別マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Kozasyu() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvKozasyu
            Dim chkcol As Integer = 1

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                If tmp_value = "" Then
                    rtn = True
                    Return rtn
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 設備マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Setubi() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvSetubi
            Dim chkcolmain As Integer = 2
            Dim chkcolsub1 As Integer = 4
            Dim chkcolsub2 As Integer = 6

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value_main As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_value_sub1 As String = IIf(dgv(chkcolsub1, cntii).Value Is Nothing, "", dgv(chkcolsub1, cntii).Value)
                Dim tmp_value_sub2 As String = IIf(dgv(chkcolsub2, cntii).Value Is Nothing, "", dgv(chkcolsub2, cntii).Value)

                If tmp_value_main = "" Or tmp_value_sub1 = "" Or tmp_value_sub2 = "" Then
                    rtn = True
                    Return rtn
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 鍵タイトルマスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Kagi() As Boolean

        End Function

        ''' <summary>
        ''' 自社口座マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Jisyakoza() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvJisyakoza
            Dim chkcolmain As Integer = 15          '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
            Dim chkcolsub1 As Integer = 17          '20160829 自社口座にデフォルト値を設定する処理を追加 16列→17列
            Dim list_narabino As New List(Of String)    '20160825 並び順の任意No自動設定処理を追加 -add

            Dim rowcnt As Integer = dgv.RowCount - 1

            '20160825 並び順の任意No自動設定処理を追加 -add sta
            '使用されている並び順をオブジェクトへ格納
            For cntii = 0 To rowcnt
                Dim tmp_basevalue As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_narabino As String = IIf(dgv(chkcolsub1, cntii).Value Is Nothing, "", dgv(chkcolsub1, cntii).Value)
                If tmp_narabino <> "" Then
                    list_narabino.Add(tmp_basevalue & "-" & tmp_narabino)
                End If
            Next
            '20160825 並び順の任意No自動設定処理を追加 -add end

            For cntii = 0 To rowcnt

                Dim tmp_value_main As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_value_sub1 As String = IIf(dgv(chkcolsub1, cntii).Value Is Nothing, "", dgv(chkcolsub1, cntii).Value)

                '20160825 並び順の任意No自動設定処理を追加 -chg sta
                'If tmp_value_main = "" Or tmp_value_sub1 = "" Then
                '    rtn = True
                '    Return rtn
                'End If

                If tmp_value_main = "" Then         '自社No未設定時は未設定項目として扱う
                    rtn = True
                    Return rtn
                ElseIf tmp_value_sub1 = "" Then     '並び順未設定時は任意のNoを自動で設定する
                    Dim naibuno As Integer = 0
                    naibuno = Me.Get_NiniNo(list_narabino, tmp_value_main)
                    RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    dgv(chkcolsub1, cntii).Value = naibuno.ToString
                    AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                End If
                '20160825 並び順の任意No自動設定処理を追加 -chg end
            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 周辺画像割付の紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Gazo() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvGazo
            '20160825 並び順の任意No自動設定処理を追加 -chg sta
            'Dim chkcol As Integer = 4
            Dim chkcolmain As Integer = 3
            Dim chkcolsub1 As Integer = 4
            '20160825 並び順の任意No自動設定処理を追加 -chg end
            Dim list_narabino As New List(Of String)    '20160825 並び順の任意No自動設定処理を追加 -add

            Dim rowcnt As Integer = dgv.RowCount - 1

            '20160825 並び順の任意No自動設定処理を追加 -add sta
            '使用されている並び順をオブジェクトへ格納
            For cntii = 0 To rowcnt
                Dim tmp_basevalue As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_narabino As String = IIf(dgv(chkcolsub1, cntii).Value Is Nothing, "", dgv(chkcolsub1, cntii).Value)
                If tmp_narabino <> "" Then
                    list_narabino.Add(tmp_basevalue & "-" & tmp_narabino)
                End If
            Next
            '20160825 並び順の任意No自動設定処理を追加 -add end

            For cntii = 0 To rowcnt

                '20160825 並び順の任意No自動設定処理を追加 -chg sta
                'Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                'If tmp_value = "" Then
                '    rtn = True
                '    Return rtn
                'End If
                Dim tmp_value_main As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_value_sub1 As String = IIf(dgv(chkcolsub1, cntii).Value Is Nothing, "", dgv(chkcolsub1, cntii).Value)
                If tmp_value_sub1 = "" Then
                    Dim naibuno As Integer = 0
                    naibuno = Me.Get_NiniNo(list_narabino, tmp_value_main)
                    RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    dgv(chkcolsub1, cntii).Value = naibuno.ToString
                    AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                End If
                '20160825 並び順の任意No自動設定処理を追加 -chg end
            Next

            Return rtn

        End Function

        ''' <summary>
        ''' FBフォーマットの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_FBInfo() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvFBInfo
            Dim fmtkbncol As Integer = 0
            Dim chkcolmain As Integer = 10
            Dim chkcolsub1 As Integer = 12

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim fmtkbn As String = IIf(dgv(fmtkbncol, cntii).Value Is Nothing, "", dgv(fmtkbncol, cntii).Value)
                Dim tmp_value_main As String = IIf(dgv(chkcolmain, cntii).Value Is Nothing, "", dgv(chkcolmain, cntii).Value)
                Dim tmp_value_sub1 As String = IIf(dgv(chkcolsub1, cntii).Value Is Nothing, "", dgv(chkcolsub1, cntii).Value)

                If tmp_value_main = "" Then
                    rtn = True
                    Return rtn
                ElseIf fmtkbn = "入出金" And tmp_value_sub1 = "" Then
                    rtn = True
                    Return rtn
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 契約分類マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_Kyrui() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvKyrui
            Dim chkcol As Integer = 1
            '20160829 必須項目未設定箇所の回避処理を追加 -add sta
            Dim obj_kyrui As New Object
            obj_kyrui = New Njc.Repository.M_ky_rui_Repository
            '20160829 必須項目未設定箇所の回避処理を追加 -add end
            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                If tmp_value = "" Then
                    '20160829 必須項目未設定箇所の回避処理を追加 -chg sta
                    'rtn = True
                    'Return rtn
                    'セル値チェンジイベント内でセル値を変更するためイベントハンドラを削除しておく
                    RemoveHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    Dim tmp_no As Integer = obj_kyrui.Get_Maxno(dgv)
                    Dim tmp_name As String = dgv(chkcol - 1, cntii).Value
                    dgv(chkcol, cntii).Value = tmp_no.ToString
                    dgv(chkcol + 1, cntii).Value = tmp_name
                    dgv(chkcol + 2, cntii).Value = "FALSE"
                    '削除したイベントハンドラを再度追加
                    AddHandler dgv.CellValueChanged, AddressOf Me.dataGridView_CellValueChanged
                    '20160829 必須項目未設定箇所の回避処理を追加 -chg end
                End If

            Next

            Return rtn

        End Function

        ''' <summary>
        ''' 都市計画・用途地域マスタの紐付未設定値チェック '20160705 未設定項目に関するメッセージ表示機能の追加 -add
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_CellBlank_TosiYoto() As Boolean

            Dim rtn As Boolean = False
            Dim dgv As DataGridView = Me.dgvTosiYoto
            Dim chkcol As Integer = 2

            Dim rowcnt As Integer = dgv.RowCount - 1

            For cntii = 0 To rowcnt

                Dim tmp_value As String = IIf(dgv(chkcol, cntii).Value Is Nothing, "", dgv(chkcol, cntii).Value)
                If tmp_value = "" Then
                    rtn = True
                    Return rtn
                End If

            Next

            Return rtn

        End Function

#End Region

#Region "DB接続情報設定"

        ''' <summary>
        ''' V7DB接続情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_V7DBInfo()

            'V7接続情報を読み取り専用で取得
            Dim regkey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\VB and VBA Program Settings\Fkanri50\Fk5Db", False)

            '20160614 V7レジストリ情報が存在しない場合の処理を追加 -chg sta
            ''値を取得 (存在しない場合は空文字)
            'Dim svname As String = DirectCast(regkey.GetValue("DATA SOURCE", ""), String)
            'Dim dbname As String = DirectCast(regkey.GetValue("INITIAL CATALOG", ""), String)
            'Dim username As String = DirectCast(regkey.GetValue("UID", ""), String)
            'Dim password As String = DirectCast(regkey.GetValue("PWD", ""), String)

            ''パスワード変換
            'password = Me.Get_V7DBPassword(password)

            Dim svname As String = ""
            Dim dbname As String = ""
            Dim username As String = ""
            Dim chgpassword As String = ""
            Dim password As String = ""

            If regkey Is Nothing Then
                'レジストリ情報が存在しない場合は空文字で出力
            Else
                '値を取得 (存在しない場合は空文字)
                svname = DirectCast(regkey.GetValue("DATA SOURCE", ""), String)
                dbname = DirectCast(regkey.GetValue("INITIAL CATALOG", ""), String)
                username = DirectCast(regkey.GetValue("UID", ""), String)
                chgpassword = DirectCast(regkey.GetValue("PWD", ""), String)
                'パスワード変換
                password = Me.Get_V7DBPassword(chgpassword)
            End If
            '20160614 V7レジストリ情報が存在しない場合の処理を追加 -chg end

            'モデルへ格納
            With fstmodelv7
                .ServerName = svname
                .InitialCatalog = dbname
                .User = username
                .Pass = password
            End With

        End Sub

        ''' <summary>
        ''' パスワード取得 (V7プログラムから引用)
        ''' </summary>
        ''' <param name="pass"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_V7DBPassword(ByVal pass As String) As String

            Dim rtn_str As String

            Dim PSW As String = "ogaryo"
            Dim tmp_pass As String = ""
            Dim n1 As Integer = 0
            Dim n2 As Integer = 0
            Dim n3 As Integer = 0

            '制御番号を形成
            For i = 1 To Len(PSW)
                n1 = n1 + Asc(Mid(PSW, i, 1))
                n1 = (n1 * 367 + 331) Mod &HFFF
                n2 = ((n2 + n1) * 743 + 599) Mod &HFFF
                n3 = ((n3 + n2) * 563 + 787) Mod &HFFF
            Next i

            '変換
            tmp_pass = pass
            Chg_Password(tmp_pass, n1, n2, n3)
            rtn_str = tmp_pass

            Return rtn_str

        End Function

        ''' <summary>
        ''' パスワード変換 (V7プログラムから引用)
        ''' </summary>
        ''' <param name="str"></param>
        ''' <param name="int1"></param>
        ''' <param name="int2"></param>
        ''' <param name="int3"></param>
        ''' <remarks></remarks>
        Private Sub Chg_Password(ByRef str As String, ByVal int1 As Integer, ByVal int2 As Integer, ByVal int3 As Integer)

            Dim R As Integer
            Dim M As Integer
            Dim N As Integer
            Dim BigNum As Integer = 32768       '疑似乱数の周期（２の巾乗）
            Dim i As Integer
            Dim c As Integer
            Dim d As Integer

            R = int1                            'オプションの乱数出発点
            M = (int2 * 4 + 1) Mod BigNum       'オプションの乗数初期化値
            N = (int3 * 2 + 1) Mod BigNum       'オプションの付加項初期化値

            ' 文字列を処理します
            For i = 1 To Len(str)
                c = Asc(Mid(str, i, 1))
                ' プリント可能な文字の一部分だけを変更します
                Select Case c
                    Case 48 To 57
                        d = c - 48
                    Case 63 To 90
                        d = c - 53
                    Case 97 To 122
                        d = c - 59
                    Case Else
                        d = -1
                End Select
                ' １文字を暗号化する準備をします
                If d >= 0 Then
                    ' 次の疑似乱数を生成します
                    R = (R * M + N) Mod BigNum
                    ' 文字と疑似乱数のビットをXorします
                    d = (R And 63) Xor d
                    ' 文字をプリント可能な範囲に戻します
                    Select Case d
                        Case 0 To 9
                            c = d + 48
                        Case 10 To 37
                            c = d + 53
                        Case 38 To 63
                            c = d + 59
                    End Select
                    ' 元の文字を結果で置き換えます
                    Mid(str, i, 1) = Chr(c)
                End If
            Next i
        End Sub

        ''' <summary>
        ''' 10DB接続情報取得 2016.02.22 10DB接続情報取得処理追加(途中まで)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_10DBInfo()

            '初期接続情報が記載されたXMLファイルパス
            Dim xmlfilepath As String = "C:\Users\sol\AppData\Roaming\n-create.co.jp\FK8Host\FK8Host.sql.Main.xml"

            'ファイル有無チェック
            If Me.Chk_FileExist(xmlfilepath) = False Then
                Exit Sub
            End If

            'ファイル読込
            Dim xmlreader As System.Xml.XmlReader = System.Xml.XmlReader.Create(xmlfilepath)
            Dim item As String = ""
            Dim data As String = ""
            Dim svname As String = ""
            Dim dbname As String = ""
            Dim username As String = ""
            Dim password As String = ""
            Dim passkey As String = "tRwmj5U4"
            Dim tmp_pass As String = ""

            While xmlreader.Read
                If xmlreader.NodeType = Xml.XmlNodeType.Element Then

                    'データ取得
                    item = xmlreader.LocalName
                    data = xmlreader.ReadString

                    'それぞれの要素で分岐
                    Select Case item
                        Case "ServerName"
                            svname = data
                        Case "InitialCatalog"
                            dbname = data
                        Case "UserID"
                            username = data
                    End Select
                End If
            End While

            'パスワード暗号化(保留)
            'Dim aaa As String = Me.Encrypt(tmp_pass, passkey)

            password = "p7s2#c1j3n"

            'モデルへ格納
            With fstmodelv10
                .ServerName = svname
                .InitialCatalog = dbname
                .User = username
                .Pass = password
            End With

            xmlreader.Close()

        End Sub

        ''' <summary>
        ''' DESで暗号化します。<br/> 2016.02.22 10DB接続情報取得処理追加(途中まで)
        ''' </summary>
        Public Function Encrypt(ByVal target As String, ByVal key As String) As String
            Dim targetBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(target)

            Dim des As New Security.Cryptography.DESCryptoServiceProvider
            Dim keyBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(key)
            des.Key = resizeBytesArray(keyBytes, des.Key.Length)
            des.IV = resizeBytesArray(keyBytes, des.IV.Length)

            Dim outStream As New IO.MemoryStream
            Dim desencrypt As Security.Cryptography.ICryptoTransform = des.CreateEncryptor
            Dim cryptStream As New Security.Cryptography.CryptoStream(outStream, desencrypt, Security.Cryptography.CryptoStreamMode.Write)

            cryptStream.Write(targetBytes, 0, targetBytes.Length)
            cryptStream.FlushFinalBlock()
            Dim outBytes As Byte() = outStream.ToArray

            cryptStream.Close()
            outStream.Close()

            Return Convert.ToBase64String(outBytes)
        End Function

        ''' <summary>
        ''' DESで復号化します。<br/> 2016.02.22 10DB接続情報取得処理追加(途中まで)
        ''' </summary>
        Public Function Decrypt(ByVal encryptedTarget As String, ByVal key As String) As String
            Dim encryptedBytes As Byte() = Convert.FromBase64String(encryptedTarget)

            Dim des As New Security.Cryptography.DESCryptoServiceProvider
            Dim keyBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(key)
            des.Key = resizeBytesArray(keyBytes, des.Key.Length)
            des.IV = resizeBytesArray(keyBytes, des.IV.Length)

            Dim inStream As New IO.MemoryStream(encryptedBytes)
            Dim desdecrypt As Security.Cryptography.ICryptoTransform = des.CreateDecryptor
            Dim cryptStream As New Security.Cryptography.CryptoStream(inStream, desdecrypt, Security.Cryptography.CryptoStreamMode.Read)

            Dim outStream As New IO.StreamReader(cryptStream, System.Text.Encoding.UTF8)
            Dim result As String = outStream.ReadToEnd

            outStream.Close()
            cryptStream.Close()
            inStream.Close()

            Return result
        End Function

        Private Function resizeBytesArray(ByVal bytes() As Byte, _
                                ByVal newSize As Integer) As Byte()
            Dim newBytes(newSize - 1) As Byte
            If bytes.Length <= newSize Then
                Dim i As Integer
                For i = 0 To bytes.Length - 1
                    newBytes(i) = bytes(i)
                Next i
            Else
                Dim pos As Integer = 0
                Dim i As Integer
                For i = 0 To bytes.Length - 1
                    newBytes(pos) = newBytes(pos) Xor bytes(i)
                    pos += 1
                    If pos >= newBytes.Length Then
                        pos = 0
                    End If
                Next i
            End If
            Return newBytes
        End Function

        ''' <summary>
        ''' 接続情報：初期値→コントロール
        ''' </summary>
        ''' <param name="cnnv7"></param>
        ''' <param name="cnnv10"></param>
        ''' <remarks></remarks>
        Public Sub Set_DefConInfo_To_Control(ByRef cnnv7 As Njc.Model.DefSQLConnection, ByRef cnnv10 As Njc.Model.DefSQLConnection)

            'V7
            Me.txtV7Server.Text = cnnv7.ServerName
            Me.txtV7Catalog.Text = cnnv7.InitialCatalog
            Me.txtV7Networklib.Text = cnnv7.NetworkLibrary
            Me.txtV7User.Text = cnnv7.User
            Me.txtV7Pass.Text = cnnv7.Pass

            '10
            Me.txtV10Server.Text = cnnv10.ServerName
            Me.txtV10Catalog.Text = cnnv10.InitialCatalog
            Me.txtV10Networklib.Text = cnnv10.NetworkLibrary
            Me.txtV10User.Text = cnnv10.User
            Me.txtV10Pass.Text = cnnv10.Pass

            '共通
            Me.txtTimeOut.Text = cnnv10.TimeOut

        End Sub

        ''' <summary>
        ''' 接続情報：コントロール→モデル
        ''' </summary>
        ''' <param name="model"></param>
        ''' <param name="flg"></param>
        ''' <remarks></remarks>
        Public Sub Set_Control_To_ConModel(ByRef model As Njc.Model.DefSQLConnection, Optional ByVal flg As Integer = 0)

            If flg = 1 Then
                '10
                model.ServerName = Me.txtV10Server.Text
                model.InitialCatalog = Me.txtV10Catalog.Text
                model.NetworkLibrary = Me.txtV10Networklib.Text
                model.User = Me.txtV10User.Text
                model.Pass = Me.txtV10Pass.Text
            Else
                'V7
                model.ServerName = Me.txtV7Server.Text
                model.InitialCatalog = Me.txtV7Catalog.Text
                model.NetworkLibrary = Me.txtV7Networklib.Text
                model.User = Me.txtV7User.Text
                model.Pass = Me.txtV7Pass.Text
            End If

            '共通
            model.TimeOut = Me.txtTimeOut.Text

        End Sub

#End Region

#Region "その他処理"

        ''' <summary>
        ''' ボタン状態変更処理
        ''' </summary>
        ''' <param name="status"></param>
        ''' <remarks></remarks>
        Private Sub Chg_BtnStatus(ByVal status As Integer)

            If singlesta Then
                btnBack.Visible = True
            Else
                btnBack.Visible = False
            End If

            btnNext.Visible = True
            btnEnd.Visible = True
            btnBack.Enabled = True
            btnNext.Enabled = True
            btnEnd.Enabled = True

            Select Case status

                Case 0
                    btnBack.Text = ""
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False

                Case 1
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"
                    If Me.btnConnectTest.Enabled Then
                        btnNext.Enabled = False
                    End If

                Case 2
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   中  止"

                Case 3
                    '※設定したテキスト名を条件に使用している箇所あり
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   実  行"
                    btnEnd.Text = "   中  止"

                Case 4
                    btnBack.Text = ""
                    btnNext.Text = "ログを開く"
                    btnEnd.Text = "   終  了"
                    btnBack.Visible = False

                Case 10
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   読  込"
                    btnEnd.Text = "   中  止"

                Case 32
                    '※設定したテキスト名を条件に使用している箇所あり
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   実行中"
                    btnEnd.Text = "キャンセル"
                    btnBack.Enabled = False
                    btnNext.Enabled = False

                Case 33
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   次  へ"
                    btnEnd.Text = "   終  了"
                    btnBack.Visible = False
                    btnEnd.Visible = False

                    '2016.04.11 呼出起動の修正 -add sta
                Case 35
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   読込中"
                    btnEnd.Text = "キャンセル"
                    btnBack.Visible = False
                    btnNext.Enabled = False

                Case 36
                    '※設定したテキスト名を条件に使用している箇所あり
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   実  行"
                    btnEnd.Text = "   中  止"
                    btnBack.Visible = False

                Case 37
                    btnBack.Text = "   戻  る"
                    btnNext.Text = "   実行中"
                    btnEnd.Text = "キャンセル"
                    btnBack.Visible = False
                    btnNext.Enabled = False
                    '2016.04.11 呼出起動の修正 -add end

                Case Else

            End Select

        End Sub

        ''' <summary>
        ''' タブオーダー(メインタブ)設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_TabOrder_Main(ByVal typeno As Integer)

            Dim order As Integer = 0

            Select Case typeno
                Case 0  '初期画面
                    Me.tabCtrlRelMain.TabIndex = order : order = order + 1
                    Me.chkRelFile.TabIndex = order : order = order + 1
                    Me.chkTempTable.TabIndex = order : order = order + 1
                    Me.txtRelationDirPath.TabIndex = order : order = order + 1
                    Me.btnRelDirSeach.TabIndex = order : order = order + 1
                    Me.txtLogDirPath.TabIndex = order : order = order + 1
                    Me.btnLogDirSeach.TabIndex = order : order = order + 1
                    Me.btnBack.TabIndex = order : order = order + 1
                    Me.btnNext.TabIndex = order : order = order + 1
                    Me.btnEnd.TabIndex = order : order = order + 1
                Case 1  '接続情報画面
                    '↓コンテナを先に設定しておく↓
                    Me.grpV7Authent.TabIndex = order : order = order + 1
                    Me.grpV10Authent.TabIndex = order : order = order + 1
                    'kakaka6 test chg s --------------------------------------------------
                    'Me.grpConnectInfo.TabIndex = order : order = order + 1
                    Me.grpV7ConnectInfo.TabIndex = order : order = order + 1
                    Me.grp10ConnectInfo.TabIndex = order : order = order + 1
                    'kakaka6 test chg e --------------------------------------------------
                    Me.tabCtrlRelMain.TabIndex = order : order = order + 1
                    '↑コンテナを先に設定しておく↑
                    Me.optV7Authent1.TabIndex = order : order = order + 1
                    Me.optV7Authent2.TabIndex = order : order = order + 1
                    Me.txtV7Server.TabIndex = order : order = order + 1
                    Me.txtV7Catalog.TabIndex = order : order = order + 1
                    Me.txtV7Networklib.TabIndex = order : order = order + 1
                    Me.txtV7User.TabIndex = order : order = order + 1
                    Me.txtV7Pass.TabIndex = order : order = order + 1
                    Me.optV10Authent1.TabIndex = order : order = order + 1
                    Me.optV10Authent2.TabIndex = order : order = order + 1
                    Me.txtV10Server.TabIndex = order : order = order + 1
                    Me.txtV10Catalog.TabIndex = order : order = order + 1
                    Me.txtV10Networklib.TabIndex = order : order = order + 1
                    Me.txtV10User.TabIndex = order : order = order + 1
                    Me.txtV10Pass.TabIndex = order : order = order + 1
                    Me.txtTimeOut.TabIndex = order : order = order + 1
                    Me.btnConnectTest.TabIndex = order : order = order + 1
                    Me.btnBack.TabIndex = order : order = order + 1
                    Me.btnNext.TabIndex = order : order = order + 1
                    Me.btnEnd.TabIndex = order : order = order + 1
                Case 2  '紐付設定画面
                    '紐付設定画面は別関数(Set_TabOrder_Sub)で処理する
                Case 3  '終了画面
                    Me.tabCtrlRelMain.TabIndex = order : order = order + 1
                    Me.btnBack.TabIndex = order : order = order + 1
                    Me.btnNext.TabIndex = order : order = order + 1
                    Me.btnEnd.TabIndex = order : order = order + 1
            End Select

        End Sub

        ''' <summary>
        ''' タブオーダー(グリッド表示タブ)設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_TabOrder_Sub(ByVal dgv As DataGridView)

            Dim order As Integer = 0

            Me.tabCtrlRelMain.TabIndex = order : order = order + 1
            Me.chkDuplicate.TabIndex = order : order = order + 1
            dgv.TabIndex = order : order = order + 1
            Me.btnBack.TabIndex = order : order = order + 1
            Me.btnNext.TabIndex = order : order = order + 1
            Me.btnEnd.TabIndex = order : order = order + 1

        End Sub

        ''' <summary>
        ''' フォルダ有無チェック
        ''' </summary>
        ''' <param name="dirpath"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_DirExist(ByVal dirpath As String) As Boolean

            Dim rtn = True

            If Not (Directory.Exists(dirpath)) Then
                rtn = False
                Return rtn
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' ファイル有無チェック
        ''' </summary>
        ''' <param name="filepath"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_FileExist(ByVal filepath As String) As Boolean

            Dim rtn = True

            'ファイル有無確認
            If Not (File.Exists(filepath)) Then
                rtn = False
                Return rtn
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' ファイル/フォルダのフルパスを設定
        ''' </summary>
        ''' <param name="path"></param>
        ''' <param name="name"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Set_Path(ByVal path As String, ByVal name As String) As String

            Dim rtn_path As String = ""

            If EtcMethod.RightStrMethod(path, 1) = "\" Then
                rtn_path = path & name
            Else
                rtn_path = path & "\" & name
            End If

            Return rtn_path

        End Function

        ''' <summary>
        ''' 紐付項目を取得してオブジェクトへ格納
        ''' </summary>
        ''' <param name="list"></param>
        ''' <remarks></remarks>
        Private Sub Set_RelItem(ByRef list As List(Of String))

            '2016.04.11 呼出起動の修正 -chg sta
            'Dim tmp_chkname As String = ""

            ''タブコントロール内
            'For Each item As Control In Me.tabPage2.Controls

            '    If item.GetType().Equals(GetType(CheckBox)) Then

            '        Dim control_chk As CheckBox = DirectCast(item, CheckBox)

            '        If control_chk.Checked Then

            '            '移行項目取得
            '            tmp_chkname = control_chk.Text

            '            'リスト格納
            '            list.Add(tmp_chkname)

            '        End If

            '    End If

            'Next

            '単体/呼出起動で処理を分岐させる
            If list_relitem_call.Count <> 0 Then
                '呼出起動時は取得しておいた紐付項目をセット
                For Each relitem In list_relitem_call
                    list.Add(relitem)
                Next
            Else
                '単体起動時は選択された紐付項目から取得
                Dim tmp_chkname As String = ""

                'タブコントロール内
                For Each item As Control In Me.tabPage2.Controls

                    If item.GetType().Equals(GetType(CheckBox)) Then

                        Dim control_chk As CheckBox = DirectCast(item, CheckBox)

                        If control_chk.Checked Then

                            '移行項目取得
                            tmp_chkname = control_chk.Text

                            'リスト格納
                            list.Add(tmp_chkname)

                        End If

                    End If

                Next

            End If
            '2016.04.11 呼出起動の修正 -chg end

        End Sub

        ''' <summary>
        ''' 設定されたフォルダ有無チェック
        ''' </summary>
        ''' <param name="errdirname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chk_SettingDir(ByRef errdirname As String) As Boolean

            Dim rtn As Boolean = True

            '読込元中間ファイル有無チェック
            Dim tmp_middirpath As String = Me.txtMidDirPath.Text
            If Me.Chk_DirExist(tmp_middirpath) Then
                MidDirPath = tmp_middirpath
            Else
                errdirname = Me.lblMidDir.Text
                rtn = False
            End If

            '出力先紐付ファイル有無チェック
            Dim tmp_reldirpath As String = Me.txtRelationDirPath.Text
            Dim tmp_relfllepath As String = Me.Set_Path(tmp_reldirpath, REL_FILENAME & ".xlsx")
            If Me.Chk_FileExist(tmp_relfllepath) Then
                RelDirPath = tmp_reldirpath
            Else
                errdirname = errdirname & "、" & Me.lblRelationDir.Text
                rtn = False
            End If

            'ログファイル格納先有無チェック
            Dim tmp_logdirpath As String = Me.txtLogDirPath.Text
            If Me.Chk_DirExist(tmp_logdirpath) Then
                Dim logfilename As String = "cvrel_log_" & Replace((Replace(Replace((String.Format(Now)), ":", ""), "/", "")), " ", "") & ".csv"
                LogDirPath = tmp_logdirpath
                LogFilePath = Me.Set_Path(LogDirPath, logfilename)

            Else
                errdirname = errdirname & "、" & Me.lblLogDir.Text
                rtn = False
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 論理値を文字列へ変換
        ''' </summary>
        ''' <param name="flg">正常(True)/異常(False)終了フラグ</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Chg_FlgToStr(ByVal flg As Boolean) As String

            Dim rtn_str As String = ""

            If flg Then
                rtn_str = LOG_NAIYO_NORMALEND
            Else
                rtn_str = LOG_NAIYO_NOTNORMALEND
            End If

            Return rtn_str

        End Function

        ''' <summary>
        ''' コマンドライン引数を各オブジェクトへ格納
        ''' </summary>
        ''' <remarks>
        '''・コマンドライン
        '''    P1 … 単独起動選択(0.呼出起動, 0以外.単体起動)
        '''　　P2 … コンバートタイプ(0.汎用, 1.既存ユーザ用(DEF))
        '''　　P3 … V7接続情報
        '''　　P4 … 10接続情報(タイムアウト値を含む)
        '''　　P5 … 紐付設定ファイル格納先フォルダ
        '''　　P6 … 中間ファイル格納先フォルダ
        '''　　P7 … 紐付項目 (「-」で連結した文字列)
        ''' </remarks>
        Private Sub Set_CmdlineInfoToObj()

            '呼出起動時に取得したコマンドラインの接続情報をオブジェクトへ格納

            '20160725 紐付設定改善 chg -sta
            ' ''20160616 コマンドライン引数の修正 -chg sta
            '' ''cmdlineの中身
            '' ''1～5  … V7接続情報
            '' ''6～10 … 10接続情報
            '' ''11    … タイムアウト値
            '' ''12    … 紐付設定ファイル格納先フォルダ
            '' ''13    … 中間ファイル格納先フォルダ
            '' ''14    … 紐付項目 (「-」で連結した文字列)

            '' ''V7接続情報
            ' ''fstmodelv7.ServerName = cmdline(2)
            ' ''fstmodelv7.InitialCatalog = cmdline(3)
            ' ''fstmodelv7.User = cmdline(4)
            ' ''fstmodelv7.Pass = cmdline(5)
            ' ''Call cnnV7.CnnSession(fstmodelv7, sqlcnnV7, cmdline(1))

            '' ''10接続情報
            ' ''fstmodelv10.ServerName = cmdline(7)
            ' ''fstmodelv10.InitialCatalog = cmdline(8)
            ' ''fstmodelv10.User = cmdline(9)
            ' ''fstmodelv10.Pass = cmdline(10)
            ' ''Call cnnV10.CnnSession(fstmodelv10, sqlcnnV10, cmdline(6))

            '' ''タイムアウト値
            ' ''fstmodelv10.TimeOut = cmdline(11)

            '' ''紐付設定ファイル格納先を取得
            ' ''RelDirPath = cmdline(12)

            '' ''中間ファイル格納先を取得
            ' ''MidDirPath = cmdline(13)

            '' ''紐付項目取得
            ' ''If UBound(cmdline) >= 14 Then

            ' ''    '「-」で連結した文字列を分割
            ' ''    Dim tmp_relitemgrp() As String = cmdline(14).Split("-")

            ' ''    For cntii = 0 To UBound(tmp_relitemgrp)

            ' ''        Dim tmp_relitem As String = tmp_relitemgrp(cntii)

            ' ''        Select Case tmp_relitem
            ' ''            Case "bkruichk"
            ' ''                list_relitem_call.Add(Me.chkBkrui.Text)
            ' ''            Case "kozochk"
            ' ''                list_relitem_call.Add(Me.chkKozo.Text)
            ' ''            Case "kagichk"
            ' ''                list_relitem_call.Add(Me.chkKagi.Text)
            ' ''            Case "tosiyotochk"
            ' ''                list_relitem_call.Add(Me.chkTosiYoto.Text)
            ' ''            Case "hyruichk"
            ' ''                list_relitem_call.Add(Me.chkHyrui.Text)
            ' ''            Case "kyruichk"
            ' ''                list_relitem_call.Add(Me.chkKyrui.Text)
            ' ''            Case "nkinkbnchk"
            ' ''                list_relitem_call.Add(Me.chkNkinkbn.Text)
            ' ''            Case "toritaiyochk"
            ' ''                list_relitem_call.Add(Me.chkToritaiyo.Text)
            ' ''            Case "nkinkomkchk"
            ' ''                list_relitem_call.Add(Me.chkNkinkomk.Text)
            ' ''            Case "setubichk"
            ' ''                list_relitem_call.Add(Me.chkSetubi.Text)
            ' ''            Case "sohokenchk"
            ' ''                '構築中
            ' ''                'list_callrelitem.Add(Me.chkSohoken.Text)
            ' ''            Case "kozasyuchk"
            ' ''                list_relitem_call.Add(Me.chkKozasyu.Text)
            ' ''                '2016.04.26 メインの方へも反映させる修正 -add sta
            ' ''            Case "jisyakozachk"
            ' ''                list_relitem_call.Add(Me.chkJisyakoza.Text)
            ' ''                '2016.04.26 メインの方へも反映させる修正 -add end
            ' ''            Case "gazochk"  '20160523 画像紐付設定処理の追加 -add
            ' ''                list_relitem_call.Add(Me.chkGazo.Text)
            ' ''            Case "fbfmtchk"  '20160525 FBフォーマット紐付設定処理の追加 -add
            ' ''                list_relitem_call.Add(Me.chkFBFmt.Text)
            ' ''        End Select

            ' ''    Next

            ' ''End If

            ' ''cmdlineの中身
            ' ''1     … V7接続情報
            ' ''2     … 10接続情報(タイムアウト値を含む)
            ' ''3    … 紐付設定ファイル格納先フォルダ
            ' ''4    … 中間ファイル格納先フォルダ
            ' ''5    … 紐付項目 (「-」で連結した文字列)

            ' ''V7接続情報
            ''Dim tmp_v7con() As String = cmdline(1).Split(",")
            ''fstmodelv7.ServerName = tmp_v7con(1)
            ''fstmodelv7.InitialCatalog = tmp_v7con(2)
            ''fstmodelv7.User = tmp_v7con(3)
            ''fstmodelv7.Pass = tmp_v7con(4)
            ''Call cnnV7.CnnSession(fstmodelv7, sqlcnnV7, tmp_v7con(0))

            ' ''10接続情報
            ''Dim tmp_10con() As String = cmdline(2).Split(",")
            ''fstmodelv10.ServerName = tmp_10con(1)
            ''fstmodelv10.InitialCatalog = tmp_10con(2)
            ''fstmodelv10.User = tmp_10con(3)
            ''fstmodelv10.Pass = tmp_10con(4)
            ''Call cnnV10.CnnSession(fstmodelv10, sqlcnnV10, tmp_10con(0))

            ' ''タイムアウト値
            ''fstmodelv10.TimeOut = tmp_10con(5)

            ' ''紐付設定ファイル格納先を取得
            ''RelDirPath = cmdline(3)

            ' ''中間ファイル格納先を取得
            ''MidDirPath = cmdline(4)

            ' ''紐付項目取得
            ''If UBound(cmdline) >= 5 Then

            ''    '「-」で連結した文字列を分割
            ''    Dim tmp_relitemgrp() As String = cmdline(5).Split("-")

            ''    For cntii = 0 To UBound(tmp_relitemgrp)

            ''        Dim tmp_relitem As String = tmp_relitemgrp(cntii)

            ''        Select Case tmp_relitem
            ''            Case "bkruichk"
            ''                list_relitem_call.Add(Me.chkBkrui.Text)
            ''            Case "kozochk"
            ''                list_relitem_call.Add(Me.chkKozo.Text)
            ''            Case "kagichk"
            ''                list_relitem_call.Add(Me.chkKagi.Text)
            ''            Case "tosiyotochk"
            ''                list_relitem_call.Add(Me.chkTosiYoto.Text)
            ''            Case "hyruichk"
            ''                list_relitem_call.Add(Me.chkHyrui.Text)
            ''            Case "kyruichk"
            ''                list_relitem_call.Add(Me.chkKyrui.Text)
            ''            Case "nkinkbnchk"
            ''                list_relitem_call.Add(Me.chkNkinkbn.Text)
            ''            Case "toritaiyochk"
            ''                list_relitem_call.Add(Me.chkToritaiyo.Text)
            ''            Case "nkinkomkchk"
            ''                list_relitem_call.Add(Me.chkNkinkomk.Text)
            ''            Case "setubichk"
            ''                list_relitem_call.Add(Me.chkSetubi.Text)
            ''            Case "sohokenchk"
            ''                '構築中
            ''                'list_callrelitem.Add(Me.chkSohoken.Text)
            ''            Case "kozasyuchk"
            ''                list_relitem_call.Add(Me.chkKozasyu.Text)
            ''                '2016.04.26 メインの方へも反映させる修正 -add sta
            ''            Case "jisyakozachk"
            ''                list_relitem_call.Add(Me.chkJisyakoza.Text)
            ''                '2016.04.26 メインの方へも反映させる修正 -add end
            ''            Case "gazochk"  '20160523 画像紐付設定処理の追加 -add
            ''                list_relitem_call.Add(Me.chkGazo.Text)
            ''            Case "fbfmtchk"  '20160525 FBフォーマット紐付設定処理の追加 -add
            ''                list_relitem_call.Add(Me.chkFBFmt.Text)
            ''        End Select

            ''    Next

            ''End If
            ' ''20160616 コマンドライン引数の修正 -chg end

            'コンバートタイプ
            If Int32.TryParse(cmdline(2).ToString, CNVNO) = False Then
                CNVNO = 1
            End If

            'V7接続情報
            Dim tmp_v7con() As String = cmdline(3).Split(",")
            fstmodelv7.ServerName = tmp_v7con(1)
            fstmodelv7.InitialCatalog = tmp_v7con(2)
            fstmodelv7.User = tmp_v7con(3)
            fstmodelv7.Pass = tmp_v7con(4)
            '20261005 V7接続情報が空の場合は接続しない(起動時に接続タイムアウトまで待たされるため) -chg sta
            '※kit2(汎用)からの呼出ではV7接続情報は常に空で、V7は使用しない
            'Call cnnV7.CnnSession(fstmodelv7, sqlcnnV7, tmp_v7con(0))
            If fstmodelv7.ServerName.Trim = "" Then
                sqlcnnV7 = New System.Data.SqlClient.SqlConnection()
            Else
                Call cnnV7.CnnSession(fstmodelv7, sqlcnnV7, tmp_v7con(0))
            End If
            '20261005 V7接続情報が空の場合は接続しない(起動時に接続タイムアウトまで待たされるため) -chg end

            '10接続情報
            Dim tmp_10con() As String = cmdline(4).Split(",")
            fstmodelv10.ServerName = tmp_10con(1)
            fstmodelv10.InitialCatalog = tmp_10con(2)
            fstmodelv10.User = tmp_10con(3)
            fstmodelv10.Pass = tmp_10con(4)
            Call cnnV10.CnnSession(fstmodelv10, sqlcnnV10, tmp_10con(0))

            'タイムアウト値
            fstmodelv10.TimeOut = tmp_10con(5)

            '紐付設定ファイル格納先を取得
            RelDirPath = cmdline(5)

            '中間ファイル格納先を取得
            MidDirPath = cmdline(6)

            '紐付項目取得
            If UBound(cmdline) >= 7 Then

                '「-」で連結した文字列を分割
                Dim tmp_relitemgrp() As String = cmdline(7).Split("-")

                For cntii = 0 To UBound(tmp_relitemgrp)

                    Dim tmp_relitem As String = tmp_relitemgrp(cntii)

                    Select Case tmp_relitem
                        Case "bkruichk"
                            list_relitem_call.Add(Me.chkBkrui.Text)
                        Case "kozochk"
                            list_relitem_call.Add(Me.chkKozo.Text)
                        Case "kagichk"
                            list_relitem_call.Add(Me.chkKagi.Text)
                        Case "tosiyotochk"
                            list_relitem_call.Add(Me.chkTosiYoto.Text)
                        Case "hyruichk"
                            list_relitem_call.Add(Me.chkHyrui.Text)
                        Case "kyruichk"
                            list_relitem_call.Add(Me.chkKyrui.Text)
                        Case "nkinkbnchk"
                            list_relitem_call.Add(Me.chkNkinkbn.Text)
                        Case "toritaiyochk"
                            list_relitem_call.Add(Me.chkToritaiyo.Text)
                        Case "nkinkomkchk"
                            list_relitem_call.Add(Me.chkNkinkomk.Text)
                        Case "setubichk"
                            list_relitem_call.Add(Me.chkSetubi.Text)
                        Case "sohokenchk"
                            'list_callrelitem.Add(Me.chkSohoken.Text)
                        Case "kozasyuchk"
                            list_relitem_call.Add(Me.chkKozasyu.Text)
                        Case "jisyakozachk"
                            list_relitem_call.Add(Me.chkJisyakoza.Text)
                        Case "gazochk"
                            list_relitem_call.Add(Me.chkGazo.Text)
                        Case "fbfmtchk"
                            list_relitem_call.Add(Me.chkFBFmt.Text)
                    End Select

                Next
                '20160829 紐付ツール起動位置修正_本体→紐付 -add sta
                '表示位置を取得
                Dim tmp_point() As String = cmdline(8).Split(",")
                Me.StartPosition = FormStartPosition.Manual
                Me.Location = New Point(Int32.Parse(tmp_point(0)), Int32.Parse(tmp_point(1)))
                '20160829 紐付ツール起動位置修正_本体→紐付 -add end
                '20161004 自社口座の口座種別取得処理の修正 -add sta
                '汎用用中間ファイル格納先フォルダパスを取得
                '20161019 画像からの呼出時のエラー修正 -chg sta
                'BaseMidDirPath = cmdline(9)
                If UBound(cmdline) >= 9 Then
                    BaseMidDirPath = cmdline(9)
                End If
                '20161019 画像からの呼出時のエラー修正 -chg end
                '20161004 自社口座の口座種別取得処理の修正 -add end
            End If
            '20160725 紐付設定改善 chg -end

        End Sub

        ''' <summary>
        ''' 紐付項目格納用のオブジェクトを初期化 '20160525 全体的な動作の修正
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_RelItemObjInit()

            'V7項目格納オブジェクト
            RelItem_M_brui = Nothing
            RelItem_M_crui = Nothing
            RelItem_M_kozasyu_Rev7 = Nothing
            RelItem_M_kozo_Rev7 = Nothing
            RelItem_M_nkbn_Rev7 = Nothing
            RelItem_M_nkin_Rev7 = Nothing
            RelItem_M_setubi_Rev7 = Nothing
            RelItem_M_toritaiyo_Rev7 = Nothing
            RelItem_M_kagi_Rev7 = Nothing
            RelItem_M_jisyakoza_Rev7 = Nothing
            RelItem_M_gazo_title_Rev7 = Nothing
            RelItem_M_FBInfo_Furiirai_Rev7 = Nothing
            RelItem_M_FBInfo_Kozafurikae_Rev7 = Nothing
            RelItem_M_FBInfo_Nssetting_Rev7 = Nothing
            RelItem_M_Keirui = Nothing
            RelItem_M_TosiYoto_Rev7 = Nothing   '20160823 前回設定値復元時エラー修正 -add

            '10項目格納オブジェクト
            RelItem_M_bk_rui = Nothing
            RelItem_M_hy_rui = Nothing
            RelItem_M_kozasyu = Nothing
            RelItem_M_kozo = Nothing
            RelItem_M_nkbn = Nothing
            RelItem_M_nkbn_z = Nothing
            RelItem_M_nkin = Nothing
            RelItem_M_nkin_z = Nothing
            RelItem_M_nkin_hendometer = Nothing
            RelItem_M_setubi_Grp = Nothing
            RelItem_M_setubi_Ms = Nothing
            RelItem_M_setubi_Komk = Nothing
            RelItem_M_setubi = Nothing
            RelItem_M_toritaiyo = Nothing
            RelItem_M_kagi = Nothing
            RelItem_M_jisyakoza = Nothing
            RelItem_M_gazo_title_kbn = Nothing
            RelItem_M_gazo_title_name = Nothing
            RelItem_M_FBInfo = Nothing
            RelItem_M_ky_rui = Nothing
            RelItem_M_Tosi = Nothing       '20160823 前回設定値復元時エラー修正 -add
            RelItem_M_Yoto = Nothing       '20160823 前回設定値復元時エラー修正 -add

        End Sub

        ''' <summary>
        ''' リストオブジェクトに格納された文字を任意の文字列で連結する '20160705 未設定項目に関するメッセージ表示機能の追加 -chg end
        ''' </summary>
        ''' <param name="list"></param>
        ''' <param name="joinstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_JoinStr(ByVal list As List(Of String), ByVal joinstr As String) As String

            Dim rtn_str As String = ""
            Dim tmp_str As String = ""

            For Each item In list
                tmp_str = tmp_str & joinstr & item
            Next

            If tmp_str <> "" Then
                tmp_str = tmp_str.Remove(0, 1)
            End If

            rtn_str = tmp_str

            Return rtn_str

        End Function

        '20160725 紐付設定改善 add -sta
        ''' <summary>
        ''' 紐付設定タブ初期化(全タブ非表示)
        ''' </summary>
        ''' <param name="initflg">初期FLG…TRUE.タブページ表示用オブジェクトインスタンス化</param>
        ''' <remarks></remarks>
        Public Sub Chg_TabPageSub(ByVal initflg As Boolean)

            If initflg Then
                tabPageManager_Sub = New Njc.Common.TabPageManager(tabCtrlRelwork)
            End If

            '起動時は全て非表示にしておく (選択された項目のみ表示させるため)
            tabPageManager_Sub.ChangeTabPageVisible(0, False)
            tabPageManager_Sub.ChangeTabPageVisible(1, False)
            tabPageManager_Sub.ChangeTabPageVisible(2, False)
            tabPageManager_Sub.ChangeTabPageVisible(3, False)
            tabPageManager_Sub.ChangeTabPageVisible(4, False)
            tabPageManager_Sub.ChangeTabPageVisible(5, False)
            tabPageManager_Sub.ChangeTabPageVisible(6, False)
            tabPageManager_Sub.ChangeTabPageVisible(7, False)
            tabPageManager_Sub.ChangeTabPageVisible(8, False)
            tabPageManager_Sub.ChangeTabPageVisible(9, False)
            tabPageManager_Sub.ChangeTabPageVisible(10, False)  '20160523 画像紐付設定処理の追加 -add
            tabPageManager_Sub.ChangeTabPageVisible(11, False)  '20160525 FBフォーマット紐付設定処理の追加 -add
            tabPageManager_Sub.ChangeTabPageVisible(12, False)  '20160526 契約分類マスタの追加 -add
            tabPageManager_Sub.ChangeTabPageVisible(13, False)  '20160531 都市計画用途地域の追加 -add

            '選択項目ラベルを非活性状態にする
            Call Chg_IniLabelColor()

            '件数表示ラベルを初期化する '20160825 紐付画面件数表示処理対応 -add
            Call Me.Chg_IniLabelText()

        End Sub

        ''' <summary>
        ''' 選択項目ラベルを非活性状態に設定
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Chg_IniLabelColor()

            For Each ctrl As Object In pnlRelSelectPanel.Controls
                '20160825 紐付画面件数表示処理対応 -chg sta
                'If ctrl.GetType().Equals(GetType(Label)) AndAlso ctrl.Name <> escRelLabel Then
                '    ctrl.ForeColor = escFColor
                '    ctrl.BackColor = Color.White
                '    'ctrl.Enable = False
                'End If
                If ctrl.GetType().Equals(GetType(Label)) AndAlso List_escRelLabel.Contains(ctrl.Name) = False Then
                    ctrl.ForeColor = escFColor
                    Console.WriteLine(ctrl.text)
                End If
                '20160825 紐付画面件数表示処理対応 -chg end
            Next
        End Sub

        ''' <summary>
        ''' 選択項目ラベルの背景色を変更
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Chg_IniBackColor()

            For Each ctrl As Object In pnlRelSelectPanel.Controls
                '20160825 紐付画面件数表示処理対応 -chg sta
                'If ctrl.GetType().Equals(GetType(Label)) AndAlso ctrl.Name <> escRelLabel AndAlso ctrl.BackColor = escBColor Then
                '    ctrl.BackColor = Color.White
                'End If
                If ctrl.GetType().Equals(GetType(Label)) AndAlso List_escRelLabel.Contains(ctrl.Name) = False AndAlso ctrl.BackColor = escBColor Then
                    ctrl.BackColor = Color.White
                End If
                '20160825 紐付画面件数表示処理対応 -chg end
            Next
        End Sub
        '20160725 紐付設定改善 add -end

        ''' <summary>
        ''' 件数表示ラベルの初期化 '20160825 紐付画面件数表示処理対応 -add
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Chg_IniLabelText()

            Dim list_cntlblname As New List(Of String) From _
                {"lblCntSumiTosiYoto", "lblCntSumiKyBunrui", "lblCntSumiNkKbn", "lblCntSumiFbFmt", "lblCntSumiTaiyo", _
                 "lblSumiZanGazo", "lblCntSumiNkin", "lblCntSumiJisyaKoza", "lblCntSumiKozo", "lblCntSumiKyoyoKagi", _
                 "lblCntSumiKozaSyubetu", "lblCntSumiSetubi", "lblCntSumiBkBunrui", "lblCntSumiHyBunrui", "lblCntAllTosiYoto", _
                 "lblCntAllKyBunrui", "lblCntAllFbFmt", "lblCntAllNkKbn", "lblCntAllGazo", "lblCntAllTaiyo", _
                 "lblCntAllJisyaKoza", "lblCntAllNkin", "lblCntAllKyoyoKagi", "lblCntAllKozo", "lblCntAllSetubi", _
                 "lblCntAllKozaSyubetu", "lblCntAllBkBunrui", "lblCntAllHyBunrui"}

            For Each ctrl As Object In pnlRelSelectPanel.Controls
                If ctrl.GetType().Equals(GetType(Label)) AndAlso list_cntlblname.Contains(ctrl.Name) Then
                    ctrl.text = ""
                End If
            Next

        End Sub

        ''' <summary>
        ''' 任意のNoを生成する '20160825 並び順の任意No自動設定処理を追加 -chg end
        ''' </summary>
        ''' <param name="list_no"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Get_NiniNo(ByRef list_no As List(Of String), Optional basevalue As String = "") As Integer

            Dim rtn_int As Integer = 0

            Dim containflg As Boolean = True
            Dim setnkinno As Integer = 0
            Do Until containflg = False
                rtn_int = rtn_int + 1
                containflg = list_no.Contains(basevalue & "-" & rtn_int.ToString)
            Loop
            list_no.Add(basevalue & "-" & rtn_int.ToString)

            Return rtn_int

        End Function

        ''' <summary>
        ''' セルコンボボックスのプルダウンリストの表示幅を設定する
        ''' </summary>
        ''' <param name="cellcombo">クリックした箇所のセルコンボボックスオブジェクト</param>
        ''' <remarks>
        ''' 20161227 コンボボックスプルダウンの表示幅修正 新規追加
        ''' 　コンボボックスに含まれている文字列から最大バイト数を取得
        ''' 　取得したバイト数を元に幅を設定(1バイト当たり幅8にする)
        ''' </remarks>
        Private Sub Set_CmbListWidth(ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックスに含まれている文字列の最大バイト数を取得する
            Dim max_byte As Integer = 0
            For Each item In cellcombo.Items
                Dim tmp_str As String = item.ToString.Replace(" ", "*").Replace("　", "**")
                Dim tmp_byte As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(tmp_str)
                If max_byte <= tmp_byte Then
                    max_byte = tmp_byte
                End If
            Next
            If max_byte = 0 Then
                Exit Sub
            End If

            '1バイト当たり幅8を設定
            Dim wid As Integer = 8
            cellcombo.DropDownWidth = max_byte * wid

        End Sub

#End Region

#Region "ログ関連設定"

        ''' <summary>
        ''' ログ出力初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_LogInit()

            Dim rowcnt As Integer = 0

            'ログ格納用テーブル初期化(DROP)
            Dim tmp_sql_drop As String = CommonQuery.Qry_TmpTbl_Drop(LOG_REL_TABLENAME)
            DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_drop, rowcnt)

            'ログ格納用テーブル生成
            Dim tmp_sql_create As String = LogSetting.Get_LogTblCreateQry()
            DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_create, rowcnt)

        End Sub

        ''' <summary>
        ''' ログ出力設定
        ''' </summary>
        ''' <param name="logdropflg"></param>
        ''' <remarks></remarks>
        Private Sub Set_LogFile(ByVal logdropflg As Boolean)

            Dim rowcnt As Integer = 0

            'ログファイル出力
            '2016.03.23 単体起動時と呼出起動時でログの出力処理を分岐させる -chg sta
            'If Me.Chk_FileExist(LogFilePath) = False Then

            '    'ログファイル(CSV)のデータ部出力
            '    Dim tmp_sql As String = LogSetting.Get_LogTblSelectQry()
            '    Dim rtn As Boolean = True
            '    Dim errstr As String = ""
            '    rtn = FileMethod.TblView_Output_CSV(Me.sqlcnnV10, LogFilePath, "", "", errstr, tmp_sql, True)

            '    'ヘッダーを加えて加工
            '    Call Me.Set_LogHeader(LogFilePath)

            'End If
            If singlesta = False Then

                '呼出起動時はコンバーターのログテーブルへ挿入する (一部のフィールドサイズが紐付＞本体側になっているので合わせること)
                Dim tmp_sql_tableinsert As String = LogSetting.Get_LogTblInsertQry()
                DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_tableinsert, rowcnt)

            ElseIf Me.Chk_FileExist(LogFilePath) = False Then

                '単体起動時は直接ログファイルを出力する
                'ログファイル(CSV)のデータ部出力
                Dim tmp_sql As String = LogSetting.Get_LogTblSelectQry()
                Dim rtn As Boolean = True
                Dim errstr As String = ""
                rtn = FileMethod.TblView_Output_CSV(Me.sqlcnnV10, LogFilePath, "", "", errstr, tmp_sql, True)

                'ヘッダーを加えて加工
                Call Me.Set_LogHeader(LogFilePath)

            End If
            '2016.03.23 単体起動時と呼出起動時でログの出力処理を分岐させる -chg end

            'ログ格納用テーブル削除(DROP)
            If logdropflg Then
                'コンバートログ
                Dim tmp_sql_logdrop As String = CommonQuery.Qry_TmpTbl_Drop(LOG_REL_TABLENAME)
                DBExec.Exec_NonQuery(Me.sqlcnnV10, tmp_sql_logdrop, rowcnt)
            End If

        End Sub

        ''' <summary>
        ''' ログファイル出力時にヘッダを挿入する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_LogHeader(ByVal filepath As String)

            Dim strtype As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")     '文字コード設定
            Dim strread As New StreamReader(filepath, strtype)                                   'ログファイルを開く
            Dim tmp_path As String = Path.GetTempFileName()                                         '仮ファイル作成
            Dim strwrite As New StreamWriter(tmp_path, False, strtype)                              '仮ファイルを開く

            '仮ファイルへヘッダー書込み
            strwrite.WriteLine(LOG_HEADER_TOTAL)

            '内容読込
            While strread.Peek() > -1
                Dim line As String = strread.ReadLine()
                strwrite.WriteLine(line)
            End While

            '終了処理
            strread.Close()
            strwrite.Close()

            '仮ファイルと入れ替え
            System.IO.File.Copy(tmp_path, filepath, True)
            System.IO.File.Delete(tmp_path)

        End Sub

#End Region

#Region "セルの着色設定"

        ''' <summary>
        ''' 選択行の着色 '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -add
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_Select_All(ByVal dgv As DataGridView, ByVal rowcurrent As Integer)

            For cntcol = 0 To dgv.ColumnCount - 1
                dgv(cntcol, rowcurrent).Style.BackColor = Color.Pink
            Next

        End Sub

        ''' <summary>
        ''' 選択セル移動時に元に着色を元に戻す '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -add
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="bordercol"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_Leave_All(ByVal dgv As DataGridView, ByVal rowcurrent As Integer, ByVal bordercol As Integer)

            For cntcol = 0 To dgv.ColumnCount - 1
                If cntcol < bordercol Then
                    dgv(cntcol, rowcurrent).Style.BackColor = Color.Azure
                Else
                    dgv(cntcol, rowcurrent).Style.BackColor = Color.Beige
                End If
            Next

        End Sub

        ''' <summary>
        ''' 入金項目の随時変動費のグレイアウト着色設定 '20160701 指摘事項まとめファイルの対応 随時変動費以外の変動費メーター分類列のグレイアウト対応 -add sta
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_Nkinkomk_Zuiji(ByVal dgv As DataGridView)

            For cntkk = 1 To dgv.RowCount
                Dim tmp_nkinkbn As String = dgv(1, cntkk - 1).Value
                If tmp_nkinkbn <> "5.随時変動" Then
                    dgv(7, cntkk - 1).Style.BackColor = Color.LightGray
                    dgv(8, cntkk - 1).Style.BackColor = Color.LightGray
                End If
            Next

        End Sub

        ''' <summary>
        ''' 入金項目コードが一括で設定されたセルの着色制御 '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add
        ''' 変更された場合は着色を元に戻す
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_Nkinkomk_CellChg(ByVal dgv As DataGridView, ByVal rowcurrent As Integer)

            If chgcolorrow.Contains(rowcurrent) Then
                If hash_gridtocolno.Contains(dgv) Then
                    Dim stacolno As Integer = 0
                    stacolno = hash_gridtocolno(dgv)
                    For cntcol = stacolno To dgv.ColumnCount - 1
                        If cntcol <= 5 Then
                            dgv(cntcol, rowcurrent).Style.BackColor = Color.Beige
                        End If
                    Next
                End If
                chgcolorrow.Remove(rowcurrent)
            End If

        End Sub

        ''' <summary>
        ''' 入金項目コードが一括で設定されたセルの着色制御 '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add
        ''' 変更されていない場合は一括設定時の着色を維持する
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="bordercol"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_Nkinkomk_CellNotChg(ByVal dgv As DataGridView, ByVal rowcurrent As Integer, ByVal bordercol As Integer)  '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add

            For cntcol = 0 To dgv.ColumnCount - 1
                If cntcol < bordercol Then
                    dgv(cntcol, rowcurrent).Style.BackColor = Color.Azure
                ElseIf chgcolorrow.Contains(rowcurrent) Then
                    dgv(cntcol, rowcurrent).Style.BackColor = Color.NavajoWhite
                Else
                    dgv(cntcol, rowcurrent).Style.BackColor = Color.Beige
                End If
            Next

        End Sub

        ''' <summary>
        ''' FBフォーマットの入手金情報以外の自社設定箇所のグレイアウト着色 '20160701 指摘事項まとめファイルの対応 セル着色処理の全体的な修正 -add
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_FBFmt(ByVal dgv As DataGridView)

            For cntkk = 1 To dgv.RowCount
                Dim tmp_fmtkbn As String = dgv(0, cntkk - 1).Value
                If tmp_fmtkbn <> "入出金" Then
                    dgv(12, cntkk - 1).Style.BackColor = Color.LightGray
                    dgv(13, cntkk - 1).Style.BackColor = Color.LightGray
                End If
            Next

        End Sub

        ''' <summary>
        ''' 周辺画像の施設分類名のグレイアウト着色 '20160701 指摘事項まとめファイルの対応 画像種別が物件時のセル着色修正 -add
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_Gazo(ByVal dgv As DataGridView)

            For cntkk = 1 To dgv.RowCount
                Dim tmp_10_syubetu As String = dgv(3, cntkk - 1).Value
                If tmp_10_syubetu = "物件" Then
                    dgv(5, cntkk - 1).Style.BackColor = Color.LightGray
                End If
            Next

        End Sub

        ''' <summary>
        ''' 部屋設備エレベーター選択時のグレイアウト着色 '20160829 エレベーター移行対応 -add
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Private Sub Set_Color_Setubi(ByVal dgv As DataGridView)

            For cntkk = 1 To dgv.RowCount
                Dim tmp_setubigrp As String = dgv(2, cntkk - 1).Value
                Dim tmp_setubi As String = dgv(4, cntkk - 1).Value
                If tmp_setubigrp = "999" Or tmp_setubi = "999" Then
                    dgv(6, cntkk - 1).Style.BackColor = Color.LightGray
                    dgv(7, cntkk - 1).Style.BackColor = Color.LightGray
                End If
            Next

        End Sub

        ''' <summary>
        ''' 入金区分読取専用制御 20160829 入金区分紐付不可行制御処理を追加
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Color_NkbnReadOnly(ByVal dgv As DataGridView)

            Dim list_readonlykbn As New List(Of String) From {"現金", "振込", "振替"}

            With dgv
                For cntii = 0 To .RowCount - 1
                    Dim tmp_kbnname As String = dgv(0, cntii).Value
                    If list_readonlykbn.Contains(tmp_kbnname) Then
                        For cntjj = 0 To .ColumnCount - 1
                            dgv(cntjj, cntii).Style.BackColor = Color.LightGray
                        Next
                    End If
                Next
            End With

        End Sub

#End Region

#Region "※汎用コンバート用処理"   '20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_紐付 -add

        ''' <summary>
        ''' 汎用コンバート時の画面制御 20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_紐付 -add
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Set_HanyoGamen()

            '紐付設定画面

            '右上部パネル内
            Me.Label22.Text = "コンバート元データ(登録データ)"

            '契約分類
            lblRelSelectKyBunrui.Location = lblRelSelectHySetubi.Location
            lblCntKyBunruiBack.Location = lblCntSetubiBack.Location
            lblCntSumiKyBunrui.Location = lblCntSumiSetubi.Location
            lblCntAllKyBunrui.Location = lblCntAllSetubi.Location
            '20161012 汎用時に紐付対象外となる項目非表示修正 -del sta
            ''都市計画・用途地域
            'lblRelSelectTosiYoto.Location = lblRelSelectJisyaKoza.Location
            'lblCntTosiYotoBack.Location = lblCntJisyaKozaBack.Location
            'lblCntSumiTosiYoto.Location = lblCntSumiJisyaKoza.Location
            'lblCntAllTosiYoto.Location = lblCntAllJisyaKoza.Location
            '20161012 汎用時に紐付対象外となる項目非表示修正 -del end
            '設備
            lblRelSelectHySetubi.Visible = False
            lblCntSetubiBack.Visible = False
            lblCntSumiSetubi.Visible = False
            lblCntAllSetubi.Visible = False

            '自社口座
            lblRelSelectJisyaKoza.Visible = False
            lblCntJisyaKozaBack.Visible = False
            lblCntSumiJisyaKoza.Visible = False
            lblCntAllJisyaKoza.Visible = False

            '周辺画像
            lblRelSelectGazo.Visible = False
            lblCntGazoBack.Visible = False
            lblSumiZanGazo.Visible = False
            lblCntAllGazo.Visible = False

            'FBフォーマット
            lblRelSelectFBFmt.Visible = False
            lblCntFbFmtBack.Visible = False
            lblCntSumiFbFmt.Visible = False
            lblCntAllFbFmt.Visible = False
            '20161012 汎用時に紐付対象外となる項目非表示修正 -add sta
            '都市計画・用途地域
            lblRelSelectTosiYoto.Visible = False
            lblCntTosiYotoBack.Visible = False
            lblCntSumiTosiYoto.Visible = False
            lblCntAllTosiYoto.Visible = False
            '20161012 汎用時に紐付対象外となる項目非表示修正 -add end
        End Sub

#End Region

    End Class

End Namespace
