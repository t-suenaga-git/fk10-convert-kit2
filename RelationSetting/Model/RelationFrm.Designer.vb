Namespace Njc.Frm

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class RelationFrm
        Inherits System.Windows.Forms.Form

        'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
        <System.Diagnostics.DebuggerNonUserCode()> _
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        'Windows フォーム デザイナーで必要です。
        Private components As System.ComponentModel.IContainer

        'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
        'Windows フォーム デザイナーを使用して変更できます。  
        'コード エディターを使って変更しないでください。
        <System.Diagnostics.DebuggerStepThrough()> _
        Private Sub InitializeComponent()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RelationFrm))
            Me.tabCtrlRelMain = New System.Windows.Forms.TabControl()
            Me.tabPage0 = New System.Windows.Forms.TabPage()
            Me.grpKagicntgrp = New System.Windows.Forms.GroupBox()
            Me.lblRelSelectKyoyoKagi = New System.Windows.Forms.Label()
            Me.lblCntKyoyoKagiBack = New System.Windows.Forms.Label()
            Me.lblCntSumiKyoyoKagi = New System.Windows.Forms.Label()
            Me.lblCntAllKyoyoKagi = New System.Windows.Forms.Label()
            Me.Label47 = New System.Windows.Forms.Label()
            Me.lblLine0 = New System.Windows.Forms.Label()
            Me.btnMidDirSeach = New System.Windows.Forms.Button()
            Me.txtMidDirPath = New System.Windows.Forms.TextBox()
            Me.lblMidDir = New System.Windows.Forms.Label()
            Me.grpDev = New System.Windows.Forms.GroupBox()
            Me.chkLogTblDrop = New System.Windows.Forms.CheckBox()
            Me.Label11 = New System.Windows.Forms.Label()
            Me.chkTempTable = New System.Windows.Forms.CheckBox()
            Me.chkRelFile = New System.Windows.Forms.CheckBox()
            Me.btnLogDirSeach = New System.Windows.Forms.Button()
            Me.txtLogDirPath = New System.Windows.Forms.TextBox()
            Me.lblLogDir = New System.Windows.Forms.Label()
            Me.btnRelDirSeach = New System.Windows.Forms.Button()
            Me.txtRelationDirPath = New System.Windows.Forms.TextBox()
            Me.lblRelationDir = New System.Windows.Forms.Label()
            Me.lblFirsttxt = New System.Windows.Forms.Label()
            Me.tabPage1 = New System.Windows.Forms.TabPage()
            Me.pnlTimeOut = New System.Windows.Forms.Panel()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.txtTimeOut = New System.Windows.Forms.TextBox()
            Me.Label36 = New System.Windows.Forms.Label()
            Me.Label35 = New System.Windows.Forms.Label()
            Me.lblNetworklib = New System.Windows.Forms.Label()
            Me.Label30 = New System.Windows.Forms.Label()
            Me.Label13 = New System.Windows.Forms.Label()
            Me.grp10ConnectInfo = New System.Windows.Forms.GroupBox()
            Me.grpV10Authent = New System.Windows.Forms.GroupBox()
            Me.optV10Authent2 = New System.Windows.Forms.RadioButton()
            Me.optV10Authent1 = New System.Windows.Forms.RadioButton()
            Me.txtV10Pass = New System.Windows.Forms.TextBox()
            Me.txtV10User = New System.Windows.Forms.TextBox()
            Me.txtV10Networklib = New System.Windows.Forms.TextBox()
            Me.txtV10Catalog = New System.Windows.Forms.TextBox()
            Me.txtV10Server = New System.Windows.Forms.TextBox()
            Me.lblV10 = New System.Windows.Forms.Label()
            Me.grpV7ConnectInfo = New System.Windows.Forms.GroupBox()
            Me.grpV7Authent = New System.Windows.Forms.GroupBox()
            Me.optV7Authent2 = New System.Windows.Forms.RadioButton()
            Me.optV7Authent1 = New System.Windows.Forms.RadioButton()
            Me.txtV7Pass = New System.Windows.Forms.TextBox()
            Me.txtV7User = New System.Windows.Forms.TextBox()
            Me.txtV7Networklib = New System.Windows.Forms.TextBox()
            Me.txtV7Catalog = New System.Windows.Forms.TextBox()
            Me.txtV7Server = New System.Windows.Forms.TextBox()
            Me.lblV7 = New System.Windows.Forms.Label()
            Me.Label57 = New System.Windows.Forms.Label()
            Me.lblDBTxt = New System.Windows.Forms.Label()
            Me.btnConnectTest = New System.Windows.Forms.Button()
            Me.tabPage2 = New System.Windows.Forms.TabPage()
            Me.Label39 = New System.Windows.Forms.Label()
            Me.grpKagi = New System.Windows.Forms.GroupBox()
            Me.optKyKagi = New System.Windows.Forms.RadioButton()
            Me.optHyKagi = New System.Windows.Forms.RadioButton()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.Label34 = New System.Windows.Forms.Label()
            Me.Label38 = New System.Windows.Forms.Label()
            Me.chkBkrui = New System.Windows.Forms.CheckBox()
            Me.chkTosiYoto = New System.Windows.Forms.CheckBox()
            Me.chkKagi = New System.Windows.Forms.CheckBox()
            Me.chkHyrui = New System.Windows.Forms.CheckBox()
            Me.chkSetubi = New System.Windows.Forms.CheckBox()
            Me.chkKyrui = New System.Windows.Forms.CheckBox()
            Me.chkKozasyu = New System.Windows.Forms.CheckBox()
            Me.chkNkinkbn = New System.Windows.Forms.CheckBox()
            Me.chkKozo = New System.Windows.Forms.CheckBox()
            Me.chkFBFmt = New System.Windows.Forms.CheckBox()
            Me.chkJisyakoza = New System.Windows.Forms.CheckBox()
            Me.chkToritaiyo = New System.Windows.Forms.CheckBox()
            Me.chkNkinkomk = New System.Windows.Forms.CheckBox()
            Me.chkGazo = New System.Windows.Forms.CheckBox()
            Me.tabPage3 = New System.Windows.Forms.TabPage()
            Me.lblHidden4 = New System.Windows.Forms.Label()
            Me.Label40 = New System.Windows.Forms.Label()
            Me.Label37 = New System.Windows.Forms.Label()
            Me.lblCaution = New System.Windows.Forms.Label()
            Me.tabCtrlRelwork = New System.Windows.Forms.TabControl()
            Me.tabPage10 = New System.Windows.Forms.TabPage()
            Me.Label6 = New System.Windows.Forms.Label()
            Me.dgvBkrui = New System.Windows.Forms.DataGridView()
            Me.ColBkrui0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColBkrui1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColBkrui2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage11 = New System.Windows.Forms.TabPage()
            Me.Label7 = New System.Windows.Forms.Label()
            Me.dgvHyrui = New System.Windows.Forms.DataGridView()
            Me.ColHyrui0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColHyrui1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColHyrui2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage12 = New System.Windows.Forms.TabPage()
            Me.Label8 = New System.Windows.Forms.Label()
            Me.dgvNkinKbn = New System.Windows.Forms.DataGridView()
            Me.ColNkinKbn0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinKbn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinKbn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinKbn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinKbn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage13 = New System.Windows.Forms.TabPage()
            Me.Label9 = New System.Windows.Forms.Label()
            Me.dgvToritaiyo = New System.Windows.Forms.DataGridView()
            Me.ColToritaiyo0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColToritaiyo1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColToritaiyo2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage14 = New System.Windows.Forms.TabPage()
            Me.Panel1 = New System.Windows.Forms.Panel()
            Me.Label44 = New System.Windows.Forms.Label()
            Me.Label43 = New System.Windows.Forms.Label()
            Me.Label31 = New System.Windows.Forms.Label()
            Me.Label42 = New System.Windows.Forms.Label()
            Me.Label17 = New System.Windows.Forms.Label()
            Me.Label18 = New System.Windows.Forms.Label()
            Me.btnSetNkinnoBulk = New System.Windows.Forms.Button()
            Me.txtNkinnoSta = New System.Windows.Forms.TextBox()
            Me.chkDuplicate = New System.Windows.Forms.CheckBox()
            Me.Label10 = New System.Windows.Forms.Label()
            Me.dgvNkinkomk = New System.Windows.Forms.DataGridView()
            Me.ColNkinkomk0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColNkinkomk8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage15 = New System.Windows.Forms.TabPage()
            Me.Label12 = New System.Windows.Forms.Label()
            Me.dgvKozo = New System.Windows.Forms.DataGridView()
            Me.ColKozo0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColKozo1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColKozo2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage16 = New System.Windows.Forms.TabPage()
            Me.Label14 = New System.Windows.Forms.Label()
            Me.dgvKozasyu = New System.Windows.Forms.DataGridView()
            Me.ColKozasyu0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColKozasyu1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColKozasyu2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage17 = New System.Windows.Forms.TabPage()
            Me.Label15 = New System.Windows.Forms.Label()
            Me.dgvSetubi = New System.Windows.Forms.DataGridView()
            Me.ColSetubi0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColSetubi1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColSetubi2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColSetubi3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColSetubi4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColSetubi5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColSetubi6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColSetubi7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage18 = New System.Windows.Forms.TabPage()
            Me.lblLine2 = New System.Windows.Forms.Label()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.dgvKagi = New System.Windows.Forms.DataGridView()
            Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage19 = New System.Windows.Forms.TabPage()
            Me.Label45 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.dgvJisyakoza = New System.Windows.Forms.DataGridView()
            Me.ColJisyakoza0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColJisyakoza17 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage20 = New System.Windows.Forms.TabPage()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.dgvGazo = New System.Windows.Forms.DataGridView()
            Me.ColGazo0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColGazo1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColGazo2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColGazo3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColGazo4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColGazo5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage21 = New System.Windows.Forms.TabPage()
            Me.Label46 = New System.Windows.Forms.Label()
            Me.Label32 = New System.Windows.Forms.Label()
            Me.dgvFBInfo = New System.Windows.Forms.DataGridView()
            Me.ColFBInfo0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColFBInfo13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage22 = New System.Windows.Forms.TabPage()
            Me.Label29 = New System.Windows.Forms.Label()
            Me.dgvKyrui = New System.Windows.Forms.DataGridView()
            Me.ColKyrui0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColKyrui1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColKyrui2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage23 = New System.Windows.Forms.TabPage()
            Me.Label33 = New System.Windows.Forms.Label()
            Me.dgvTosiYoto = New System.Windows.Forms.DataGridView()
            Me.ColTosiYoto0 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColTosiYoto1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColTosiYoto2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.ColTosiYoto3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.tabPage4 = New System.Windows.Forms.TabPage()
            Me.lblFinaltxt = New System.Windows.Forms.Label()
            Me.lblRelItemRead = New System.Windows.Forms.Label()
            Me.pgbRelItemRead = New System.Windows.Forms.ProgressBar()
            Me.btnBack = New System.Windows.Forms.Button()
            Me.btnNext = New System.Windows.Forms.Button()
            Me.btnEnd = New System.Windows.Forms.Button()
            Me.DataGridView1 = New System.Windows.Forms.DataGridView()
            Me.Label16 = New System.Windows.Forms.Label()
            Me.Label19 = New System.Windows.Forms.Label()
            Me.Label20 = New System.Windows.Forms.Label()
            Me.Label21 = New System.Windows.Forms.Label()
            Me.Label22 = New System.Windows.Forms.Label()
            Me.Label23 = New System.Windows.Forms.Label()
            Me.Label24 = New System.Windows.Forms.Label()
            Me.Label25 = New System.Windows.Forms.Label()
            Me.Label26 = New System.Windows.Forms.Label()
            Me.Label27 = New System.Windows.Forms.Label()
            Me.Label28 = New System.Windows.Forms.Label()
            Me.lbltest = New System.Windows.Forms.Label()
            Me.lblHidden3 = New System.Windows.Forms.Label()
            Me.pnlHaisyoku = New System.Windows.Forms.Panel()
            Me.pnlRelSelectPanel = New System.Windows.Forms.Panel()
            Me.lblCntAllTosiYoto = New System.Windows.Forms.Label()
            Me.lblCntSumiTosiYoto = New System.Windows.Forms.Label()
            Me.lblCntAllKyBunrui = New System.Windows.Forms.Label()
            Me.lblTitleRelTotalCnt = New System.Windows.Forms.Label()
            Me.lblCntAllFbFmt = New System.Windows.Forms.Label()
            Me.lblCntSumiKyBunrui = New System.Windows.Forms.Label()
            Me.lblCntAllNkKbn = New System.Windows.Forms.Label()
            Me.lblCntAllGazo = New System.Windows.Forms.Label()
            Me.lblTitleRelCnt = New System.Windows.Forms.Label()
            Me.lblCntAllTaiyo = New System.Windows.Forms.Label()
            Me.lblCntSumiNkKbn = New System.Windows.Forms.Label()
            Me.lblCntAllJisyaKoza = New System.Windows.Forms.Label()
            Me.lblCntSumiFbFmt = New System.Windows.Forms.Label()
            Me.lblCntAllNkin = New System.Windows.Forms.Label()
            Me.lblTitleRelKomk = New System.Windows.Forms.Label()
            Me.lblCntSumiTaiyo = New System.Windows.Forms.Label()
            Me.lblCntAllKozo = New System.Windows.Forms.Label()
            Me.lblRelSelectTosiYoto = New System.Windows.Forms.Label()
            Me.lblCntAllSetubi = New System.Windows.Forms.Label()
            Me.lblSumiZanGazo = New System.Windows.Forms.Label()
            Me.lblCntAllKozaSyubetu = New System.Windows.Forms.Label()
            Me.lblRelSelectJisyaKoza = New System.Windows.Forms.Label()
            Me.lblCntSumiNkin = New System.Windows.Forms.Label()
            Me.lblRelSelectKyBunrui = New System.Windows.Forms.Label()
            Me.lblCntSumiJisyaKoza = New System.Windows.Forms.Label()
            Me.lblRelSelectGazo = New System.Windows.Forms.Label()
            Me.lblCntSumiKozo = New System.Windows.Forms.Label()
            Me.lblRelSelectFBFmt = New System.Windows.Forms.Label()
            Me.lblCntSumiKozaSyubetu = New System.Windows.Forms.Label()
            Me.lblRelSelectNkin = New System.Windows.Forms.Label()
            Me.lblCntSumiSetubi = New System.Windows.Forms.Label()
            Me.lblRelSelectHySetubi = New System.Windows.Forms.Label()
            Me.lblRelSelectKozo = New System.Windows.Forms.Label()
            Me.lblRelSelectKozaSyubetu = New System.Windows.Forms.Label()
            Me.lblCntAllBkBunrui = New System.Windows.Forms.Label()
            Me.lblTitleRelSelect = New System.Windows.Forms.Label()
            Me.lblRelSelectBkBunrui = New System.Windows.Forms.Label()
            Me.lblCntSumiBkBunrui = New System.Windows.Forms.Label()
            Me.lblRelSelectTaiyo = New System.Windows.Forms.Label()
            Me.lblRelSelectHyBunrui = New System.Windows.Forms.Label()
            Me.lblRelSelectNkKbn = New System.Windows.Forms.Label()
            Me.lblCntTosiYotoBack = New System.Windows.Forms.Label()
            Me.lblCntKyBunruiBack = New System.Windows.Forms.Label()
            Me.lblCntSumiHyBunrui = New System.Windows.Forms.Label()
            Me.lblCntFbFmtBack = New System.Windows.Forms.Label()
            Me.lblCntGazoBack = New System.Windows.Forms.Label()
            Me.lblCntAllHyBunrui = New System.Windows.Forms.Label()
            Me.lblCntNkKbnBack = New System.Windows.Forms.Label()
            Me.lblCntTaiyoBack = New System.Windows.Forms.Label()
            Me.lblCntNkinBack = New System.Windows.Forms.Label()
            Me.lblCntJisyaKozaBack = New System.Windows.Forms.Label()
            Me.lblCntSetubiBack = New System.Windows.Forms.Label()
            Me.lblCntKozaSyubetuBack = New System.Windows.Forms.Label()
            Me.lblCntKozoBack = New System.Windows.Forms.Label()
            Me.lblCntHyBunruiBack = New System.Windows.Forms.Label()
            Me.lblCntBkBunruiBack = New System.Windows.Forms.Label()
            Me.tabCtrlRelMain.SuspendLayout()
            Me.tabPage0.SuspendLayout()
            Me.grpKagicntgrp.SuspendLayout()
            Me.grpDev.SuspendLayout()
            Me.tabPage1.SuspendLayout()
            Me.pnlTimeOut.SuspendLayout()
            Me.grp10ConnectInfo.SuspendLayout()
            Me.grpV10Authent.SuspendLayout()
            Me.grpV7ConnectInfo.SuspendLayout()
            Me.grpV7Authent.SuspendLayout()
            Me.tabPage2.SuspendLayout()
            Me.grpKagi.SuspendLayout()
            Me.tabPage3.SuspendLayout()
            Me.tabCtrlRelwork.SuspendLayout()
            Me.tabPage10.SuspendLayout()
            CType(Me.dgvBkrui, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage11.SuspendLayout()
            CType(Me.dgvHyrui, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage12.SuspendLayout()
            CType(Me.dgvNkinKbn, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage13.SuspendLayout()
            CType(Me.dgvToritaiyo, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage14.SuspendLayout()
            Me.Panel1.SuspendLayout()
            CType(Me.dgvNkinkomk, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage15.SuspendLayout()
            CType(Me.dgvKozo, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage16.SuspendLayout()
            CType(Me.dgvKozasyu, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage17.SuspendLayout()
            CType(Me.dgvSetubi, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage18.SuspendLayout()
            CType(Me.dgvKagi, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage19.SuspendLayout()
            CType(Me.dgvJisyakoza, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage20.SuspendLayout()
            CType(Me.dgvGazo, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage21.SuspendLayout()
            CType(Me.dgvFBInfo, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage22.SuspendLayout()
            CType(Me.dgvKyrui, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage23.SuspendLayout()
            CType(Me.dgvTosiYoto, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.tabPage4.SuspendLayout()
            CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.pnlHaisyoku.SuspendLayout()
            Me.pnlRelSelectPanel.SuspendLayout()
            Me.SuspendLayout()
            '
            'tabCtrlRelMain
            '
            Me.tabCtrlRelMain.Controls.Add(Me.tabPage0)
            Me.tabCtrlRelMain.Controls.Add(Me.tabPage1)
            Me.tabCtrlRelMain.Controls.Add(Me.tabPage2)
            Me.tabCtrlRelMain.Controls.Add(Me.tabPage3)
            Me.tabCtrlRelMain.Controls.Add(Me.tabPage4)
            Me.tabCtrlRelMain.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
            Me.tabCtrlRelMain.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.tabCtrlRelMain.Location = New System.Drawing.Point(30, 40)
            Me.tabCtrlRelMain.Name = "tabCtrlRelMain"
            Me.tabCtrlRelMain.SelectedIndex = 0
            Me.tabCtrlRelMain.Size = New System.Drawing.Size(920, 580)
            Me.tabCtrlRelMain.TabIndex = 64
            '
            'tabPage0
            '
            Me.tabPage0.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage0.Controls.Add(Me.grpKagicntgrp)
            Me.tabPage0.Controls.Add(Me.Label47)
            Me.tabPage0.Controls.Add(Me.lblLine0)
            Me.tabPage0.Controls.Add(Me.btnMidDirSeach)
            Me.tabPage0.Controls.Add(Me.txtMidDirPath)
            Me.tabPage0.Controls.Add(Me.lblMidDir)
            Me.tabPage0.Controls.Add(Me.grpDev)
            Me.tabPage0.Controls.Add(Me.btnLogDirSeach)
            Me.tabPage0.Controls.Add(Me.txtLogDirPath)
            Me.tabPage0.Controls.Add(Me.lblLogDir)
            Me.tabPage0.Controls.Add(Me.btnRelDirSeach)
            Me.tabPage0.Controls.Add(Me.txtRelationDirPath)
            Me.tabPage0.Controls.Add(Me.lblRelationDir)
            Me.tabPage0.Controls.Add(Me.lblFirsttxt)
            Me.tabPage0.Location = New System.Drawing.Point(4, 27)
            Me.tabPage0.Name = "tabPage0"
            Me.tabPage0.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage0.Size = New System.Drawing.Size(912, 549)
            Me.tabPage0.TabIndex = 0
            Me.tabPage0.Text = "初期設定"
            '
            'grpKagicntgrp
            '
            Me.grpKagicntgrp.Controls.Add(Me.lblRelSelectKyoyoKagi)
            Me.grpKagicntgrp.Controls.Add(Me.lblCntKyoyoKagiBack)
            Me.grpKagicntgrp.Controls.Add(Me.lblCntSumiKyoyoKagi)
            Me.grpKagicntgrp.Controls.Add(Me.lblCntAllKyoyoKagi)
            Me.grpKagicntgrp.Location = New System.Drawing.Point(488, 272)
            Me.grpKagicntgrp.Name = "grpKagicntgrp"
            Me.grpKagicntgrp.Size = New System.Drawing.Size(208, 56)
            Me.grpKagicntgrp.TabIndex = 231
            Me.grpKagicntgrp.TabStop = False
            Me.grpKagicntgrp.Text = "鍵非表示用"
            '
            'lblRelSelectKyoyoKagi
            '
            Me.lblRelSelectKyoyoKagi.BackColor = System.Drawing.Color.White
            Me.lblRelSelectKyoyoKagi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectKyoyoKagi.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectKyoyoKagi.Location = New System.Drawing.Point(24, 24)
            Me.lblRelSelectKyoyoKagi.Name = "lblRelSelectKyoyoKagi"
            Me.lblRelSelectKyoyoKagi.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectKyoyoKagi.TabIndex = 146
            Me.lblRelSelectKyoyoKagi.Text = "共用鍵設定"
            Me.lblRelSelectKyoyoKagi.UseCompatibleTextRendering = True
            '
            'lblCntKyoyoKagiBack
            '
            Me.lblCntKyoyoKagiBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntKyoyoKagiBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntKyoyoKagiBack.Location = New System.Drawing.Point(146, 27)
            Me.lblCntKyoyoKagiBack.Name = "lblCntKyoyoKagiBack"
            Me.lblCntKyoyoKagiBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntKyoyoKagiBack.TabIndex = 230
            Me.lblCntKyoyoKagiBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntKyoyoKagiBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSumiKyoyoKagi
            '
            Me.lblCntSumiKyoyoKagi.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiKyoyoKagi.Location = New System.Drawing.Point(165, 28)
            Me.lblCntSumiKyoyoKagi.Name = "lblCntSumiKyoyoKagi"
            Me.lblCntSumiKyoyoKagi.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiKyoyoKagi.TabIndex = 192
            Me.lblCntSumiKyoyoKagi.Text = "9,999"
            Me.lblCntSumiKyoyoKagi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllKyoyoKagi
            '
            Me.lblCntAllKyoyoKagi.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllKyoyoKagi.Location = New System.Drawing.Point(224, 28)
            Me.lblCntAllKyoyoKagi.Name = "lblCntAllKyoyoKagi"
            Me.lblCntAllKyoyoKagi.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllKyoyoKagi.TabIndex = 193
            Me.lblCntAllKyoyoKagi.Text = "9,999"
            Me.lblCntAllKyoyoKagi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'Label47
            '
            Me.Label47.Font = New System.Drawing.Font("メイリオ", 9.75!)
            Me.Label47.ForeColor = System.Drawing.Color.Navy
            Me.Label47.Location = New System.Drawing.Point(35, 49)
            Me.Label47.Name = "Label47"
            Me.Label47.Size = New System.Drawing.Size(716, 41)
            Me.Label47.TabIndex = 145
            Me.Label47.Text = "紐付作業を行います。各種ファイルの出力先を指定して下さい。"
            Me.Label47.UseCompatibleTextRendering = True
            '
            'lblLine0
            '
            Me.lblLine0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblLine0.Location = New System.Drawing.Point(-4, -3)
            Me.lblLine0.Name = "lblLine0"
            Me.lblLine0.Size = New System.Drawing.Size(916, 2)
            Me.lblLine0.TabIndex = 144
            '
            'btnMidDirSeach
            '
            Me.btnMidDirSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.btnMidDirSeach.Location = New System.Drawing.Point(596, 117)
            Me.btnMidDirSeach.Name = "btnMidDirSeach"
            Me.btnMidDirSeach.Size = New System.Drawing.Size(32, 24)
            Me.btnMidDirSeach.TabIndex = 120
            Me.btnMidDirSeach.Text = "..."
            Me.btnMidDirSeach.UseVisualStyleBackColor = True
            '
            'txtMidDirPath
            '
            Me.txtMidDirPath.Location = New System.Drawing.Point(245, 117)
            Me.txtMidDirPath.Name = "txtMidDirPath"
            Me.txtMidDirPath.Size = New System.Drawing.Size(345, 25)
            Me.txtMidDirPath.TabIndex = 119
            '
            'lblMidDir
            '
            Me.lblMidDir.Location = New System.Drawing.Point(74, 120)
            Me.lblMidDir.Name = "lblMidDir"
            Me.lblMidDir.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.lblMidDir.Size = New System.Drawing.Size(165, 20)
            Me.lblMidDir.TabIndex = 118
            Me.lblMidDir.Text = "読込元中間ファイル格納先"
            Me.lblMidDir.UseCompatibleTextRendering = True
            '
            'grpDev
            '
            Me.grpDev.Controls.Add(Me.chkLogTblDrop)
            Me.grpDev.Controls.Add(Me.Label11)
            Me.grpDev.Controls.Add(Me.chkTempTable)
            Me.grpDev.Controls.Add(Me.chkRelFile)
            Me.grpDev.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.grpDev.Location = New System.Drawing.Point(76, 260)
            Me.grpDev.Name = "grpDev"
            Me.grpDev.Size = New System.Drawing.Size(379, 170)
            Me.grpDev.TabIndex = 117
            Me.grpDev.TabStop = False
            Me.grpDev.Text = "開発用"
            Me.grpDev.Visible = False
            '
            'chkLogTblDrop
            '
            Me.chkLogTblDrop.AutoSize = True
            Me.chkLogTblDrop.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.chkLogTblDrop.Location = New System.Drawing.Point(24, 128)
            Me.chkLogTblDrop.Name = "chkLogTblDrop"
            Me.chkLogTblDrop.Size = New System.Drawing.Size(159, 22)
            Me.chkLogTblDrop.TabIndex = 119
            Me.chkLogTblDrop.Text = "ログテーブルを削除する"
            Me.chkLogTblDrop.UseVisualStyleBackColor = True
            '
            'Label11
            '
            Me.Label11.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label11.Location = New System.Drawing.Point(41, 57)
            Me.Label11.Name = "Label11"
            Me.Label11.Size = New System.Drawing.Size(310, 40)
            Me.Label11.TabIndex = 118
            Me.Label11.Text = "紐付け設定ファイルを既に作成している、または用意している場合はチェックをOFFにして下さい。"
            Me.Label11.UseCompatibleTextRendering = True
            '
            'chkTempTable
            '
            Me.chkTempTable.AutoSize = True
            Me.chkTempTable.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.chkTempTable.Location = New System.Drawing.Point(24, 100)
            Me.chkTempTable.Name = "chkTempTable"
            Me.chkTempTable.Size = New System.Drawing.Size(111, 22)
            Me.chkTempTable.TabIndex = 117
            Me.chkTempTable.Text = "仮テーブル登録"
            Me.chkTempTable.UseVisualStyleBackColor = True
            '
            'chkRelFile
            '
            Me.chkRelFile.AutoSize = True
            Me.chkRelFile.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.chkRelFile.Location = New System.Drawing.Point(24, 32)
            Me.chkRelFile.Name = "chkRelFile"
            Me.chkRelFile.Size = New System.Drawing.Size(159, 22)
            Me.chkRelFile.TabIndex = 115
            Me.chkRelFile.Text = "紐付け設定ファイル作成"
            Me.chkRelFile.UseVisualStyleBackColor = True
            '
            'btnLogDirSeach
            '
            Me.btnLogDirSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.btnLogDirSeach.Location = New System.Drawing.Point(596, 178)
            Me.btnLogDirSeach.Name = "btnLogDirSeach"
            Me.btnLogDirSeach.Size = New System.Drawing.Size(32, 24)
            Me.btnLogDirSeach.TabIndex = 113
            Me.btnLogDirSeach.Text = "..."
            Me.btnLogDirSeach.UseVisualStyleBackColor = True
            '
            'txtLogDirPath
            '
            Me.txtLogDirPath.Location = New System.Drawing.Point(245, 178)
            Me.txtLogDirPath.Name = "txtLogDirPath"
            Me.txtLogDirPath.Size = New System.Drawing.Size(345, 25)
            Me.txtLogDirPath.TabIndex = 112
            '
            'lblLogDir
            '
            Me.lblLogDir.Location = New System.Drawing.Point(74, 181)
            Me.lblLogDir.Name = "lblLogDir"
            Me.lblLogDir.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.lblLogDir.Size = New System.Drawing.Size(165, 20)
            Me.lblLogDir.TabIndex = 111
            Me.lblLogDir.Text = "ログファイル格納先"
            Me.lblLogDir.UseCompatibleTextRendering = True
            '
            'btnRelDirSeach
            '
            Me.btnRelDirSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.btnRelDirSeach.Location = New System.Drawing.Point(596, 146)
            Me.btnRelDirSeach.Name = "btnRelDirSeach"
            Me.btnRelDirSeach.Size = New System.Drawing.Size(32, 24)
            Me.btnRelDirSeach.TabIndex = 110
            Me.btnRelDirSeach.Text = "..."
            Me.btnRelDirSeach.UseVisualStyleBackColor = True
            '
            'txtRelationDirPath
            '
            Me.txtRelationDirPath.Location = New System.Drawing.Point(245, 146)
            Me.txtRelationDirPath.Name = "txtRelationDirPath"
            Me.txtRelationDirPath.Size = New System.Drawing.Size(345, 25)
            Me.txtRelationDirPath.TabIndex = 109
            '
            'lblRelationDir
            '
            Me.lblRelationDir.Location = New System.Drawing.Point(76, 149)
            Me.lblRelationDir.Name = "lblRelationDir"
            Me.lblRelationDir.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.lblRelationDir.Size = New System.Drawing.Size(165, 20)
            Me.lblRelationDir.TabIndex = 108
            Me.lblRelationDir.Text = "紐付ファイル格納先"
            Me.lblRelationDir.UseCompatibleTextRendering = True
            '
            'lblFirsttxt
            '
            Me.lblFirsttxt.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Bold)
            Me.lblFirsttxt.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
            Me.lblFirsttxt.Location = New System.Drawing.Point(30, 20)
            Me.lblFirsttxt.Name = "lblFirsttxt"
            Me.lblFirsttxt.Size = New System.Drawing.Size(716, 41)
            Me.lblFirsttxt.TabIndex = 93
            Me.lblFirsttxt.Text = "紐付作業を行います。各種ファイルの出力先を指定して下さい。"
            Me.lblFirsttxt.UseCompatibleTextRendering = True
            '
            'tabPage1
            '
            Me.tabPage1.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage1.Controls.Add(Me.pnlTimeOut)
            Me.tabPage1.Controls.Add(Me.Label36)
            Me.tabPage1.Controls.Add(Me.Label35)
            Me.tabPage1.Controls.Add(Me.lblNetworklib)
            Me.tabPage1.Controls.Add(Me.Label30)
            Me.tabPage1.Controls.Add(Me.Label13)
            Me.tabPage1.Controls.Add(Me.grp10ConnectInfo)
            Me.tabPage1.Controls.Add(Me.grpV7ConnectInfo)
            Me.tabPage1.Controls.Add(Me.Label57)
            Me.tabPage1.Controls.Add(Me.lblDBTxt)
            Me.tabPage1.Controls.Add(Me.btnConnectTest)
            Me.tabPage1.Location = New System.Drawing.Point(4, 27)
            Me.tabPage1.Name = "tabPage1"
            Me.tabPage1.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage1.Size = New System.Drawing.Size(912, 549)
            Me.tabPage1.TabIndex = 1
            Me.tabPage1.Text = "DB接続"
            '
            'pnlTimeOut
            '
            Me.pnlTimeOut.Controls.Add(Me.Label4)
            Me.pnlTimeOut.Controls.Add(Me.txtTimeOut)
            Me.pnlTimeOut.Location = New System.Drawing.Point(749, 20)
            Me.pnlTimeOut.Name = "pnlTimeOut"
            Me.pnlTimeOut.Size = New System.Drawing.Size(140, 73)
            Me.pnlTimeOut.TabIndex = 116
            '
            'Label4
            '
            Me.Label4.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label4.Location = New System.Drawing.Point(12, 7)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(114, 19)
            Me.Label4.TabIndex = 115
            Me.Label4.Text = "タイムアウト値"
            Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label4.UseCompatibleTextRendering = True
            '
            'txtTimeOut
            '
            Me.txtTimeOut.Location = New System.Drawing.Point(54, 32)
            Me.txtTimeOut.Multiline = True
            Me.txtTimeOut.Name = "txtTimeOut"
            Me.txtTimeOut.Size = New System.Drawing.Size(59, 27)
            Me.txtTimeOut.TabIndex = 108
            Me.txtTimeOut.Text = "30"
            Me.txtTimeOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            '
            'Label36
            '
            Me.Label36.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label36.Location = New System.Drawing.Point(34, 374)
            Me.Label36.Name = "Label36"
            Me.Label36.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.Label36.Size = New System.Drawing.Size(150, 20)
            Me.Label36.TabIndex = 114
            Me.Label36.Text = "パスワード"
            Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label36.UseCompatibleTextRendering = True
            '
            'Label35
            '
            Me.Label35.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label35.Location = New System.Drawing.Point(34, 334)
            Me.Label35.Name = "Label35"
            Me.Label35.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.Label35.Size = New System.Drawing.Size(150, 20)
            Me.Label35.TabIndex = 113
            Me.Label35.Text = "ログイン"
            Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label35.UseCompatibleTextRendering = True
            '
            'lblNetworklib
            '
            Me.lblNetworklib.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblNetworklib.Location = New System.Drawing.Point(34, 417)
            Me.lblNetworklib.Name = "lblNetworklib"
            Me.lblNetworklib.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.lblNetworklib.Size = New System.Drawing.Size(150, 20)
            Me.lblNetworklib.TabIndex = 112
            Me.lblNetworklib.Text = "ネットワークライブラリ"
            Me.lblNetworklib.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.lblNetworklib.UseCompatibleTextRendering = True
            '
            'Label30
            '
            Me.Label30.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label30.Location = New System.Drawing.Point(34, 254)
            Me.Label30.Name = "Label30"
            Me.Label30.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.Label30.Size = New System.Drawing.Size(150, 20)
            Me.Label30.TabIndex = 111
            Me.Label30.Text = "サーバー名"
            Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label30.UseCompatibleTextRendering = True
            '
            'Label13
            '
            Me.Label13.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label13.Location = New System.Drawing.Point(34, 294)
            Me.Label13.Name = "Label13"
            Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.Yes
            Me.Label13.Size = New System.Drawing.Size(150, 20)
            Me.Label13.TabIndex = 110
            Me.Label13.Text = "カタログ名"
            Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label13.UseCompatibleTextRendering = True
            '
            'grp10ConnectInfo
            '
            Me.grp10ConnectInfo.Controls.Add(Me.grpV10Authent)
            Me.grp10ConnectInfo.Controls.Add(Me.txtV10Pass)
            Me.grp10ConnectInfo.Controls.Add(Me.txtV10User)
            Me.grp10ConnectInfo.Controls.Add(Me.txtV10Networklib)
            Me.grp10ConnectInfo.Controls.Add(Me.txtV10Catalog)
            Me.grp10ConnectInfo.Controls.Add(Me.txtV10Server)
            Me.grp10ConnectInfo.Controls.Add(Me.lblV10)
            Me.grp10ConnectInfo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.grp10ConnectInfo.Location = New System.Drawing.Point(529, 132)
            Me.grp10ConnectInfo.Name = "grp10ConnectInfo"
            Me.grp10ConnectInfo.Size = New System.Drawing.Size(300, 330)
            Me.grp10ConnectInfo.TabIndex = 107
            Me.grp10ConnectInfo.TabStop = False
            Me.grp10ConnectInfo.Text = "【移行先接続情報】"
            '
            'grpV10Authent
            '
            Me.grpV10Authent.Controls.Add(Me.optV10Authent2)
            Me.grpV10Authent.Controls.Add(Me.optV10Authent1)
            Me.grpV10Authent.Location = New System.Drawing.Point(6, 56)
            Me.grpV10Authent.Name = "grpV10Authent"
            Me.grpV10Authent.Size = New System.Drawing.Size(288, 50)
            Me.grpV10Authent.TabIndex = 78
            Me.grpV10Authent.TabStop = False
            '
            'optV10Authent2
            '
            Me.optV10Authent2.AutoSize = True
            Me.optV10Authent2.Location = New System.Drawing.Point(160, 18)
            Me.optV10Authent2.Name = "optV10Authent2"
            Me.optV10Authent2.Size = New System.Drawing.Size(114, 24)
            Me.optV10Authent2.TabIndex = 55
            Me.optV10Authent2.Text = "Windows 認証"
            Me.optV10Authent2.UseVisualStyleBackColor = True
            '
            'optV10Authent1
            '
            Me.optV10Authent1.AutoSize = True
            Me.optV10Authent1.Checked = True
            Me.optV10Authent1.Location = New System.Drawing.Point(12, 18)
            Me.optV10Authent1.Name = "optV10Authent1"
            Me.optV10Authent1.Size = New System.Drawing.Size(123, 24)
            Me.optV10Authent1.TabIndex = 54
            Me.optV10Authent1.TabStop = True
            Me.optV10Authent1.Text = "SQLServer 認証"
            Me.optV10Authent1.UseVisualStyleBackColor = True
            '
            'txtV10Pass
            '
            Me.txtV10Pass.Location = New System.Drawing.Point(10, 238)
            Me.txtV10Pass.Name = "txtV10Pass"
            Me.txtV10Pass.Size = New System.Drawing.Size(280, 27)
            Me.txtV10Pass.TabIndex = 77
            Me.txtV10Pass.Text = "***************"
            '
            'txtV10User
            '
            Me.txtV10User.Location = New System.Drawing.Point(10, 201)
            Me.txtV10User.Name = "txtV10User"
            Me.txtV10User.Size = New System.Drawing.Size(280, 27)
            Me.txtV10User.TabIndex = 76
            Me.txtV10User.Text = "sa"
            '
            'txtV10Networklib
            '
            Me.txtV10Networklib.Location = New System.Drawing.Point(10, 280)
            Me.txtV10Networklib.Name = "txtV10Networklib"
            Me.txtV10Networklib.Size = New System.Drawing.Size(280, 27)
            Me.txtV10Networklib.TabIndex = 75
            Me.txtV10Networklib.Text = "共有メモリ"
            '
            'txtV10Catalog
            '
            Me.txtV10Catalog.Location = New System.Drawing.Point(10, 160)
            Me.txtV10Catalog.Name = "txtV10Catalog"
            Me.txtV10Catalog.Size = New System.Drawing.Size(280, 27)
            Me.txtV10Catalog.TabIndex = 74
            Me.txtV10Catalog.Text = "fk8db"
            '
            'txtV10Server
            '
            Me.txtV10Server.Location = New System.Drawing.Point(10, 120)
            Me.txtV10Server.Name = "txtV10Server"
            Me.txtV10Server.Size = New System.Drawing.Size(280, 27)
            Me.txtV10Server.TabIndex = 73
            Me.txtV10Server.Text = "PC-NJC\SQL2012"
            '
            'lblV10
            '
            Me.lblV10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblV10.Image = CType(resources.GetObject("lblV10.Image"), System.Drawing.Image)
            Me.lblV10.Location = New System.Drawing.Point(6, 30)
            Me.lblV10.Name = "lblV10"
            Me.lblV10.Size = New System.Drawing.Size(288, 24)
            Me.lblV10.TabIndex = 72
            Me.lblV10.Text = "賃貸革命10"
            Me.lblV10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'grpV7ConnectInfo
            '
            Me.grpV7ConnectInfo.Controls.Add(Me.grpV7Authent)
            Me.grpV7ConnectInfo.Controls.Add(Me.txtV7Pass)
            Me.grpV7ConnectInfo.Controls.Add(Me.txtV7User)
            Me.grpV7ConnectInfo.Controls.Add(Me.txtV7Networklib)
            Me.grpV7ConnectInfo.Controls.Add(Me.txtV7Catalog)
            Me.grpV7ConnectInfo.Controls.Add(Me.txtV7Server)
            Me.grpV7ConnectInfo.Controls.Add(Me.lblV7)
            Me.grpV7ConnectInfo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.grpV7ConnectInfo.Location = New System.Drawing.Point(190, 132)
            Me.grpV7ConnectInfo.Name = "grpV7ConnectInfo"
            Me.grpV7ConnectInfo.Size = New System.Drawing.Size(300, 330)
            Me.grpV7ConnectInfo.TabIndex = 106
            Me.grpV7ConnectInfo.TabStop = False
            Me.grpV7ConnectInfo.Text = "【移行元接続情報】"
            '
            'grpV7Authent
            '
            Me.grpV7Authent.Controls.Add(Me.optV7Authent2)
            Me.grpV7Authent.Controls.Add(Me.optV7Authent1)
            Me.grpV7Authent.Location = New System.Drawing.Point(6, 56)
            Me.grpV7Authent.Name = "grpV7Authent"
            Me.grpV7Authent.Size = New System.Drawing.Size(288, 50)
            Me.grpV7Authent.TabIndex = 70
            Me.grpV7Authent.TabStop = False
            '
            'optV7Authent2
            '
            Me.optV7Authent2.AutoSize = True
            Me.optV7Authent2.Location = New System.Drawing.Point(160, 18)
            Me.optV7Authent2.Name = "optV7Authent2"
            Me.optV7Authent2.Size = New System.Drawing.Size(114, 24)
            Me.optV7Authent2.TabIndex = 55
            Me.optV7Authent2.Text = "Windows 認証"
            Me.optV7Authent2.UseVisualStyleBackColor = True
            '
            'optV7Authent1
            '
            Me.optV7Authent1.AutoSize = True
            Me.optV7Authent1.Checked = True
            Me.optV7Authent1.Location = New System.Drawing.Point(12, 18)
            Me.optV7Authent1.Name = "optV7Authent1"
            Me.optV7Authent1.Size = New System.Drawing.Size(123, 24)
            Me.optV7Authent1.TabIndex = 54
            Me.optV7Authent1.TabStop = True
            Me.optV7Authent1.Text = "SQLServer 認証"
            Me.optV7Authent1.UseVisualStyleBackColor = True
            '
            'txtV7Pass
            '
            Me.txtV7Pass.Location = New System.Drawing.Point(10, 240)
            Me.txtV7Pass.Name = "txtV7Pass"
            Me.txtV7Pass.Size = New System.Drawing.Size(280, 27)
            Me.txtV7Pass.TabIndex = 64
            Me.txtV7Pass.Text = "***************"
            '
            'txtV7User
            '
            Me.txtV7User.Location = New System.Drawing.Point(10, 200)
            Me.txtV7User.Name = "txtV7User"
            Me.txtV7User.Size = New System.Drawing.Size(280, 27)
            Me.txtV7User.TabIndex = 63
            Me.txtV7User.Text = "sa"
            '
            'txtV7Networklib
            '
            Me.txtV7Networklib.Location = New System.Drawing.Point(10, 280)
            Me.txtV7Networklib.Name = "txtV7Networklib"
            Me.txtV7Networklib.Size = New System.Drawing.Size(280, 27)
            Me.txtV7Networklib.TabIndex = 62
            Me.txtV7Networklib.Text = "共有メモリ"
            '
            'txtV7Catalog
            '
            Me.txtV7Catalog.Location = New System.Drawing.Point(10, 160)
            Me.txtV7Catalog.Name = "txtV7Catalog"
            Me.txtV7Catalog.Size = New System.Drawing.Size(280, 27)
            Me.txtV7Catalog.TabIndex = 61
            Me.txtV7Catalog.Text = "fk5dtsql"
            '
            'txtV7Server
            '
            Me.txtV7Server.Location = New System.Drawing.Point(10, 120)
            Me.txtV7Server.Name = "txtV7Server"
            Me.txtV7Server.Size = New System.Drawing.Size(280, 27)
            Me.txtV7Server.TabIndex = 60
            Me.txtV7Server.Text = "PC-NJC\SQL2008"
            '
            'lblV7
            '
            Me.lblV7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblV7.Image = CType(resources.GetObject("lblV7.Image"), System.Drawing.Image)
            Me.lblV7.Location = New System.Drawing.Point(10, 30)
            Me.lblV7.Name = "lblV7"
            Me.lblV7.Size = New System.Drawing.Size(284, 24)
            Me.lblV7.TabIndex = 53
            Me.lblV7.Text = "賃貸革命V7"
            Me.lblV7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'Label57
            '
            Me.Label57.BackColor = System.Drawing.SystemColors.Menu
            Me.Label57.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label57.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
            Me.Label57.Location = New System.Drawing.Point(30, 90)
            Me.Label57.Name = "Label57"
            Me.Label57.Size = New System.Drawing.Size(820, 24)
            Me.Label57.TabIndex = 105
            Me.Label57.Text = "※正常接続が確認できない場合、[次へ] に進むことはできません。"
            Me.Label57.UseCompatibleTextRendering = True
            '
            'lblDBTxt
            '
            Me.lblDBTxt.BackColor = System.Drawing.SystemColors.Menu
            Me.lblDBTxt.ForeColor = System.Drawing.Color.Navy
            Me.lblDBTxt.Location = New System.Drawing.Point(30, 20)
            Me.lblDBTxt.Name = "lblDBTxt"
            Me.lblDBTxt.Size = New System.Drawing.Size(820, 80)
            Me.lblDBTxt.TabIndex = 62
            Me.lblDBTxt.Text = "データベースの接続設定を行います。([初期値読込] を押すと接続情報の初期設定値を読み込みます)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "「賃貸革命V7」と「賃貸革命10」の接続情報を設定し、[接続確" & _
        "認] ボタンを押して下さい。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "正常接続が確認できましたら [次へ] ボタンを押して下さい。"
            Me.lblDBTxt.UseCompatibleTextRendering = True
            '
            'btnConnectTest
            '
            Me.btnConnectTest.Font = New System.Drawing.Font("メイリオ", 11.25!)
            Me.btnConnectTest.Image = CType(resources.GetObject("btnConnectTest.Image"), System.Drawing.Image)
            Me.btnConnectTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnConnectTest.Location = New System.Drawing.Point(769, 506)
            Me.btnConnectTest.Name = "btnConnectTest"
            Me.btnConnectTest.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
            Me.btnConnectTest.Size = New System.Drawing.Size(120, 30)
            Me.btnConnectTest.TabIndex = 61
            Me.btnConnectTest.Text = " 接続確認"
            Me.btnConnectTest.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
            Me.btnConnectTest.UseVisualStyleBackColor = True
            '
            'tabPage2
            '
            Me.tabPage2.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage2.Controls.Add(Me.Label39)
            Me.tabPage2.Controls.Add(Me.grpKagi)
            Me.tabPage2.Controls.Add(Me.Label1)
            Me.tabPage2.Controls.Add(Me.Label34)
            Me.tabPage2.Controls.Add(Me.Label38)
            Me.tabPage2.Controls.Add(Me.chkBkrui)
            Me.tabPage2.Controls.Add(Me.chkTosiYoto)
            Me.tabPage2.Controls.Add(Me.chkKagi)
            Me.tabPage2.Controls.Add(Me.chkHyrui)
            Me.tabPage2.Controls.Add(Me.chkSetubi)
            Me.tabPage2.Controls.Add(Me.chkKyrui)
            Me.tabPage2.Controls.Add(Me.chkKozasyu)
            Me.tabPage2.Controls.Add(Me.chkNkinkbn)
            Me.tabPage2.Controls.Add(Me.chkKozo)
            Me.tabPage2.Controls.Add(Me.chkFBFmt)
            Me.tabPage2.Controls.Add(Me.chkJisyakoza)
            Me.tabPage2.Controls.Add(Me.chkToritaiyo)
            Me.tabPage2.Controls.Add(Me.chkNkinkomk)
            Me.tabPage2.Controls.Add(Me.chkGazo)
            Me.tabPage2.Location = New System.Drawing.Point(4, 27)
            Me.tabPage2.Name = "tabPage2"
            Me.tabPage2.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage2.Size = New System.Drawing.Size(912, 549)
            Me.tabPage2.TabIndex = 2
            Me.tabPage2.Text = "対象項目選択"
            '
            'Label39
            '
            Me.Label39.AutoSize = True
            Me.Label39.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label39.Location = New System.Drawing.Point(59, 332)
            Me.Label39.Name = "Label39"
            Me.Label39.Size = New System.Drawing.Size(68, 18)
            Me.Label39.TabIndex = 107
            Me.Label39.Text = "鍵情報選択"
            '
            'grpKagi
            '
            Me.grpKagi.Controls.Add(Me.optKyKagi)
            Me.grpKagi.Controls.Add(Me.optHyKagi)
            Me.grpKagi.Location = New System.Drawing.Point(76, 353)
            Me.grpKagi.Name = "grpKagi"
            Me.grpKagi.Size = New System.Drawing.Size(121, 70)
            Me.grpKagi.TabIndex = 100
            Me.grpKagi.TabStop = False
            Me.grpKagi.Visible = False
            '
            'optKyKagi
            '
            Me.optKyKagi.AutoSize = True
            Me.optKyKagi.Location = New System.Drawing.Point(16, 42)
            Me.optKyKagi.Name = "optKyKagi"
            Me.optKyKagi.Size = New System.Drawing.Size(86, 22)
            Me.optKyKagi.TabIndex = 1
            Me.optKyKagi.Text = "契約鍵情報"
            Me.optKyKagi.UseVisualStyleBackColor = True
            '
            'optHyKagi
            '
            Me.optHyKagi.AutoSize = True
            Me.optHyKagi.Checked = True
            Me.optHyKagi.Location = New System.Drawing.Point(16, 16)
            Me.optHyKagi.Name = "optHyKagi"
            Me.optHyKagi.Size = New System.Drawing.Size(86, 22)
            Me.optHyKagi.TabIndex = 0
            Me.optHyKagi.TabStop = True
            Me.optHyKagi.Text = "部屋鍵情報"
            Me.optHyKagi.UseVisualStyleBackColor = True
            '
            'Label1
            '
            Me.Label1.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label1.Location = New System.Drawing.Point(203, 367)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(532, 56)
            Me.Label1.TabIndex = 99
            Me.Label1.Text = "V7の鍵情報は部屋と契約で別々に登録できますが、10では共通情報となっています。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(例えば契約の鍵情報を編集した場合、部屋情報の鍵情報にも反映されます)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "その" & _
        "為、10では部屋か契約のどちらかのみの移行となります。どちらを移行するか選択して下さい。"
            Me.Label1.UseCompatibleTextRendering = True
            Me.Label1.Visible = False
            '
            'Label34
            '
            Me.Label34.AutoSize = True
            Me.Label34.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label34.Location = New System.Drawing.Point(39, 117)
            Me.Label34.Name = "Label34"
            Me.Label34.Size = New System.Drawing.Size(100, 23)
            Me.Label34.TabIndex = 106
            Me.Label34.Text = "紐付設定項目"
            '
            'Label38
            '
            Me.Label38.BackColor = System.Drawing.SystemColors.Menu
            Me.Label38.ForeColor = System.Drawing.Color.Navy
            Me.Label38.Location = New System.Drawing.Point(30, 20)
            Me.Label38.Name = "Label38"
            Me.Label38.Size = New System.Drawing.Size(820, 80)
            Me.Label38.TabIndex = 109
            Me.Label38.Text = "「賃貸革命V7」と「賃貸革命10」の項目の紐付設定を行います。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "紐付設定を行う項目にチェックを入れて[次へ] ボタンを押して下さい。"
            Me.Label38.UseCompatibleTextRendering = True
            '
            'chkBkrui
            '
            Me.chkBkrui.AutoSize = True
            Me.chkBkrui.Location = New System.Drawing.Point(62, 153)
            Me.chkBkrui.Name = "chkBkrui"
            Me.chkBkrui.Size = New System.Drawing.Size(111, 22)
            Me.chkBkrui.TabIndex = 0
            Me.chkBkrui.Text = "物件分類マスタ"
            Me.chkBkrui.UseVisualStyleBackColor = True
            '
            'chkTosiYoto
            '
            Me.chkTosiYoto.AutoSize = True
            Me.chkTosiYoto.Location = New System.Drawing.Point(387, 237)
            Me.chkTosiYoto.Name = "chkTosiYoto"
            Me.chkTosiYoto.Size = New System.Drawing.Size(171, 22)
            Me.chkTosiYoto.TabIndex = 105
            Me.chkTosiYoto.Text = "都市計画・用途地域マスタ"
            Me.chkTosiYoto.UseVisualStyleBackColor = True
            '
            'chkKagi
            '
            Me.chkKagi.AutoSize = True
            Me.chkKagi.Location = New System.Drawing.Point(226, 249)
            Me.chkKagi.Name = "chkKagi"
            Me.chkKagi.Size = New System.Drawing.Size(123, 22)
            Me.chkKagi.TabIndex = 8
            Me.chkKagi.Text = "鍵タイトルマスタ"
            Me.chkKagi.UseVisualStyleBackColor = True
            '
            'chkHyrui
            '
            Me.chkHyrui.AutoSize = True
            Me.chkHyrui.Location = New System.Drawing.Point(62, 185)
            Me.chkHyrui.Name = "chkHyrui"
            Me.chkHyrui.Size = New System.Drawing.Size(111, 22)
            Me.chkHyrui.TabIndex = 1
            Me.chkHyrui.Text = "部屋分類マスタ"
            Me.chkHyrui.UseVisualStyleBackColor = True
            '
            'chkSetubi
            '
            Me.chkSetubi.AutoSize = True
            Me.chkSetubi.Location = New System.Drawing.Point(226, 217)
            Me.chkSetubi.Name = "chkSetubi"
            Me.chkSetubi.Size = New System.Drawing.Size(87, 22)
            Me.chkSetubi.TabIndex = 7
            Me.chkSetubi.Text = "設備マスタ"
            Me.chkSetubi.UseVisualStyleBackColor = True
            '
            'chkKyrui
            '
            Me.chkKyrui.AutoSize = True
            Me.chkKyrui.Location = New System.Drawing.Point(387, 209)
            Me.chkKyrui.Name = "chkKyrui"
            Me.chkKyrui.Size = New System.Drawing.Size(111, 22)
            Me.chkKyrui.TabIndex = 104
            Me.chkKyrui.Text = "契約分類マスタ"
            Me.chkKyrui.UseVisualStyleBackColor = True
            '
            'chkKozasyu
            '
            Me.chkKozasyu.AutoSize = True
            Me.chkKozasyu.Location = New System.Drawing.Point(226, 185)
            Me.chkKozasyu.Name = "chkKozasyu"
            Me.chkKozasyu.Size = New System.Drawing.Size(111, 22)
            Me.chkKozasyu.TabIndex = 6
            Me.chkKozasyu.Text = "口座種別マスタ"
            Me.chkKozasyu.UseVisualStyleBackColor = True
            '
            'chkNkinkbn
            '
            Me.chkNkinkbn.AutoSize = True
            Me.chkNkinkbn.Location = New System.Drawing.Point(62, 217)
            Me.chkNkinkbn.Name = "chkNkinkbn"
            Me.chkNkinkbn.Size = New System.Drawing.Size(111, 22)
            Me.chkNkinkbn.TabIndex = 2
            Me.chkNkinkbn.Text = "入金区分マスタ"
            Me.chkNkinkbn.UseVisualStyleBackColor = True
            '
            'chkKozo
            '
            Me.chkKozo.AutoSize = True
            Me.chkKozo.Location = New System.Drawing.Point(226, 153)
            Me.chkKozo.Name = "chkKozo"
            Me.chkKozo.Size = New System.Drawing.Size(87, 22)
            Me.chkKozo.TabIndex = 5
            Me.chkKozo.Text = "構造マスタ"
            Me.chkKozo.UseVisualStyleBackColor = True
            '
            'chkFBFmt
            '
            Me.chkFBFmt.AutoSize = True
            Me.chkFBFmt.Location = New System.Drawing.Point(387, 181)
            Me.chkFBFmt.Name = "chkFBFmt"
            Me.chkFBFmt.Size = New System.Drawing.Size(138, 22)
            Me.chkFBFmt.TabIndex = 103
            Me.chkFBFmt.Text = "FBフォーマット割付"
            Me.chkFBFmt.UseVisualStyleBackColor = True
            '
            'chkJisyakoza
            '
            Me.chkJisyakoza.AutoSize = True
            Me.chkJisyakoza.Location = New System.Drawing.Point(226, 281)
            Me.chkJisyakoza.Name = "chkJisyakoza"
            Me.chkJisyakoza.Size = New System.Drawing.Size(111, 22)
            Me.chkJisyakoza.TabIndex = 101
            Me.chkJisyakoza.Text = "自社口座マスタ"
            Me.chkJisyakoza.UseVisualStyleBackColor = True
            '
            'chkToritaiyo
            '
            Me.chkToritaiyo.AutoSize = True
            Me.chkToritaiyo.Location = New System.Drawing.Point(62, 249)
            Me.chkToritaiyo.Name = "chkToritaiyo"
            Me.chkToritaiyo.Size = New System.Drawing.Size(111, 22)
            Me.chkToritaiyo.TabIndex = 3
            Me.chkToritaiyo.Text = "取引態様マスタ"
            Me.chkToritaiyo.UseVisualStyleBackColor = True
            '
            'chkNkinkomk
            '
            Me.chkNkinkomk.AutoSize = True
            Me.chkNkinkomk.Location = New System.Drawing.Point(62, 281)
            Me.chkNkinkomk.Name = "chkNkinkomk"
            Me.chkNkinkomk.Size = New System.Drawing.Size(111, 22)
            Me.chkNkinkomk.TabIndex = 4
            Me.chkNkinkomk.Text = "入金項目マスタ"
            Me.chkNkinkomk.UseVisualStyleBackColor = True
            '
            'chkGazo
            '
            Me.chkGazo.AutoSize = True
            Me.chkGazo.Location = New System.Drawing.Point(387, 153)
            Me.chkGazo.Name = "chkGazo"
            Me.chkGazo.Size = New System.Drawing.Size(75, 22)
            Me.chkGazo.TabIndex = 102
            Me.chkGazo.Text = "画像割付"
            Me.chkGazo.UseVisualStyleBackColor = True
            '
            'tabPage3
            '
            Me.tabPage3.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage3.Controls.Add(Me.lblHidden4)
            Me.tabPage3.Controls.Add(Me.Label40)
            Me.tabPage3.Controls.Add(Me.Label37)
            Me.tabPage3.Controls.Add(Me.lblCaution)
            Me.tabPage3.Controls.Add(Me.tabCtrlRelwork)
            Me.tabPage3.Location = New System.Drawing.Point(4, 27)
            Me.tabPage3.Name = "tabPage3"
            Me.tabPage3.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage3.Size = New System.Drawing.Size(912, 549)
            Me.tabPage3.TabIndex = 10
            Me.tabPage3.Text = "紐付設定"
            '
            'lblHidden4
            '
            Me.lblHidden4.Location = New System.Drawing.Point(3, 81)
            Me.lblHidden4.Name = "lblHidden4"
            Me.lblHidden4.Size = New System.Drawing.Size(15, 26)
            Me.lblHidden4.TabIndex = 146
            Me.lblHidden4.Text = "　"
            '
            'Label40
            '
            Me.Label40.ForeColor = System.Drawing.Color.Black
            Me.Label40.Location = New System.Drawing.Point(10, 29)
            Me.Label40.Name = "Label40"
            Me.Label40.Size = New System.Drawing.Size(889, 20)
            Me.Label40.TabIndex = 145
            Me.Label40.Text = "・初期画面表示時に、関連する10項目値が予めセットされます。(適正な項目値が設定されていない場合、手動で変更して下さい。)"
            Me.Label40.UseCompatibleTextRendering = True
            '
            'Label37
            '
            Me.Label37.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label37.ForeColor = System.Drawing.Color.Red
            Me.Label37.Location = New System.Drawing.Point(10, 49)
            Me.Label37.Name = "Label37"
            Me.Label37.Size = New System.Drawing.Size(889, 20)
            Me.Label37.TabIndex = 144
            Me.Label37.Text = "※項目数(列)またはデータ数(行)が多い場合、表示内容が画面外に隠れている場合があります。スクロールバーを動かして見落とさないようにして下さい。"
            Me.Label37.UseCompatibleTextRendering = True
            '
            'lblCaution
            '
            Me.lblCaution.ForeColor = System.Drawing.Color.Black
            Me.lblCaution.Location = New System.Drawing.Point(10, 9)
            Me.lblCaution.Name = "lblCaution"
            Me.lblCaution.Size = New System.Drawing.Size(889, 20)
            Me.lblCaution.TabIndex = 124
            Me.lblCaution.Text = "・紐付が未設定の項目は移行されません。全ての項目について紐付設定を行って下さい。"
            Me.lblCaution.UseCompatibleTextRendering = True
            '
            'tabCtrlRelwork
            '
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage10)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage11)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage12)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage13)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage14)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage15)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage16)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage17)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage18)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage19)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage20)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage21)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage22)
            Me.tabCtrlRelwork.Controls.Add(Me.tabPage23)
            Me.tabCtrlRelwork.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
            Me.tabCtrlRelwork.Location = New System.Drawing.Point(13, 81)
            Me.tabCtrlRelwork.Name = "tabCtrlRelwork"
            Me.tabCtrlRelwork.SelectedIndex = 0
            Me.tabCtrlRelwork.Size = New System.Drawing.Size(893, 462)
            Me.tabCtrlRelwork.TabIndex = 123
            '
            'tabPage10
            '
            Me.tabPage10.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage10.Controls.Add(Me.Label6)
            Me.tabPage10.Controls.Add(Me.dgvBkrui)
            Me.tabPage10.Location = New System.Drawing.Point(4, 27)
            Me.tabPage10.Name = "tabPage10"
            Me.tabPage10.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage10.Size = New System.Drawing.Size(885, 431)
            Me.tabPage10.TabIndex = 0
            Me.tabPage10.Text = "物件分類マスタ"
            '
            'Label6
            '
            Me.Label6.Location = New System.Drawing.Point(15, 10)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(840, 55)
            Me.Label6.TabIndex = 98
            Me.Label6.Text = "「物件分類」の紐付設定を行って下さい。([物件分類No]の最大設定値=9999)"
            Me.Label6.UseCompatibleTextRendering = True
            '
            'dgvBkrui
            '
            Me.dgvBkrui.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvBkrui.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBkrui0, Me.ColBkrui1, Me.ColBkrui2})
            Me.dgvBkrui.Location = New System.Drawing.Point(10, 73)
            Me.dgvBkrui.Name = "dgvBkrui"
            Me.dgvBkrui.RowTemplate.Height = 21
            Me.dgvBkrui.Size = New System.Drawing.Size(865, 352)
            Me.dgvBkrui.TabIndex = 94
            '
            'ColBkrui0
            '
            Me.ColBkrui0.HeaderText = "物件分類名"
            Me.ColBkrui0.Name = "ColBkrui0"
            Me.ColBkrui0.Width = 150
            '
            'ColBkrui1
            '
            Me.ColBkrui1.HeaderText = "物件分類No"
            Me.ColBkrui1.Name = "ColBkrui1"
            Me.ColBkrui1.Width = 150
            '
            'ColBkrui2
            '
            Me.ColBkrui2.HeaderText = "物件分類名"
            Me.ColBkrui2.Name = "ColBkrui2"
            Me.ColBkrui2.Width = 150
            '
            'tabPage11
            '
            Me.tabPage11.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage11.Controls.Add(Me.Label7)
            Me.tabPage11.Controls.Add(Me.dgvHyrui)
            Me.tabPage11.Location = New System.Drawing.Point(4, 27)
            Me.tabPage11.Name = "tabPage11"
            Me.tabPage11.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage11.Size = New System.Drawing.Size(885, 431)
            Me.tabPage11.TabIndex = 1
            Me.tabPage11.Text = "部屋分類マスタ"
            '
            'Label7
            '
            Me.Label7.Location = New System.Drawing.Point(15, 10)
            Me.Label7.Name = "Label7"
            Me.Label7.Size = New System.Drawing.Size(840, 55)
            Me.Label7.TabIndex = 99
            Me.Label7.Text = "「部屋分類」の紐付設定を行って下さい。([部屋分類No]の最大設定値=9999)"
            Me.Label7.UseCompatibleTextRendering = True
            '
            'dgvHyrui
            '
            Me.dgvHyrui.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvHyrui.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColHyrui0, Me.ColHyrui1, Me.ColHyrui2})
            Me.dgvHyrui.Location = New System.Drawing.Point(10, 73)
            Me.dgvHyrui.Name = "dgvHyrui"
            Me.dgvHyrui.RowTemplate.Height = 21
            Me.dgvHyrui.Size = New System.Drawing.Size(865, 352)
            Me.dgvHyrui.TabIndex = 95
            '
            'ColHyrui0
            '
            Me.ColHyrui0.HeaderText = "部屋分類名"
            Me.ColHyrui0.Name = "ColHyrui0"
            Me.ColHyrui0.Width = 200
            '
            'ColHyrui1
            '
            Me.ColHyrui1.HeaderText = "部屋分類No"
            Me.ColHyrui1.Name = "ColHyrui1"
            Me.ColHyrui1.Width = 200
            '
            'ColHyrui2
            '
            Me.ColHyrui2.HeaderText = "部屋分類名"
            Me.ColHyrui2.Name = "ColHyrui2"
            Me.ColHyrui2.Width = 200
            '
            'tabPage12
            '
            Me.tabPage12.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage12.Controls.Add(Me.Label8)
            Me.tabPage12.Controls.Add(Me.dgvNkinKbn)
            Me.tabPage12.Location = New System.Drawing.Point(4, 27)
            Me.tabPage12.Name = "tabPage12"
            Me.tabPage12.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage12.Size = New System.Drawing.Size(885, 431)
            Me.tabPage12.TabIndex = 2
            Me.tabPage12.Text = "入金区分マスタ"
            '
            'Label8
            '
            Me.Label8.Location = New System.Drawing.Point(15, 10)
            Me.Label8.Name = "Label8"
            Me.Label8.Size = New System.Drawing.Size(840, 55)
            Me.Label8.TabIndex = 99
            Me.Label8.Text = "「入金区分」の紐付設定を行って下さい。([入金区分No]の最大設定値=20)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "※属性 (1:現金, 2:振込, 3:振替, 4:その他, 99:移動) を設定す" & _
        "る必要があります。"
            Me.Label8.UseCompatibleTextRendering = True
            '
            'dgvNkinKbn
            '
            Me.dgvNkinKbn.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvNkinKbn.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNkinKbn0, Me.ColNkinKbn1, Me.ColNkinKbn2, Me.ColNkinKbn3, Me.ColNkinKbn4})
            Me.dgvNkinKbn.Location = New System.Drawing.Point(10, 73)
            Me.dgvNkinKbn.Name = "dgvNkinKbn"
            Me.dgvNkinKbn.RowTemplate.Height = 21
            Me.dgvNkinKbn.Size = New System.Drawing.Size(865, 352)
            Me.dgvNkinKbn.TabIndex = 96
            '
            'ColNkinKbn0
            '
            Me.ColNkinKbn0.HeaderText = "入金区分名"
            Me.ColNkinKbn0.Name = "ColNkinKbn0"
            Me.ColNkinKbn0.Width = 150
            '
            'ColNkinKbn1
            '
            Me.ColNkinKbn1.HeaderText = "入金区分No"
            Me.ColNkinKbn1.Name = "ColNkinKbn1"
            Me.ColNkinKbn1.Width = 150
            '
            'ColNkinKbn2
            '
            Me.ColNkinKbn2.HeaderText = "入金区分名"
            Me.ColNkinKbn2.Name = "ColNkinKbn2"
            Me.ColNkinKbn2.Width = 150
            '
            'ColNkinKbn3
            '
            Me.ColNkinKbn3.HeaderText = "入金区分属性No"
            Me.ColNkinKbn3.Name = "ColNkinKbn3"
            Me.ColNkinKbn3.Width = 150
            '
            'ColNkinKbn4
            '
            Me.ColNkinKbn4.HeaderText = "入金区分属性名"
            Me.ColNkinKbn4.Name = "ColNkinKbn4"
            Me.ColNkinKbn4.Width = 150
            '
            'tabPage13
            '
            Me.tabPage13.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage13.Controls.Add(Me.Label9)
            Me.tabPage13.Controls.Add(Me.dgvToritaiyo)
            Me.tabPage13.Location = New System.Drawing.Point(4, 27)
            Me.tabPage13.Name = "tabPage13"
            Me.tabPage13.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage13.Size = New System.Drawing.Size(885, 431)
            Me.tabPage13.TabIndex = 10
            Me.tabPage13.Text = "取引態様マスタ"
            '
            'Label9
            '
            Me.Label9.Location = New System.Drawing.Point(15, 10)
            Me.Label9.Name = "Label9"
            Me.Label9.Size = New System.Drawing.Size(840, 55)
            Me.Label9.TabIndex = 100
            Me.Label9.Text = "「取引態様」の紐付設定を行って下さい。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "・V7の「媒介」「仲介」は、10の「一般媒介」に自動設定します。"
            Me.Label9.UseCompatibleTextRendering = True
            '
            'dgvToritaiyo
            '
            Me.dgvToritaiyo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvToritaiyo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColToritaiyo0, Me.ColToritaiyo1, Me.ColToritaiyo2})
            Me.dgvToritaiyo.Location = New System.Drawing.Point(10, 73)
            Me.dgvToritaiyo.Name = "dgvToritaiyo"
            Me.dgvToritaiyo.RowTemplate.Height = 21
            Me.dgvToritaiyo.Size = New System.Drawing.Size(865, 352)
            Me.dgvToritaiyo.TabIndex = 99
            '
            'ColToritaiyo0
            '
            Me.ColToritaiyo0.HeaderText = "取引態様名"
            Me.ColToritaiyo0.Name = "ColToritaiyo0"
            Me.ColToritaiyo0.Width = 200
            '
            'ColToritaiyo1
            '
            Me.ColToritaiyo1.HeaderText = "取引態様No"
            Me.ColToritaiyo1.Name = "ColToritaiyo1"
            Me.ColToritaiyo1.Width = 200
            '
            'ColToritaiyo2
            '
            Me.ColToritaiyo2.HeaderText = "取引態様名"
            Me.ColToritaiyo2.Name = "ColToritaiyo2"
            Me.ColToritaiyo2.Width = 200
            '
            'tabPage14
            '
            Me.tabPage14.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage14.Controls.Add(Me.Panel1)
            Me.tabPage14.Controls.Add(Me.chkDuplicate)
            Me.tabPage14.Controls.Add(Me.Label10)
            Me.tabPage14.Controls.Add(Me.dgvNkinkomk)
            Me.tabPage14.Location = New System.Drawing.Point(4, 27)
            Me.tabPage14.Name = "tabPage14"
            Me.tabPage14.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage14.Size = New System.Drawing.Size(885, 431)
            Me.tabPage14.TabIndex = 11
            Me.tabPage14.Text = "入金項目マスタ"
            '
            'Panel1
            '
            Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.Panel1.Controls.Add(Me.Label44)
            Me.Panel1.Controls.Add(Me.Label43)
            Me.Panel1.Controls.Add(Me.Label31)
            Me.Panel1.Controls.Add(Me.Label42)
            Me.Panel1.Controls.Add(Me.Label17)
            Me.Panel1.Controls.Add(Me.Label18)
            Me.Panel1.Controls.Add(Me.btnSetNkinnoBulk)
            Me.Panel1.Controls.Add(Me.txtNkinnoSta)
            Me.Panel1.Location = New System.Drawing.Point(10, 67)
            Me.Panel1.Name = "Panel1"
            Me.Panel1.Size = New System.Drawing.Size(865, 103)
            Me.Panel1.TabIndex = 134
            '
            'Label44
            '
            Me.Label44.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label44.Location = New System.Drawing.Point(221, 55)
            Me.Label44.Name = "Label44"
            Me.Label44.Size = New System.Drawing.Size(632, 20)
            Me.Label44.TabIndex = 135
            Me.Label44.Text = "※[変動費分類メーターNo]は除きます。(手動で設定して下さい)"
            Me.Label44.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label44.UseCompatibleTextRendering = True
            '
            'Label43
            '
            Me.Label43.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label43.Location = New System.Drawing.Point(16, 62)
            Me.Label43.Name = "Label43"
            Me.Label43.Size = New System.Drawing.Size(110, 24)
            Me.Label43.TabIndex = 134
            Me.Label43.Text = "コード一括設定"
            Me.Label43.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label43.UseCompatibleTextRendering = True
            '
            'Label31
            '
            Me.Label31.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label31.Location = New System.Drawing.Point(221, 76)
            Me.Label31.Name = "Label31"
            Me.Label31.Size = New System.Drawing.Size(632, 20)
            Me.Label31.TabIndex = 133
            Me.Label31.Text = "例) 開始Noが「500」かつ入金区分が「契約時(範囲 2000~2999)」の場合、「2500,2501,2502, …」を設定します。"
            Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label31.UseCompatibleTextRendering = True
            '
            'Label42
            '
            Me.Label42.AutoSize = True
            Me.Label42.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label42.Location = New System.Drawing.Point(3, 3)
            Me.Label42.Name = "Label42"
            Me.Label42.Size = New System.Drawing.Size(116, 18)
            Me.Label42.TabIndex = 133
            Me.Label42.Text = "任意コード一括設定"
            '
            'Label17
            '
            Me.Label17.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label17.Location = New System.Drawing.Point(221, 21)
            Me.Label17.Name = "Label17"
            Me.Label17.Size = New System.Drawing.Size(632, 35)
            Me.Label17.TabIndex = 129
            Me.Label17.Text = "賃貸革命10の入金項目が未設定の箇所へ、開始コードを基準に任意のコード(既に登録されているコード以外の重複しないコード)を一括設定します。任意のコードが設定された" & _
        "データはコンバート時に新規の入金項目データとして移行されます。"
            Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label17.UseCompatibleTextRendering = True
            '
            'Label18
            '
            Me.Label18.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label18.Location = New System.Drawing.Point(16, 31)
            Me.Label18.Name = "Label18"
            Me.Label18.Size = New System.Drawing.Size(110, 24)
            Me.Label18.TabIndex = 130
            Me.Label18.Text = "開始コード(下3桁)"
            Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.Label18.UseCompatibleTextRendering = True
            '
            'btnSetNkinnoBulk
            '
            Me.btnSetNkinnoBulk.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.btnSetNkinnoBulk.Image = CType(resources.GetObject("btnSetNkinnoBulk.Image"), System.Drawing.Image)
            Me.btnSetNkinnoBulk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnSetNkinnoBulk.Location = New System.Drawing.Point(132, 61)
            Me.btnSetNkinnoBulk.Name = "btnSetNkinnoBulk"
            Me.btnSetNkinnoBulk.Size = New System.Drawing.Size(67, 27)
            Me.btnSetNkinnoBulk.TabIndex = 127
            Me.btnSetNkinnoBulk.Text = " 実行  "
            Me.btnSetNkinnoBulk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.btnSetNkinnoBulk.UseVisualStyleBackColor = True
            '
            'txtNkinnoSta
            '
            Me.txtNkinnoSta.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.txtNkinnoSta.Location = New System.Drawing.Point(132, 32)
            Me.txtNkinnoSta.Multiline = True
            Me.txtNkinnoSta.Name = "txtNkinnoSta"
            Me.txtNkinnoSta.Size = New System.Drawing.Size(67, 23)
            Me.txtNkinnoSta.TabIndex = 132
            Me.txtNkinnoSta.Text = "500"
            Me.txtNkinnoSta.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
            '
            'chkDuplicate
            '
            Me.chkDuplicate.AutoSize = True
            Me.chkDuplicate.Location = New System.Drawing.Point(764, 15)
            Me.chkDuplicate.Name = "chkDuplicate"
            Me.chkDuplicate.Size = New System.Drawing.Size(111, 22)
            Me.chkDuplicate.TabIndex = 125
            Me.chkDuplicate.Text = "重複を許可する"
            Me.chkDuplicate.UseVisualStyleBackColor = True
            Me.chkDuplicate.Visible = False
            '
            'Label10
            '
            Me.Label10.Location = New System.Drawing.Point(15, 10)
            Me.Label10.Name = "Label10"
            Me.Label10.Size = New System.Drawing.Size(625, 54)
            Me.Label10.TabIndex = 126
            Me.Label10.Text = "「入金項目」の紐付設定を行って下さい。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "  ・[入金項目No(名称)]-[入金項目属性No(名称)]-[変動費メーターNo(名称)]の順に入力して下さい。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " " & _
        " ・同内容(項目No + 属性No + メーターNo)は設定できません。"
            Me.Label10.UseCompatibleTextRendering = True
            '
            'dgvNkinkomk
            '
            Me.dgvNkinkomk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvNkinkomk.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNkinkomk0, Me.ColNkinkomk1, Me.ColNkinkomk2, Me.ColNkinkomk3, Me.ColNkinkomk4, Me.ColNkinkomk5, Me.ColNkinkomk6, Me.ColNkinkomk7, Me.ColNkinkomk8})
            Me.dgvNkinkomk.Location = New System.Drawing.Point(10, 176)
            Me.dgvNkinkomk.Name = "dgvNkinkomk"
            Me.dgvNkinkomk.RowTemplate.Height = 21
            Me.dgvNkinkomk.Size = New System.Drawing.Size(865, 250)
            Me.dgvNkinkomk.TabIndex = 100
            '
            'ColNkinkomk0
            '
            Me.ColNkinkomk0.HeaderText = "入金項目名称"
            Me.ColNkinkomk0.Name = "ColNkinkomk0"
            '
            'ColNkinkomk1
            '
            Me.ColNkinkomk1.HeaderText = "入金項目区分"
            Me.ColNkinkomk1.Name = "ColNkinkomk1"
            '
            'ColNkinkomk2
            '
            Me.ColNkinkomk2.HeaderText = "入金項目No"
            Me.ColNkinkomk2.Name = "ColNkinkomk2"
            '
            'ColNkinkomk3
            '
            Me.ColNkinkomk3.HeaderText = "入金項目名称"
            Me.ColNkinkomk3.Name = "ColNkinkomk3"
            '
            'ColNkinkomk4
            '
            Me.ColNkinkomk4.HeaderText = "入金項目属性No"
            Me.ColNkinkomk4.Name = "ColNkinkomk4"
            '
            'ColNkinkomk5
            '
            Me.ColNkinkomk5.HeaderText = "入金項目属性名称"
            Me.ColNkinkomk5.Name = "ColNkinkomk5"
            '
            'ColNkinkomk6
            '
            Me.ColNkinkomk6.HeaderText = "入金項目区分"
            Me.ColNkinkomk6.Name = "ColNkinkomk6"
            '
            'ColNkinkomk7
            '
            Me.ColNkinkomk7.HeaderText = "変動費メーター分類No"
            Me.ColNkinkomk7.Name = "ColNkinkomk7"
            '
            'ColNkinkomk8
            '
            Me.ColNkinkomk8.HeaderText = "変動費メーター分類名"
            Me.ColNkinkomk8.Name = "ColNkinkomk8"
            '
            'tabPage15
            '
            Me.tabPage15.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage15.Controls.Add(Me.Label12)
            Me.tabPage15.Controls.Add(Me.dgvKozo)
            Me.tabPage15.Location = New System.Drawing.Point(4, 27)
            Me.tabPage15.Name = "tabPage15"
            Me.tabPage15.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage15.Size = New System.Drawing.Size(885, 431)
            Me.tabPage15.TabIndex = 12
            Me.tabPage15.Text = "物件構造マスタ"
            '
            'Label12
            '
            Me.Label12.Location = New System.Drawing.Point(15, 10)
            Me.Label12.Name = "Label12"
            Me.Label12.Size = New System.Drawing.Size(840, 55)
            Me.Label12.TabIndex = 102
            Me.Label12.Text = "「構造」の紐付設定を行って下さい。"
            Me.Label12.UseCompatibleTextRendering = True
            '
            'dgvKozo
            '
            Me.dgvKozo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvKozo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColKozo0, Me.ColKozo1, Me.ColKozo2})
            Me.dgvKozo.Location = New System.Drawing.Point(10, 73)
            Me.dgvKozo.Name = "dgvKozo"
            Me.dgvKozo.RowTemplate.Height = 21
            Me.dgvKozo.Size = New System.Drawing.Size(865, 352)
            Me.dgvKozo.TabIndex = 101
            '
            'ColKozo0
            '
            Me.ColKozo0.HeaderText = "構造名"
            Me.ColKozo0.Name = "ColKozo0"
            Me.ColKozo0.Width = 200
            '
            'ColKozo1
            '
            Me.ColKozo1.HeaderText = "構造No"
            Me.ColKozo1.Name = "ColKozo1"
            Me.ColKozo1.Width = 200
            '
            'ColKozo2
            '
            Me.ColKozo2.HeaderText = "構造名"
            Me.ColKozo2.Name = "ColKozo2"
            Me.ColKozo2.Width = 200
            '
            'tabPage16
            '
            Me.tabPage16.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage16.Controls.Add(Me.Label14)
            Me.tabPage16.Controls.Add(Me.dgvKozasyu)
            Me.tabPage16.Location = New System.Drawing.Point(4, 27)
            Me.tabPage16.Name = "tabPage16"
            Me.tabPage16.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage16.Size = New System.Drawing.Size(885, 431)
            Me.tabPage16.TabIndex = 13
            Me.tabPage16.Text = "口座種別マスタ"
            '
            'Label14
            '
            Me.Label14.Location = New System.Drawing.Point(15, 10)
            Me.Label14.Name = "Label14"
            Me.Label14.Size = New System.Drawing.Size(840, 55)
            Me.Label14.TabIndex = 103
            Me.Label14.Text = "「口座種別」の紐付設定を行って下さい。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "・初期表示時の紐付かない項目については「その他」が設定されます。"
            Me.Label14.UseCompatibleTextRendering = True
            '
            'dgvKozasyu
            '
            Me.dgvKozasyu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvKozasyu.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColKozasyu0, Me.ColKozasyu1, Me.ColKozasyu2})
            Me.dgvKozasyu.Location = New System.Drawing.Point(10, 73)
            Me.dgvKozasyu.Name = "dgvKozasyu"
            Me.dgvKozasyu.RowTemplate.Height = 21
            Me.dgvKozasyu.Size = New System.Drawing.Size(865, 352)
            Me.dgvKozasyu.TabIndex = 102
            '
            'ColKozasyu0
            '
            Me.ColKozasyu0.HeaderText = "口座種別名称"
            Me.ColKozasyu0.Name = "ColKozasyu0"
            Me.ColKozasyu0.Width = 200
            '
            'ColKozasyu1
            '
            Me.ColKozasyu1.HeaderText = "口座種別No"
            Me.ColKozasyu1.Name = "ColKozasyu1"
            Me.ColKozasyu1.Width = 200
            '
            'ColKozasyu2
            '
            Me.ColKozasyu2.HeaderText = "口座種別名称"
            Me.ColKozasyu2.Name = "ColKozasyu2"
            Me.ColKozasyu2.Width = 200
            '
            'tabPage17
            '
            Me.tabPage17.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage17.Controls.Add(Me.Label15)
            Me.tabPage17.Controls.Add(Me.dgvSetubi)
            Me.tabPage17.Location = New System.Drawing.Point(4, 27)
            Me.tabPage17.Name = "tabPage17"
            Me.tabPage17.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage17.Size = New System.Drawing.Size(885, 431)
            Me.tabPage17.TabIndex = 14
            Me.tabPage17.Text = "部屋設備"
            '
            'Label15
            '
            Me.Label15.Location = New System.Drawing.Point(15, 10)
            Me.Label15.Name = "Label15"
            Me.Label15.Size = New System.Drawing.Size(860, 57)
            Me.Label15.TabIndex = 104
            Me.Label15.Text = "「設備」の紐付設定を行って下さい。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "・[設備グループNo(名)]-[設備No(名)]-[設備内容No(名)]の順に入力して下さい。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "・紐付されなかったデータは" & _
        "新規の設備データとして移行します。"
            Me.Label15.UseCompatibleTextRendering = True
            '
            'dgvSetubi
            '
            Me.dgvSetubi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvSetubi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSetubi0, Me.ColSetubi1, Me.ColSetubi2, Me.ColSetubi3, Me.ColSetubi4, Me.ColSetubi5, Me.ColSetubi6, Me.ColSetubi7})
            Me.dgvSetubi.Location = New System.Drawing.Point(10, 73)
            Me.dgvSetubi.Name = "dgvSetubi"
            Me.dgvSetubi.RowTemplate.Height = 21
            Me.dgvSetubi.Size = New System.Drawing.Size(865, 352)
            Me.dgvSetubi.TabIndex = 103
            '
            'ColSetubi0
            '
            Me.ColSetubi0.HeaderText = "設備名"
            Me.ColSetubi0.Name = "ColSetubi0"
            '
            'ColSetubi1
            '
            Me.ColSetubi1.HeaderText = "設備内容"
            Me.ColSetubi1.Name = "ColSetubi1"
            '
            'ColSetubi2
            '
            Me.ColSetubi2.HeaderText = "設備グループNo"
            Me.ColSetubi2.Name = "ColSetubi2"
            '
            'ColSetubi3
            '
            Me.ColSetubi3.HeaderText = "設備グループ名"
            Me.ColSetubi3.Name = "ColSetubi3"
            '
            'ColSetubi4
            '
            Me.ColSetubi4.HeaderText = "設備No"
            Me.ColSetubi4.Name = "ColSetubi4"
            '
            'ColSetubi5
            '
            Me.ColSetubi5.HeaderText = "設備名"
            Me.ColSetubi5.Name = "ColSetubi5"
            '
            'ColSetubi6
            '
            Me.ColSetubi6.HeaderText = "設備内容No"
            Me.ColSetubi6.Name = "ColSetubi6"
            '
            'ColSetubi7
            '
            Me.ColSetubi7.HeaderText = "設備内容"
            Me.ColSetubi7.Name = "ColSetubi7"
            '
            'tabPage18
            '
            Me.tabPage18.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage18.Controls.Add(Me.lblLine2)
            Me.tabPage18.Controls.Add(Me.Label2)
            Me.tabPage18.Controls.Add(Me.dgvKagi)
            Me.tabPage18.Location = New System.Drawing.Point(4, 27)
            Me.tabPage18.Name = "tabPage18"
            Me.tabPage18.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage18.Size = New System.Drawing.Size(885, 431)
            Me.tabPage18.TabIndex = 15
            Me.tabPage18.Text = "共用鍵設定"
            '
            'lblLine2
            '
            Me.lblLine2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblLine2.Location = New System.Drawing.Point(-1, 0)
            Me.lblLine2.Name = "lblLine2"
            Me.lblLine2.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.lblLine2.Size = New System.Drawing.Size(885, 2)
            Me.lblLine2.TabIndex = 146
            '
            'Label2
            '
            Me.Label2.Location = New System.Drawing.Point(15, 10)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(840, 55)
            Me.Label2.TabIndex = 97
            Me.Label2.Text = "物件に紐付く「共用鍵」と部屋(契約)に紐付く「専用鍵」のどちらかに設定することができます。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "共用の場合は[共用チェック]列にチェックを入れて下さい。"
            Me.Label2.UseCompatibleTextRendering = True
            '
            'dgvKagi
            '
            Me.dgvKagi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvKagi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2})
            Me.dgvKagi.Location = New System.Drawing.Point(10, 73)
            Me.dgvKagi.Name = "dgvKagi"
            Me.dgvKagi.RowTemplate.Height = 21
            Me.dgvKagi.Size = New System.Drawing.Size(865, 352)
            Me.dgvKagi.TabIndex = 95
            '
            'DataGridViewTextBoxColumn1
            '
            Me.DataGridViewTextBoxColumn1.HeaderText = "行No"
            Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
            Me.DataGridViewTextBoxColumn1.Width = 150
            '
            'DataGridViewTextBoxColumn2
            '
            Me.DataGridViewTextBoxColumn2.HeaderText = "鍵名"
            Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
            Me.DataGridViewTextBoxColumn2.Width = 150
            '
            'tabPage19
            '
            Me.tabPage19.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage19.Controls.Add(Me.Label45)
            Me.tabPage19.Controls.Add(Me.Label5)
            Me.tabPage19.Controls.Add(Me.dgvJisyakoza)
            Me.tabPage19.Location = New System.Drawing.Point(4, 27)
            Me.tabPage19.Name = "tabPage19"
            Me.tabPage19.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage19.Size = New System.Drawing.Size(885, 431)
            Me.tabPage19.TabIndex = 16
            Me.tabPage19.Text = "自社口座マスタ"
            '
            'Label45
            '
            Me.Label45.ForeColor = System.Drawing.Color.Red
            Me.Label45.Location = New System.Drawing.Point(15, 29)
            Me.Label45.Name = "Label45"
            Me.Label45.Size = New System.Drawing.Size(840, 41)
            Me.Label45.TabIndex = 99
            Me.Label45.Text = "※10では各口座情報(振込・振替・依頼人口座(総合振込))に自社口座を紐付ける必要があります。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "※[自社口座No]は10の口座の並び順(登録順)となります。尚、" & _
        "未設定の場合は任意Noを自動で設定します。"
            Me.Label45.UseCompatibleTextRendering = True
            '
            'Label5
            '
            Me.Label5.ForeColor = System.Drawing.Color.Black
            Me.Label5.Location = New System.Drawing.Point(15, 10)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(840, 19)
            Me.Label5.TabIndex = 98
            Me.Label5.Text = "「自社口座」の紐付設定を行って下さい。"
            Me.Label5.UseCompatibleTextRendering = True
            '
            'dgvJisyakoza
            '
            Me.dgvJisyakoza.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvJisyakoza.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColJisyakoza0, Me.ColJisyakoza1, Me.ColJisyakoza2, Me.ColJisyakoza3, Me.ColJisyakoza4, Me.ColJisyakoza5, Me.ColJisyakoza6, Me.ColJisyakoza7, Me.ColJisyakoza8, Me.ColJisyakoza9, Me.ColJisyakoza10, Me.ColJisyakoza11, Me.ColJisyakoza12, Me.ColJisyakoza13, Me.ColJisyakoza14, Me.ColJisyakoza15, Me.ColJisyakoza16, Me.ColJisyakoza17})
            Me.dgvJisyakoza.Location = New System.Drawing.Point(10, 73)
            Me.dgvJisyakoza.Name = "dgvJisyakoza"
            Me.dgvJisyakoza.RowTemplate.Height = 21
            Me.dgvJisyakoza.Size = New System.Drawing.Size(865, 352)
            Me.dgvJisyakoza.TabIndex = 96
            '
            'ColJisyakoza0
            '
            Me.ColJisyakoza0.HeaderText = "口座取得元"
            Me.ColJisyakoza0.Name = "ColJisyakoza0"
            '
            'ColJisyakoza1
            '
            Me.ColJisyakoza1.HeaderText = "各口座設定No"
            Me.ColJisyakoza1.Name = "ColJisyakoza1"
            '
            'ColJisyakoza2
            '
            Me.ColJisyakoza2.HeaderText = "各口座設定名称"
            Me.ColJisyakoza2.Name = "ColJisyakoza2"
            '
            'ColJisyakoza3
            '
            Me.ColJisyakoza3.HeaderText = "金融機関No"
            Me.ColJisyakoza3.Name = "ColJisyakoza3"
            '
            'ColJisyakoza4
            '
            Me.ColJisyakoza4.HeaderText = "金融機関名"
            Me.ColJisyakoza4.Name = "ColJisyakoza4"
            '
            'ColJisyakoza5
            '
            Me.ColJisyakoza5.HeaderText = "支店No"
            Me.ColJisyakoza5.Name = "ColJisyakoza5"
            '
            'ColJisyakoza6
            '
            Me.ColJisyakoza6.HeaderText = "支店名"
            Me.ColJisyakoza6.Name = "ColJisyakoza6"
            '
            'ColJisyakoza7
            '
            Me.ColJisyakoza7.HeaderText = "口座種別"
            Me.ColJisyakoza7.Name = "ColJisyakoza7"
            '
            'ColJisyakoza8
            '
            Me.ColJisyakoza8.HeaderText = "口座番号"
            Me.ColJisyakoza8.Name = "ColJisyakoza8"
            '
            'ColJisyakoza9
            '
            Me.ColJisyakoza9.HeaderText = "口座名義"
            Me.ColJisyakoza9.Name = "ColJisyakoza9"
            '
            'ColJisyakoza10
            '
            Me.ColJisyakoza10.HeaderText = "口座名義カナ"
            Me.ColJisyakoza10.Name = "ColJisyakoza10"
            '
            'ColJisyakoza11
            '
            Me.ColJisyakoza11.HeaderText = "ゆうちょ記号1"
            Me.ColJisyakoza11.Name = "ColJisyakoza11"
            '
            'ColJisyakoza12
            '
            Me.ColJisyakoza12.HeaderText = "ゆうちょ記号2"
            Me.ColJisyakoza12.Name = "ColJisyakoza12"
            '
            'ColJisyakoza13
            '
            Me.ColJisyakoza13.HeaderText = "ゆうちょ口座番号"
            Me.ColJisyakoza13.Name = "ColJisyakoza13"
            '
            'ColJisyakoza14
            '
            Me.ColJisyakoza14.HeaderText = "備考"
            Me.ColJisyakoza14.Name = "ColJisyakoza14"
            '
            'ColJisyakoza15
            '
            Me.ColJisyakoza15.HeaderText = "自社No"
            Me.ColJisyakoza15.Name = "ColJisyakoza15"
            '
            'ColJisyakoza16
            '
            Me.ColJisyakoza16.HeaderText = "自社名"
            Me.ColJisyakoza16.Name = "ColJisyakoza16"
            '
            'ColJisyakoza17
            '
            Me.ColJisyakoza17.HeaderText = "並び順"
            Me.ColJisyakoza17.Name = "ColJisyakoza17"
            '
            'tabPage20
            '
            Me.tabPage20.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage20.Controls.Add(Me.Label3)
            Me.tabPage20.Controls.Add(Me.dgvGazo)
            Me.tabPage20.Location = New System.Drawing.Point(4, 27)
            Me.tabPage20.Name = "tabPage20"
            Me.tabPage20.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage20.Size = New System.Drawing.Size(885, 431)
            Me.tabPage20.TabIndex = 17
            Me.tabPage20.Text = "周辺画像"
            '
            'Label3
            '
            Me.Label3.Location = New System.Drawing.Point(15, 10)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(840, 80)
            Me.Label3.TabIndex = 99
            Me.Label3.Text = "「画像」の割り付けを行って下さい。(紐付必須項目です)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "  ・賃貸革命V7の「物件画像」を、10の「物件画像」か「周辺画像」に割り付けることができます。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "  " & _
        "・「周辺画像」に割り付けた場合、「施設分類名」を設定することができます。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "※[ボタン位置]列…V7の物件情報の物件画像登録ボタンの位置のことです。(左上から右" & _
        "下へ1～20)"
            Me.Label3.UseCompatibleTextRendering = True
            '
            'dgvGazo
            '
            Me.dgvGazo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvGazo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColGazo0, Me.ColGazo1, Me.ColGazo2, Me.ColGazo3, Me.ColGazo4, Me.ColGazo5})
            Me.dgvGazo.Location = New System.Drawing.Point(10, 93)
            Me.dgvGazo.Name = "dgvGazo"
            Me.dgvGazo.RowTemplate.Height = 21
            Me.dgvGazo.Size = New System.Drawing.Size(865, 332)
            Me.dgvGazo.TabIndex = 97
            '
            'ColGazo0
            '
            Me.ColGazo0.HeaderText = "画像種別"
            Me.ColGazo0.Name = "ColGazo0"
            Me.ColGazo0.Width = 150
            '
            'ColGazo1
            '
            Me.ColGazo1.HeaderText = "ボタン位置"
            Me.ColGazo1.Name = "ColGazo1"
            Me.ColGazo1.Width = 150
            '
            'ColGazo2
            '
            Me.ColGazo2.HeaderText = "ボタン名"
            Me.ColGazo2.Name = "ColGazo2"
            Me.ColGazo2.Width = 150
            '
            'ColGazo3
            '
            Me.ColGazo3.HeaderText = "移行先画像種別"
            Me.ColGazo3.Name = "ColGazo3"
            Me.ColGazo3.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
            Me.ColGazo3.Width = 150
            '
            'ColGazo4
            '
            Me.ColGazo4.HeaderText = "並び順"
            Me.ColGazo4.Name = "ColGazo4"
            Me.ColGazo4.Width = 150
            '
            'ColGazo5
            '
            Me.ColGazo5.HeaderText = "施設分類名"
            Me.ColGazo5.Name = "ColGazo5"
            '
            'tabPage21
            '
            Me.tabPage21.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage21.Controls.Add(Me.Label46)
            Me.tabPage21.Controls.Add(Me.Label32)
            Me.tabPage21.Controls.Add(Me.dgvFBInfo)
            Me.tabPage21.Location = New System.Drawing.Point(4, 27)
            Me.tabPage21.Name = "tabPage21"
            Me.tabPage21.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage21.Size = New System.Drawing.Size(885, 431)
            Me.tabPage21.TabIndex = 18
            Me.tabPage21.Text = "FBデータフォーマット名"
            '
            'Label46
            '
            Me.Label46.ForeColor = System.Drawing.Color.Red
            Me.Label46.Location = New System.Drawing.Point(15, 30)
            Me.Label46.Name = "Label46"
            Me.Label46.Size = New System.Drawing.Size(840, 20)
            Me.Label46.TabIndex = 100
            Me.Label46.Text = "※入出金取得設定の場合、自社・支店を割り付ける必要があります。"
            Me.Label46.UseCompatibleTextRendering = True
            '
            'Label32
            '
            Me.Label32.Location = New System.Drawing.Point(15, 10)
            Me.Label32.Name = "Label32"
            Me.Label32.Size = New System.Drawing.Size(840, 20)
            Me.Label32.TabIndex = 99
            Me.Label32.Text = "賃貸革命V7のデータフォーマットを10で用意されているデータフォーマットへ割り付けて下さい。"
            Me.Label32.UseCompatibleTextRendering = True
            '
            'dgvFBInfo
            '
            Me.dgvFBInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvFBInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFBInfo0, Me.ColFBInfo1, Me.ColFBInfo2, Me.ColFBInfo3, Me.ColFBInfo4, Me.ColFBInfo5, Me.ColFBInfo6, Me.ColFBInfo7, Me.ColFBInfo8, Me.ColFBInfo9, Me.ColFBInfo10, Me.ColFBInfo11, Me.ColFBInfo12, Me.ColFBInfo13})
            Me.dgvFBInfo.Location = New System.Drawing.Point(10, 73)
            Me.dgvFBInfo.Name = "dgvFBInfo"
            Me.dgvFBInfo.RowTemplate.Height = 21
            Me.dgvFBInfo.Size = New System.Drawing.Size(865, 352)
            Me.dgvFBInfo.TabIndex = 96
            '
            'ColFBInfo0
            '
            Me.ColFBInfo0.HeaderText = "フォーマット種別"
            Me.ColFBInfo0.Name = "ColFBInfo0"
            Me.ColFBInfo0.Width = 150
            '
            'ColFBInfo1
            '
            Me.ColFBInfo1.HeaderText = "各口座No"
            Me.ColFBInfo1.Name = "ColFBInfo1"
            Me.ColFBInfo1.Width = 150
            '
            'ColFBInfo2
            '
            Me.ColFBInfo2.HeaderText = "各口座名称"
            Me.ColFBInfo2.Name = "ColFBInfo2"
            Me.ColFBInfo2.Width = 150
            '
            'ColFBInfo3
            '
            Me.ColFBInfo3.HeaderText = "金融機関No"
            Me.ColFBInfo3.Name = "ColFBInfo3"
            Me.ColFBInfo3.Width = 150
            '
            'ColFBInfo4
            '
            Me.ColFBInfo4.HeaderText = "金融機関名"
            Me.ColFBInfo4.Name = "ColFBInfo4"
            Me.ColFBInfo4.Width = 150
            '
            'ColFBInfo5
            '
            Me.ColFBInfo5.HeaderText = "支店No"
            Me.ColFBInfo5.Name = "ColFBInfo5"
            Me.ColFBInfo5.Width = 150
            '
            'ColFBInfo6
            '
            Me.ColFBInfo6.HeaderText = "支店名"
            Me.ColFBInfo6.Name = "ColFBInfo6"
            Me.ColFBInfo6.Width = 150
            '
            'ColFBInfo7
            '
            Me.ColFBInfo7.HeaderText = "口座種別"
            Me.ColFBInfo7.Name = "ColFBInfo7"
            Me.ColFBInfo7.Width = 150
            '
            'ColFBInfo8
            '
            Me.ColFBInfo8.HeaderText = "口座番号"
            Me.ColFBInfo8.Name = "ColFBInfo8"
            Me.ColFBInfo8.Width = 150
            '
            'ColFBInfo9
            '
            Me.ColFBInfo9.HeaderText = "口座名義"
            Me.ColFBInfo9.Name = "ColFBInfo9"
            Me.ColFBInfo9.Width = 150
            '
            'ColFBInfo10
            '
            Me.ColFBInfo10.HeaderText = "フォーマットNo"
            Me.ColFBInfo10.Name = "ColFBInfo10"
            Me.ColFBInfo10.Width = 150
            '
            'ColFBInfo11
            '
            Me.ColFBInfo11.HeaderText = "フォーマット名"
            Me.ColFBInfo11.Name = "ColFBInfo11"
            Me.ColFBInfo11.Width = 150
            '
            'ColFBInfo12
            '
            Me.ColFBInfo12.HeaderText = "自社No"
            Me.ColFBInfo12.Name = "ColFBInfo12"
            Me.ColFBInfo12.Width = 150
            '
            'ColFBInfo13
            '
            Me.ColFBInfo13.HeaderText = "自社名"
            Me.ColFBInfo13.Name = "ColFBInfo13"
            Me.ColFBInfo13.Width = 150
            '
            'tabPage22
            '
            Me.tabPage22.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage22.Controls.Add(Me.Label29)
            Me.tabPage22.Controls.Add(Me.dgvKyrui)
            Me.tabPage22.Location = New System.Drawing.Point(4, 27)
            Me.tabPage22.Name = "tabPage22"
            Me.tabPage22.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage22.Size = New System.Drawing.Size(885, 431)
            Me.tabPage22.TabIndex = 19
            Me.tabPage22.Text = "契約分類マスタ"
            '
            'Label29
            '
            Me.Label29.Location = New System.Drawing.Point(15, 10)
            Me.Label29.Name = "Label29"
            Me.Label29.Size = New System.Drawing.Size(840, 55)
            Me.Label29.TabIndex = 99
            Me.Label29.Text = "「契約分類」の紐付設定を行って下さい。([契約分類No]の最大設定値=9999)"
            Me.Label29.UseCompatibleTextRendering = True
            '
            'dgvKyrui
            '
            Me.dgvKyrui.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvKyrui.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColKyrui0, Me.ColKyrui1, Me.ColKyrui2})
            Me.dgvKyrui.Location = New System.Drawing.Point(10, 73)
            Me.dgvKyrui.Name = "dgvKyrui"
            Me.dgvKyrui.RowTemplate.Height = 21
            Me.dgvKyrui.Size = New System.Drawing.Size(865, 352)
            Me.dgvKyrui.TabIndex = 97
            '
            'ColKyrui0
            '
            Me.ColKyrui0.HeaderText = "契約分類名称"
            Me.ColKyrui0.Name = "ColKyrui0"
            Me.ColKyrui0.Width = 200
            '
            'ColKyrui1
            '
            Me.ColKyrui1.HeaderText = "契約分類No"
            Me.ColKyrui1.Name = "ColKyrui1"
            Me.ColKyrui1.Width = 200
            '
            'ColKyrui2
            '
            Me.ColKyrui2.HeaderText = "契約分類名称"
            Me.ColKyrui2.Name = "ColKyrui2"
            Me.ColKyrui2.Width = 200
            '
            'tabPage23
            '
            Me.tabPage23.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage23.Controls.Add(Me.Label33)
            Me.tabPage23.Controls.Add(Me.dgvTosiYoto)
            Me.tabPage23.Location = New System.Drawing.Point(4, 27)
            Me.tabPage23.Name = "tabPage23"
            Me.tabPage23.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage23.Size = New System.Drawing.Size(885, 431)
            Me.tabPage23.TabIndex = 20
            Me.tabPage23.Text = "都市計画・用途地域マスタ"
            '
            'Label33
            '
            Me.Label33.Location = New System.Drawing.Point(15, 10)
            Me.Label33.Name = "Label33"
            Me.Label33.Size = New System.Drawing.Size(840, 55)
            Me.Label33.TabIndex = 100
            Me.Label33.Text = "「都市計画」「用途地域」の紐付設定を行って下さい。"
            Me.Label33.UseCompatibleTextRendering = True
            '
            'dgvTosiYoto
            '
            Me.dgvTosiYoto.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvTosiYoto.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTosiYoto0, Me.ColTosiYoto1, Me.ColTosiYoto2, Me.ColTosiYoto3})
            Me.dgvTosiYoto.Location = New System.Drawing.Point(10, 73)
            Me.dgvTosiYoto.Name = "dgvTosiYoto"
            Me.dgvTosiYoto.RowTemplate.Height = 21
            Me.dgvTosiYoto.Size = New System.Drawing.Size(865, 352)
            Me.dgvTosiYoto.TabIndex = 98
            '
            'ColTosiYoto0
            '
            Me.ColTosiYoto0.HeaderText = "用途地域・都市計画名"
            Me.ColTosiYoto0.Name = "ColTosiYoto0"
            Me.ColTosiYoto0.Width = 200
            '
            'ColTosiYoto1
            '
            Me.ColTosiYoto1.HeaderText = "用途地域・都市計画区分"
            Me.ColTosiYoto1.Name = "ColTosiYoto1"
            Me.ColTosiYoto1.Width = 200
            '
            'ColTosiYoto2
            '
            Me.ColTosiYoto2.HeaderText = "用途地域・都市計画No"
            Me.ColTosiYoto2.Name = "ColTosiYoto2"
            Me.ColTosiYoto2.Width = 200
            '
            'ColTosiYoto3
            '
            Me.ColTosiYoto3.HeaderText = "用途地域・都市計画名"
            Me.ColTosiYoto3.Name = "ColTosiYoto3"
            Me.ColTosiYoto3.Width = 200
            '
            'tabPage4
            '
            Me.tabPage4.BackColor = System.Drawing.SystemColors.Menu
            Me.tabPage4.Controls.Add(Me.lblFinaltxt)
            Me.tabPage4.Location = New System.Drawing.Point(4, 27)
            Me.tabPage4.Name = "tabPage4"
            Me.tabPage4.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage4.Size = New System.Drawing.Size(912, 549)
            Me.tabPage4.TabIndex = 11
            Me.tabPage4.Text = "終了"
            '
            'lblFinaltxt
            '
            Me.lblFinaltxt.Location = New System.Drawing.Point(32, 48)
            Me.lblFinaltxt.Name = "lblFinaltxt"
            Me.lblFinaltxt.Size = New System.Drawing.Size(133, 18)
            Me.lblFinaltxt.TabIndex = 93
            Me.lblFinaltxt.Text = "紐付けが完了しました。"
            Me.lblFinaltxt.UseCompatibleTextRendering = True
            '
            'lblRelItemRead
            '
            Me.lblRelItemRead.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelItemRead.Location = New System.Drawing.Point(400, 645)
            Me.lblRelItemRead.Name = "lblRelItemRead"
            Me.lblRelItemRead.Size = New System.Drawing.Size(350, 15)
            Me.lblRelItemRead.TabIndex = 125
            Me.lblRelItemRead.Text = "紐付項目を読込中です…"
            Me.lblRelItemRead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.lblRelItemRead.UseCompatibleTextRendering = True
            '
            'pgbRelItemRead
            '
            Me.pgbRelItemRead.Location = New System.Drawing.Point(30, 640)
            Me.pgbRelItemRead.Name = "pgbRelItemRead"
            Me.pgbRelItemRead.Size = New System.Drawing.Size(350, 20)
            Me.pgbRelItemRead.TabIndex = 124
            '
            'btnBack
            '
            Me.btnBack.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnBack.Font = New System.Drawing.Font("メイリオ", 11.25!)
            Me.btnBack.Image = CType(resources.GetObject("btnBack.Image"), System.Drawing.Image)
            Me.btnBack.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnBack.Location = New System.Drawing.Point(817, 640)
            Me.btnBack.Name = "btnBack"
            Me.btnBack.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
            Me.btnBack.Size = New System.Drawing.Size(120, 30)
            Me.btnBack.TabIndex = 129
            Me.btnBack.Text = "   戻  る"
            Me.btnBack.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
            Me.btnBack.UseVisualStyleBackColor = True
            '
            'btnNext
            '
            Me.btnNext.Font = New System.Drawing.Font("メイリオ", 11.25!)
            Me.btnNext.Image = CType(resources.GetObject("btnNext.Image"), System.Drawing.Image)
            Me.btnNext.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnNext.Location = New System.Drawing.Point(952, 640)
            Me.btnNext.Name = "btnNext"
            Me.btnNext.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
            Me.btnNext.Size = New System.Drawing.Size(120, 30)
            Me.btnNext.TabIndex = 128
            Me.btnNext.Text = "   次  へ"
            Me.btnNext.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
            Me.btnNext.UseVisualStyleBackColor = True
            '
            'btnEnd
            '
            Me.btnEnd.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnEnd.Font = New System.Drawing.Font("メイリオ", 11.25!)
            Me.btnEnd.Image = CType(resources.GetObject("btnEnd.Image"), System.Drawing.Image)
            Me.btnEnd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnEnd.Location = New System.Drawing.Point(1089, 640)
            Me.btnEnd.Name = "btnEnd"
            Me.btnEnd.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
            Me.btnEnd.Size = New System.Drawing.Size(120, 30)
            Me.btnEnd.TabIndex = 127
            Me.btnEnd.Text = "   終  了"
            Me.btnEnd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
            Me.btnEnd.UseVisualStyleBackColor = True
            '
            'DataGridView1
            '
            Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.DataGridView1.Location = New System.Drawing.Point(956, 512)
            Me.DataGridView1.Name = "DataGridView1"
            Me.DataGridView1.RowTemplate.Height = 21
            Me.DataGridView1.Size = New System.Drawing.Size(278, 104)
            Me.DataGridView1.TabIndex = 130
            '
            'Label16
            '
            Me.Label16.BackColor = System.Drawing.Color.PowderBlue
            Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.Label16.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label16.Location = New System.Drawing.Point(25, 28)
            Me.Label16.Name = "Label16"
            Me.Label16.Size = New System.Drawing.Size(20, 15)
            Me.Label16.TabIndex = 131
            Me.Label16.UseCompatibleTextRendering = True
            '
            'Label19
            '
            Me.Label19.BackColor = System.Drawing.Color.Bisque
            Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.Label19.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label19.Location = New System.Drawing.Point(25, 46)
            Me.Label19.Name = "Label19"
            Me.Label19.Size = New System.Drawing.Size(20, 15)
            Me.Label19.TabIndex = 132
            Me.Label19.UseCompatibleTextRendering = True
            '
            'Label20
            '
            Me.Label20.BackColor = System.Drawing.Color.LightGray
            Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.Label20.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label20.Location = New System.Drawing.Point(25, 64)
            Me.Label20.Name = "Label20"
            Me.Label20.Size = New System.Drawing.Size(20, 15)
            Me.Label20.TabIndex = 133
            Me.Label20.UseCompatibleTextRendering = True
            '
            'Label21
            '
            Me.Label21.BackColor = System.Drawing.Color.Pink
            Me.Label21.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.Label21.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label21.Location = New System.Drawing.Point(25, 82)
            Me.Label21.Name = "Label21"
            Me.Label21.Size = New System.Drawing.Size(20, 15)
            Me.Label21.TabIndex = 134
            Me.Label21.UseCompatibleTextRendering = True
            '
            'Label22
            '
            Me.Label22.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label22.Location = New System.Drawing.Point(60, 28)
            Me.Label22.Name = "Label22"
            Me.Label22.Size = New System.Drawing.Size(190, 15)
            Me.Label22.TabIndex = 135
            Me.Label22.Text = "コンバート元データ(賃貸革命V7)"
            Me.Label22.UseCompatibleTextRendering = True
            '
            'Label23
            '
            Me.Label23.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label23.Location = New System.Drawing.Point(60, 46)
            Me.Label23.Name = "Label23"
            Me.Label23.Size = New System.Drawing.Size(190, 15)
            Me.Label23.TabIndex = 136
            Me.Label23.Text = "コンバート先データ(賃貸革命10)"
            Me.Label23.UseCompatibleTextRendering = True
            '
            'Label24
            '
            Me.Label24.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label24.Location = New System.Drawing.Point(60, 64)
            Me.Label24.Name = "Label24"
            Me.Label24.Size = New System.Drawing.Size(190, 15)
            Me.Label24.TabIndex = 137
            Me.Label24.Text = "設定不可"
            Me.Label24.UseCompatibleTextRendering = True
            '
            'Label25
            '
            Me.Label25.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label25.Location = New System.Drawing.Point(60, 82)
            Me.Label25.Name = "Label25"
            Me.Label25.Size = New System.Drawing.Size(190, 15)
            Me.Label25.TabIndex = 138
            Me.Label25.Text = "選択行"
            Me.Label25.UseCompatibleTextRendering = True
            '
            'Label26
            '
            Me.Label26.BackColor = System.Drawing.Color.NavajoWhite
            Me.Label26.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.Label26.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label26.Location = New System.Drawing.Point(25, 100)
            Me.Label26.Name = "Label26"
            Me.Label26.Size = New System.Drawing.Size(20, 15)
            Me.Label26.TabIndex = 139
            Me.Label26.UseCompatibleTextRendering = True
            '
            'Label27
            '
            Me.Label27.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label27.Location = New System.Drawing.Point(60, 100)
            Me.Label27.Name = "Label27"
            Me.Label27.Size = New System.Drawing.Size(190, 15)
            Me.Label27.TabIndex = 140
            Me.Label27.Text = "一括設定行(入金項目のみ)"
            Me.Label27.UseCompatibleTextRendering = True
            '
            'Label28
            '
            Me.Label28.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.Label28.Location = New System.Drawing.Point(10, 5)
            Me.Label28.Name = "Label28"
            Me.Label28.Size = New System.Drawing.Size(200, 18)
            Me.Label28.TabIndex = 141
            Me.Label28.Text = "※配色に関して"
            Me.Label28.UseCompatibleTextRendering = True
            '
            'lbltest
            '
            Me.lbltest.BackColor = System.Drawing.Color.DodgerBlue
            Me.lbltest.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lbltest.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lbltest.Location = New System.Drawing.Point(1184, 6)
            Me.lbltest.Name = "lbltest"
            Me.lbltest.Size = New System.Drawing.Size(24, 18)
            Me.lbltest.TabIndex = 142
            Me.lbltest.UseCompatibleTextRendering = True
            Me.lbltest.Visible = False
            '
            'lblHidden3
            '
            Me.lblHidden3.Location = New System.Drawing.Point(6, 11)
            Me.lblHidden3.Name = "lblHidden3"
            Me.lblHidden3.Size = New System.Drawing.Size(18, 55)
            Me.lblHidden3.TabIndex = 143
            Me.lblHidden3.Text = "　"
            '
            'pnlHaisyoku
            '
            Me.pnlHaisyoku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.pnlHaisyoku.Controls.Add(Me.Label28)
            Me.pnlHaisyoku.Controls.Add(Me.Label16)
            Me.pnlHaisyoku.Controls.Add(Me.Label19)
            Me.pnlHaisyoku.Controls.Add(Me.Label20)
            Me.pnlHaisyoku.Controls.Add(Me.Label27)
            Me.pnlHaisyoku.Controls.Add(Me.Label21)
            Me.pnlHaisyoku.Controls.Add(Me.Label26)
            Me.pnlHaisyoku.Controls.Add(Me.Label22)
            Me.pnlHaisyoku.Controls.Add(Me.Label25)
            Me.pnlHaisyoku.Controls.Add(Me.Label23)
            Me.pnlHaisyoku.Controls.Add(Me.Label24)
            Me.pnlHaisyoku.Location = New System.Drawing.Point(956, 63)
            Me.pnlHaisyoku.Name = "pnlHaisyoku"
            Me.pnlHaisyoku.Size = New System.Drawing.Size(278, 128)
            Me.pnlHaisyoku.TabIndex = 144
            '
            'pnlRelSelectPanel
            '
            Me.pnlRelSelectPanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllTosiYoto)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiTosiYoto)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllKyBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblTitleRelTotalCnt)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllFbFmt)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiKyBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllNkKbn)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllGazo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblTitleRelCnt)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllTaiyo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiNkKbn)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllJisyaKoza)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiFbFmt)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllNkin)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblTitleRelKomk)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiTaiyo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllKozo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectTosiYoto)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllSetubi)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblSumiZanGazo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllKozaSyubetu)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectJisyaKoza)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiNkin)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectKyBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiJisyaKoza)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectGazo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiKozo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectFBFmt)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiKozaSyubetu)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectNkin)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiSetubi)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectHySetubi)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectKozo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectKozaSyubetu)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllBkBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblTitleRelSelect)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectBkBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiBkBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectTaiyo)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectHyBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblRelSelectNkKbn)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntTosiYotoBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntKyBunruiBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSumiHyBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntFbFmtBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntGazoBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntAllHyBunrui)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntNkKbnBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntTaiyoBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntNkinBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntJisyaKozaBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntSetubiBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntKozaSyubetuBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntKozoBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntHyBunruiBack)
            Me.pnlRelSelectPanel.Controls.Add(Me.lblCntBkBunruiBack)
            Me.pnlRelSelectPanel.Location = New System.Drawing.Point(956, 197)
            Me.pnlRelSelectPanel.Name = "pnlRelSelectPanel"
            Me.pnlRelSelectPanel.Size = New System.Drawing.Size(278, 307)
            Me.pnlRelSelectPanel.TabIndex = 145
            '
            'lblCntAllTosiYoto
            '
            Me.lblCntAllTosiYoto.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllTosiYoto.Location = New System.Drawing.Point(215, 278)
            Me.lblCntAllTosiYoto.Name = "lblCntAllTosiYoto"
            Me.lblCntAllTosiYoto.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllTosiYoto.TabIndex = 218
            Me.lblCntAllTosiYoto.Text = "9,999"
            Me.lblCntAllTosiYoto.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSumiTosiYoto
            '
            Me.lblCntSumiTosiYoto.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiTosiYoto.Location = New System.Drawing.Point(156, 278)
            Me.lblCntSumiTosiYoto.Name = "lblCntSumiTosiYoto"
            Me.lblCntSumiTosiYoto.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiTosiYoto.TabIndex = 217
            Me.lblCntSumiTosiYoto.Text = "9,999"
            Me.lblCntSumiTosiYoto.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllKyBunrui
            '
            Me.lblCntAllKyBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllKyBunrui.Location = New System.Drawing.Point(215, 259)
            Me.lblCntAllKyBunrui.Name = "lblCntAllKyBunrui"
            Me.lblCntAllKyBunrui.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllKyBunrui.TabIndex = 213
            Me.lblCntAllKyBunrui.Text = "9,999"
            Me.lblCntAllKyBunrui.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblTitleRelTotalCnt
            '
            Me.lblTitleRelTotalCnt.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblTitleRelTotalCnt.Location = New System.Drawing.Point(203, 28)
            Me.lblTitleRelTotalCnt.Name = "lblTitleRelTotalCnt"
            Me.lblTitleRelTotalCnt.Size = New System.Drawing.Size(60, 15)
            Me.lblTitleRelTotalCnt.TabIndex = 224
            Me.lblTitleRelTotalCnt.Text = "紐付総数"
            Me.lblTitleRelTotalCnt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllFbFmt
            '
            Me.lblCntAllFbFmt.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllFbFmt.Location = New System.Drawing.Point(215, 241)
            Me.lblCntAllFbFmt.Name = "lblCntAllFbFmt"
            Me.lblCntAllFbFmt.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllFbFmt.TabIndex = 208
            Me.lblCntAllFbFmt.Text = "9,999"
            Me.lblCntAllFbFmt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSumiKyBunrui
            '
            Me.lblCntSumiKyBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiKyBunrui.Location = New System.Drawing.Point(156, 259)
            Me.lblCntSumiKyBunrui.Name = "lblCntSumiKyBunrui"
            Me.lblCntSumiKyBunrui.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiKyBunrui.TabIndex = 212
            Me.lblCntSumiKyBunrui.Text = "9,999"
            Me.lblCntSumiKyBunrui.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllNkKbn
            '
            Me.lblCntAllNkKbn.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllNkKbn.Location = New System.Drawing.Point(215, 90)
            Me.lblCntAllNkKbn.Name = "lblCntAllNkKbn"
            Me.lblCntAllNkKbn.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllNkKbn.TabIndex = 163
            Me.lblCntAllNkKbn.Text = "9,999"
            Me.lblCntAllNkKbn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllGazo
            '
            Me.lblCntAllGazo.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllGazo.Location = New System.Drawing.Point(215, 222)
            Me.lblCntAllGazo.Name = "lblCntAllGazo"
            Me.lblCntAllGazo.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllGazo.TabIndex = 203
            Me.lblCntAllGazo.Text = "9,999"
            Me.lblCntAllGazo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblTitleRelCnt
            '
            Me.lblTitleRelCnt.AccessibleRole = System.Windows.Forms.AccessibleRole.Outline
            Me.lblTitleRelCnt.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblTitleRelCnt.Location = New System.Drawing.Point(141, 28)
            Me.lblTitleRelCnt.Name = "lblTitleRelCnt"
            Me.lblTitleRelCnt.Size = New System.Drawing.Size(60, 17)
            Me.lblTitleRelCnt.TabIndex = 223
            Me.lblTitleRelCnt.Text = "紐付済"
            Me.lblTitleRelCnt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllTaiyo
            '
            Me.lblCntAllTaiyo.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllTaiyo.Location = New System.Drawing.Point(215, 109)
            Me.lblCntAllTaiyo.Name = "lblCntAllTaiyo"
            Me.lblCntAllTaiyo.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllTaiyo.TabIndex = 168
            Me.lblCntAllTaiyo.Text = "9,999"
            Me.lblCntAllTaiyo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSumiNkKbn
            '
            Me.lblCntSumiNkKbn.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiNkKbn.Location = New System.Drawing.Point(156, 90)
            Me.lblCntSumiNkKbn.Name = "lblCntSumiNkKbn"
            Me.lblCntSumiNkKbn.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiNkKbn.TabIndex = 162
            Me.lblCntSumiNkKbn.Text = "9,999"
            Me.lblCntSumiNkKbn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllJisyaKoza
            '
            Me.lblCntAllJisyaKoza.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllJisyaKoza.Location = New System.Drawing.Point(215, 203)
            Me.lblCntAllJisyaKoza.Name = "lblCntAllJisyaKoza"
            Me.lblCntAllJisyaKoza.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllJisyaKoza.TabIndex = 198
            Me.lblCntAllJisyaKoza.Text = "9,999"
            Me.lblCntAllJisyaKoza.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSumiFbFmt
            '
            Me.lblCntSumiFbFmt.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiFbFmt.Location = New System.Drawing.Point(156, 241)
            Me.lblCntSumiFbFmt.Name = "lblCntSumiFbFmt"
            Me.lblCntSumiFbFmt.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiFbFmt.TabIndex = 207
            Me.lblCntSumiFbFmt.Text = "9,999"
            Me.lblCntSumiFbFmt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllNkin
            '
            Me.lblCntAllNkin.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllNkin.Location = New System.Drawing.Point(215, 128)
            Me.lblCntAllNkin.Name = "lblCntAllNkin"
            Me.lblCntAllNkin.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllNkin.TabIndex = 173
            Me.lblCntAllNkin.Text = "9,999"
            Me.lblCntAllNkin.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblTitleRelKomk
            '
            Me.lblTitleRelKomk.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblTitleRelKomk.Location = New System.Drawing.Point(12, 28)
            Me.lblTitleRelKomk.Name = "lblTitleRelKomk"
            Me.lblTitleRelKomk.Size = New System.Drawing.Size(123, 15)
            Me.lblTitleRelKomk.TabIndex = 222
            Me.lblTitleRelKomk.Text = "紐付項目内容"
            Me.lblTitleRelKomk.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSumiTaiyo
            '
            Me.lblCntSumiTaiyo.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiTaiyo.Location = New System.Drawing.Point(156, 109)
            Me.lblCntSumiTaiyo.Name = "lblCntSumiTaiyo"
            Me.lblCntSumiTaiyo.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiTaiyo.TabIndex = 167
            Me.lblCntSumiTaiyo.Text = "9,999"
            Me.lblCntSumiTaiyo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllKozo
            '
            Me.lblCntAllKozo.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllKozo.Location = New System.Drawing.Point(215, 147)
            Me.lblCntAllKozo.Name = "lblCntAllKozo"
            Me.lblCntAllKozo.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllKozo.TabIndex = 178
            Me.lblCntAllKozo.Text = "9,999"
            Me.lblCntAllKozo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectTosiYoto
            '
            Me.lblRelSelectTosiYoto.BackColor = System.Drawing.Color.White
            Me.lblRelSelectTosiYoto.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectTosiYoto.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectTosiYoto.Location = New System.Drawing.Point(15, 276)
            Me.lblRelSelectTosiYoto.Name = "lblRelSelectTosiYoto"
            Me.lblRelSelectTosiYoto.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectTosiYoto.TabIndex = 151
            Me.lblRelSelectTosiYoto.Text = "都市計画・用途地域"
            Me.lblRelSelectTosiYoto.UseCompatibleTextRendering = True
            '
            'lblCntAllSetubi
            '
            Me.lblCntAllSetubi.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllSetubi.Location = New System.Drawing.Point(215, 185)
            Me.lblCntAllSetubi.Name = "lblCntAllSetubi"
            Me.lblCntAllSetubi.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllSetubi.TabIndex = 188
            Me.lblCntAllSetubi.Text = "9,999"
            Me.lblCntAllSetubi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblSumiZanGazo
            '
            Me.lblSumiZanGazo.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblSumiZanGazo.Location = New System.Drawing.Point(156, 222)
            Me.lblSumiZanGazo.Name = "lblSumiZanGazo"
            Me.lblSumiZanGazo.Size = New System.Drawing.Size(40, 15)
            Me.lblSumiZanGazo.TabIndex = 202
            Me.lblSumiZanGazo.Text = "9,999"
            Me.lblSumiZanGazo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllKozaSyubetu
            '
            Me.lblCntAllKozaSyubetu.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllKozaSyubetu.Location = New System.Drawing.Point(215, 166)
            Me.lblCntAllKozaSyubetu.Name = "lblCntAllKozaSyubetu"
            Me.lblCntAllKozaSyubetu.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllKozaSyubetu.TabIndex = 183
            Me.lblCntAllKozaSyubetu.Text = "9,999"
            Me.lblCntAllKozaSyubetu.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectJisyaKoza
            '
            Me.lblRelSelectJisyaKoza.BackColor = System.Drawing.Color.White
            Me.lblRelSelectJisyaKoza.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectJisyaKoza.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectJisyaKoza.Location = New System.Drawing.Point(15, 200)
            Me.lblRelSelectJisyaKoza.Name = "lblRelSelectJisyaKoza"
            Me.lblRelSelectJisyaKoza.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectJisyaKoza.TabIndex = 147
            Me.lblRelSelectJisyaKoza.Text = "自社口座マスタ"
            Me.lblRelSelectJisyaKoza.UseCompatibleTextRendering = True
            '
            'lblCntSumiNkin
            '
            Me.lblCntSumiNkin.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiNkin.Location = New System.Drawing.Point(156, 128)
            Me.lblCntSumiNkin.Name = "lblCntSumiNkin"
            Me.lblCntSumiNkin.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiNkin.TabIndex = 172
            Me.lblCntSumiNkin.Text = "9,999"
            Me.lblCntSumiNkin.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectKyBunrui
            '
            Me.lblRelSelectKyBunrui.BackColor = System.Drawing.Color.White
            Me.lblRelSelectKyBunrui.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectKyBunrui.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectKyBunrui.Location = New System.Drawing.Point(15, 257)
            Me.lblRelSelectKyBunrui.Name = "lblRelSelectKyBunrui"
            Me.lblRelSelectKyBunrui.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectKyBunrui.TabIndex = 150
            Me.lblRelSelectKyBunrui.Text = "契約分類マスタ"
            Me.lblRelSelectKyBunrui.UseCompatibleTextRendering = True
            '
            'lblCntSumiJisyaKoza
            '
            Me.lblCntSumiJisyaKoza.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiJisyaKoza.Location = New System.Drawing.Point(156, 203)
            Me.lblCntSumiJisyaKoza.Name = "lblCntSumiJisyaKoza"
            Me.lblCntSumiJisyaKoza.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiJisyaKoza.TabIndex = 197
            Me.lblCntSumiJisyaKoza.Text = "9,999"
            Me.lblCntSumiJisyaKoza.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectGazo
            '
            Me.lblRelSelectGazo.BackColor = System.Drawing.Color.White
            Me.lblRelSelectGazo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectGazo.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectGazo.Location = New System.Drawing.Point(15, 219)
            Me.lblRelSelectGazo.Name = "lblRelSelectGazo"
            Me.lblRelSelectGazo.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectGazo.TabIndex = 148
            Me.lblRelSelectGazo.Text = "周辺画像"
            Me.lblRelSelectGazo.UseCompatibleTextRendering = True
            '
            'lblCntSumiKozo
            '
            Me.lblCntSumiKozo.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiKozo.Location = New System.Drawing.Point(156, 147)
            Me.lblCntSumiKozo.Name = "lblCntSumiKozo"
            Me.lblCntSumiKozo.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiKozo.TabIndex = 177
            Me.lblCntSumiKozo.Text = "9,999"
            Me.lblCntSumiKozo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectFBFmt
            '
            Me.lblRelSelectFBFmt.BackColor = System.Drawing.Color.White
            Me.lblRelSelectFBFmt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectFBFmt.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectFBFmt.Location = New System.Drawing.Point(15, 238)
            Me.lblRelSelectFBFmt.Name = "lblRelSelectFBFmt"
            Me.lblRelSelectFBFmt.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectFBFmt.TabIndex = 149
            Me.lblRelSelectFBFmt.Text = "FBフォーマット名"
            Me.lblRelSelectFBFmt.UseCompatibleTextRendering = True
            '
            'lblCntSumiKozaSyubetu
            '
            Me.lblCntSumiKozaSyubetu.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiKozaSyubetu.Location = New System.Drawing.Point(156, 166)
            Me.lblCntSumiKozaSyubetu.Name = "lblCntSumiKozaSyubetu"
            Me.lblCntSumiKozaSyubetu.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiKozaSyubetu.TabIndex = 182
            Me.lblCntSumiKozaSyubetu.Text = "9,999"
            Me.lblCntSumiKozaSyubetu.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectNkin
            '
            Me.lblRelSelectNkin.BackColor = System.Drawing.Color.White
            Me.lblRelSelectNkin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectNkin.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectNkin.Location = New System.Drawing.Point(15, 124)
            Me.lblRelSelectNkin.Name = "lblRelSelectNkin"
            Me.lblRelSelectNkin.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectNkin.TabIndex = 142
            Me.lblRelSelectNkin.Text = "入金項目マスタ"
            Me.lblRelSelectNkin.UseCompatibleTextRendering = True
            '
            'lblCntSumiSetubi
            '
            Me.lblCntSumiSetubi.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiSetubi.Location = New System.Drawing.Point(156, 185)
            Me.lblCntSumiSetubi.Name = "lblCntSumiSetubi"
            Me.lblCntSumiSetubi.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiSetubi.TabIndex = 187
            Me.lblCntSumiSetubi.Text = "9,999"
            Me.lblCntSumiSetubi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectHySetubi
            '
            Me.lblRelSelectHySetubi.BackColor = System.Drawing.Color.White
            Me.lblRelSelectHySetubi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectHySetubi.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectHySetubi.Location = New System.Drawing.Point(15, 181)
            Me.lblRelSelectHySetubi.Name = "lblRelSelectHySetubi"
            Me.lblRelSelectHySetubi.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectHySetubi.TabIndex = 145
            Me.lblRelSelectHySetubi.Text = "部屋設備"
            Me.lblRelSelectHySetubi.UseCompatibleTextRendering = True
            '
            'lblRelSelectKozo
            '
            Me.lblRelSelectKozo.BackColor = System.Drawing.Color.White
            Me.lblRelSelectKozo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectKozo.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectKozo.Location = New System.Drawing.Point(15, 143)
            Me.lblRelSelectKozo.Name = "lblRelSelectKozo"
            Me.lblRelSelectKozo.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectKozo.TabIndex = 143
            Me.lblRelSelectKozo.Text = "物件構造マスタ"
            Me.lblRelSelectKozo.UseCompatibleTextRendering = True
            '
            'lblRelSelectKozaSyubetu
            '
            Me.lblRelSelectKozaSyubetu.BackColor = System.Drawing.Color.White
            Me.lblRelSelectKozaSyubetu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectKozaSyubetu.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectKozaSyubetu.Location = New System.Drawing.Point(15, 162)
            Me.lblRelSelectKozaSyubetu.Name = "lblRelSelectKozaSyubetu"
            Me.lblRelSelectKozaSyubetu.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectKozaSyubetu.TabIndex = 144
            Me.lblRelSelectKozaSyubetu.Text = "口座種別マスタ"
            Me.lblRelSelectKozaSyubetu.UseCompatibleTextRendering = True
            '
            'lblCntAllBkBunrui
            '
            Me.lblCntAllBkBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllBkBunrui.Location = New System.Drawing.Point(215, 51)
            Me.lblCntAllBkBunrui.Name = "lblCntAllBkBunrui"
            Me.lblCntAllBkBunrui.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllBkBunrui.TabIndex = 153
            Me.lblCntAllBkBunrui.Text = "9,999"
            Me.lblCntAllBkBunrui.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblTitleRelSelect
            '
            Me.lblTitleRelSelect.Font = New System.Drawing.Font("メイリオ", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblTitleRelSelect.Location = New System.Drawing.Point(10, 5)
            Me.lblTitleRelSelect.Name = "lblTitleRelSelect"
            Me.lblTitleRelSelect.Size = New System.Drawing.Size(248, 18)
            Me.lblTitleRelSelect.TabIndex = 141
            Me.lblTitleRelSelect.Text = "紐付設定する項目を選択して下さい"
            Me.lblTitleRelSelect.UseCompatibleTextRendering = True
            '
            'lblRelSelectBkBunrui
            '
            Me.lblRelSelectBkBunrui.BackColor = System.Drawing.Color.White
            Me.lblRelSelectBkBunrui.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectBkBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectBkBunrui.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.lblRelSelectBkBunrui.Location = New System.Drawing.Point(15, 48)
            Me.lblRelSelectBkBunrui.Name = "lblRelSelectBkBunrui"
            Me.lblRelSelectBkBunrui.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectBkBunrui.TabIndex = 135
            Me.lblRelSelectBkBunrui.Text = "物件分類マスタ"
            Me.lblRelSelectBkBunrui.UseCompatibleTextRendering = True
            '
            'lblCntSumiBkBunrui
            '
            Me.lblCntSumiBkBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiBkBunrui.Location = New System.Drawing.Point(156, 51)
            Me.lblCntSumiBkBunrui.Name = "lblCntSumiBkBunrui"
            Me.lblCntSumiBkBunrui.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiBkBunrui.TabIndex = 152
            Me.lblCntSumiBkBunrui.Text = "9,999"
            Me.lblCntSumiBkBunrui.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblRelSelectTaiyo
            '
            Me.lblRelSelectTaiyo.BackColor = System.Drawing.Color.White
            Me.lblRelSelectTaiyo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectTaiyo.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectTaiyo.Location = New System.Drawing.Point(15, 105)
            Me.lblRelSelectTaiyo.Name = "lblRelSelectTaiyo"
            Me.lblRelSelectTaiyo.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectTaiyo.TabIndex = 138
            Me.lblRelSelectTaiyo.Text = "取引態様マスタ"
            Me.lblRelSelectTaiyo.UseCompatibleTextRendering = True
            '
            'lblRelSelectHyBunrui
            '
            Me.lblRelSelectHyBunrui.BackColor = System.Drawing.Color.White
            Me.lblRelSelectHyBunrui.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectHyBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectHyBunrui.Location = New System.Drawing.Point(15, 67)
            Me.lblRelSelectHyBunrui.Name = "lblRelSelectHyBunrui"
            Me.lblRelSelectHyBunrui.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectHyBunrui.TabIndex = 136
            Me.lblRelSelectHyBunrui.Text = "部屋分類マスタ"
            Me.lblRelSelectHyBunrui.UseCompatibleTextRendering = True
            '
            'lblRelSelectNkKbn
            '
            Me.lblRelSelectNkKbn.BackColor = System.Drawing.Color.White
            Me.lblRelSelectNkKbn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.lblRelSelectNkKbn.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.lblRelSelectNkKbn.Location = New System.Drawing.Point(15, 86)
            Me.lblRelSelectNkKbn.Name = "lblRelSelectNkKbn"
            Me.lblRelSelectNkKbn.Size = New System.Drawing.Size(120, 20)
            Me.lblRelSelectNkKbn.TabIndex = 137
            Me.lblRelSelectNkKbn.Text = "入金区分マスタ"
            Me.lblRelSelectNkKbn.UseCompatibleTextRendering = True
            '
            'lblCntTosiYotoBack
            '
            Me.lblCntTosiYotoBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntTosiYotoBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntTosiYotoBack.Location = New System.Drawing.Point(137, 278)
            Me.lblCntTosiYotoBack.Name = "lblCntTosiYotoBack"
            Me.lblCntTosiYotoBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntTosiYotoBack.TabIndex = 238
            Me.lblCntTosiYotoBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntTosiYotoBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntKyBunruiBack
            '
            Me.lblCntKyBunruiBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntKyBunruiBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntKyBunruiBack.Location = New System.Drawing.Point(137, 259)
            Me.lblCntKyBunruiBack.Name = "lblCntKyBunruiBack"
            Me.lblCntKyBunruiBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntKyBunruiBack.TabIndex = 237
            Me.lblCntKyBunruiBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntKyBunruiBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSumiHyBunrui
            '
            Me.lblCntSumiHyBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSumiHyBunrui.Location = New System.Drawing.Point(156, 70)
            Me.lblCntSumiHyBunrui.Name = "lblCntSumiHyBunrui"
            Me.lblCntSumiHyBunrui.Size = New System.Drawing.Size(40, 15)
            Me.lblCntSumiHyBunrui.TabIndex = 157
            Me.lblCntSumiHyBunrui.Text = "9,999"
            Me.lblCntSumiHyBunrui.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntFbFmtBack
            '
            Me.lblCntFbFmtBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntFbFmtBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntFbFmtBack.Location = New System.Drawing.Point(137, 240)
            Me.lblCntFbFmtBack.Name = "lblCntFbFmtBack"
            Me.lblCntFbFmtBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntFbFmtBack.TabIndex = 236
            Me.lblCntFbFmtBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntFbFmtBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntGazoBack
            '
            Me.lblCntGazoBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntGazoBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntGazoBack.Location = New System.Drawing.Point(137, 221)
            Me.lblCntGazoBack.Name = "lblCntGazoBack"
            Me.lblCntGazoBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntGazoBack.TabIndex = 235
            Me.lblCntGazoBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntGazoBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntAllHyBunrui
            '
            Me.lblCntAllHyBunrui.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntAllHyBunrui.Location = New System.Drawing.Point(215, 70)
            Me.lblCntAllHyBunrui.Name = "lblCntAllHyBunrui"
            Me.lblCntAllHyBunrui.Size = New System.Drawing.Size(40, 15)
            Me.lblCntAllHyBunrui.TabIndex = 158
            Me.lblCntAllHyBunrui.Text = "9,999"
            Me.lblCntAllHyBunrui.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntNkKbnBack
            '
            Me.lblCntNkKbnBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntNkKbnBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntNkKbnBack.Location = New System.Drawing.Point(137, 89)
            Me.lblCntNkKbnBack.Name = "lblCntNkKbnBack"
            Me.lblCntNkKbnBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntNkKbnBack.TabIndex = 234
            Me.lblCntNkKbnBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntNkKbnBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntTaiyoBack
            '
            Me.lblCntTaiyoBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntTaiyoBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntTaiyoBack.Location = New System.Drawing.Point(137, 108)
            Me.lblCntTaiyoBack.Name = "lblCntTaiyoBack"
            Me.lblCntTaiyoBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntTaiyoBack.TabIndex = 233
            Me.lblCntTaiyoBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntTaiyoBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntNkinBack
            '
            Me.lblCntNkinBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntNkinBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntNkinBack.Location = New System.Drawing.Point(137, 127)
            Me.lblCntNkinBack.Name = "lblCntNkinBack"
            Me.lblCntNkinBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntNkinBack.TabIndex = 232
            Me.lblCntNkinBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntNkinBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntJisyaKozaBack
            '
            Me.lblCntJisyaKozaBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntJisyaKozaBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntJisyaKozaBack.Location = New System.Drawing.Point(137, 202)
            Me.lblCntJisyaKozaBack.Name = "lblCntJisyaKozaBack"
            Me.lblCntJisyaKozaBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntJisyaKozaBack.TabIndex = 231
            Me.lblCntJisyaKozaBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntJisyaKozaBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntSetubiBack
            '
            Me.lblCntSetubiBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntSetubiBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntSetubiBack.Location = New System.Drawing.Point(137, 184)
            Me.lblCntSetubiBack.Name = "lblCntSetubiBack"
            Me.lblCntSetubiBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntSetubiBack.TabIndex = 229
            Me.lblCntSetubiBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntSetubiBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntKozaSyubetuBack
            '
            Me.lblCntKozaSyubetuBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntKozaSyubetuBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntKozaSyubetuBack.Location = New System.Drawing.Point(137, 165)
            Me.lblCntKozaSyubetuBack.Name = "lblCntKozaSyubetuBack"
            Me.lblCntKozaSyubetuBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntKozaSyubetuBack.TabIndex = 228
            Me.lblCntKozaSyubetuBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntKozaSyubetuBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntKozoBack
            '
            Me.lblCntKozoBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntKozoBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntKozoBack.Location = New System.Drawing.Point(137, 146)
            Me.lblCntKozoBack.Name = "lblCntKozoBack"
            Me.lblCntKozoBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntKozoBack.TabIndex = 227
            Me.lblCntKozoBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntKozoBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntHyBunruiBack
            '
            Me.lblCntHyBunruiBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntHyBunruiBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntHyBunruiBack.Location = New System.Drawing.Point(137, 70)
            Me.lblCntHyBunruiBack.Name = "lblCntHyBunruiBack"
            Me.lblCntHyBunruiBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntHyBunruiBack.TabIndex = 226
            Me.lblCntHyBunruiBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntHyBunruiBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lblCntBkBunruiBack
            '
            Me.lblCntBkBunruiBack.BackColor = System.Drawing.SystemColors.Control
            Me.lblCntBkBunruiBack.Font = New System.Drawing.Font("メイリオ", 9.0!)
            Me.lblCntBkBunruiBack.Location = New System.Drawing.Point(137, 51)
            Me.lblCntBkBunruiBack.Name = "lblCntBkBunruiBack"
            Me.lblCntBkBunruiBack.Size = New System.Drawing.Size(137, 15)
            Me.lblCntBkBunruiBack.TabIndex = 225
            Me.lblCntBkBunruiBack.Text = "(    x,xxx  /  x,xxx    )"
            Me.lblCntBkBunruiBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'RelationFrm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1264, 682)
            Me.Controls.Add(Me.pnlRelSelectPanel)
            Me.Controls.Add(Me.pnlHaisyoku)
            Me.Controls.Add(Me.lblHidden3)
            Me.Controls.Add(Me.lbltest)
            Me.Controls.Add(Me.lblRelItemRead)
            Me.Controls.Add(Me.DataGridView1)
            Me.Controls.Add(Me.pgbRelItemRead)
            Me.Controls.Add(Me.btnBack)
            Me.Controls.Add(Me.btnNext)
            Me.Controls.Add(Me.btnEnd)
            Me.Controls.Add(Me.tabCtrlRelMain)
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.Location = New System.Drawing.Point(10, 73)
            Me.MaximizeBox = False
            Me.Name = "RelationFrm"
            Me.Text = "紐付設定"
            Me.tabCtrlRelMain.ResumeLayout(False)
            Me.tabPage0.ResumeLayout(False)
            Me.tabPage0.PerformLayout()
            Me.grpKagicntgrp.ResumeLayout(False)
            Me.grpDev.ResumeLayout(False)
            Me.grpDev.PerformLayout()
            Me.tabPage1.ResumeLayout(False)
            Me.pnlTimeOut.ResumeLayout(False)
            Me.pnlTimeOut.PerformLayout()
            Me.grp10ConnectInfo.ResumeLayout(False)
            Me.grp10ConnectInfo.PerformLayout()
            Me.grpV10Authent.ResumeLayout(False)
            Me.grpV10Authent.PerformLayout()
            Me.grpV7ConnectInfo.ResumeLayout(False)
            Me.grpV7ConnectInfo.PerformLayout()
            Me.grpV7Authent.ResumeLayout(False)
            Me.grpV7Authent.PerformLayout()
            Me.tabPage2.ResumeLayout(False)
            Me.tabPage2.PerformLayout()
            Me.grpKagi.ResumeLayout(False)
            Me.grpKagi.PerformLayout()
            Me.tabPage3.ResumeLayout(False)
            Me.tabCtrlRelwork.ResumeLayout(False)
            Me.tabPage10.ResumeLayout(False)
            CType(Me.dgvBkrui, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage11.ResumeLayout(False)
            CType(Me.dgvHyrui, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage12.ResumeLayout(False)
            CType(Me.dgvNkinKbn, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage13.ResumeLayout(False)
            CType(Me.dgvToritaiyo, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage14.ResumeLayout(False)
            Me.tabPage14.PerformLayout()
            Me.Panel1.ResumeLayout(False)
            Me.Panel1.PerformLayout()
            CType(Me.dgvNkinkomk, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage15.ResumeLayout(False)
            CType(Me.dgvKozo, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage16.ResumeLayout(False)
            CType(Me.dgvKozasyu, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage17.ResumeLayout(False)
            CType(Me.dgvSetubi, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage18.ResumeLayout(False)
            CType(Me.dgvKagi, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage19.ResumeLayout(False)
            CType(Me.dgvJisyakoza, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage20.ResumeLayout(False)
            CType(Me.dgvGazo, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage21.ResumeLayout(False)
            CType(Me.dgvFBInfo, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage22.ResumeLayout(False)
            CType(Me.dgvKyrui, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage23.ResumeLayout(False)
            CType(Me.dgvTosiYoto, System.ComponentModel.ISupportInitialize).EndInit()
            Me.tabPage4.ResumeLayout(False)
            CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.pnlHaisyoku.ResumeLayout(False)
            Me.pnlRelSelectPanel.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub
        Friend WithEvents tabCtrlRelMain As System.Windows.Forms.TabControl
        Friend WithEvents tabPage1 As System.Windows.Forms.TabPage
        Friend WithEvents tabPage0 As System.Windows.Forms.TabPage
        Friend WithEvents tabPage2 As System.Windows.Forms.TabPage
        Friend WithEvents btnConnectTest As System.Windows.Forms.Button
        Friend WithEvents lblDBTxt As System.Windows.Forms.Label
        Friend WithEvents tabPage3 As System.Windows.Forms.TabPage
        Friend WithEvents lblFirsttxt As System.Windows.Forms.Label
        Friend WithEvents btnRelDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtRelationDirPath As System.Windows.Forms.TextBox
        Friend WithEvents lblRelationDir As System.Windows.Forms.Label
        Friend WithEvents lblLogDir As System.Windows.Forms.Label
        Friend WithEvents txtLogDirPath As System.Windows.Forms.TextBox
        Friend WithEvents btnLogDirSeach As System.Windows.Forms.Button
        Friend WithEvents grpDev As System.Windows.Forms.GroupBox
        Friend WithEvents Label11 As System.Windows.Forms.Label
        Friend WithEvents chkTempTable As System.Windows.Forms.CheckBox
        Friend WithEvents chkRelFile As System.Windows.Forms.CheckBox
        Friend WithEvents tabPage4 As System.Windows.Forms.TabPage
        Friend WithEvents lblFinaltxt As System.Windows.Forms.Label
        Friend WithEvents tabCtrlRelwork As System.Windows.Forms.TabControl
        Friend WithEvents tabPage10 As System.Windows.Forms.TabPage
        Friend WithEvents dgvBkrui As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage11 As System.Windows.Forms.TabPage
        Friend WithEvents dgvHyrui As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage12 As System.Windows.Forms.TabPage
        Friend WithEvents dgvNkinKbn As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage13 As System.Windows.Forms.TabPage
        Friend WithEvents dgvToritaiyo As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage14 As System.Windows.Forms.TabPage
        Friend WithEvents dgvNkinkomk As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage15 As System.Windows.Forms.TabPage
        Friend WithEvents dgvKozo As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage16 As System.Windows.Forms.TabPage
        Friend WithEvents dgvKozasyu As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage17 As System.Windows.Forms.TabPage
        Friend WithEvents dgvSetubi As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage18 As System.Windows.Forms.TabPage
        Friend WithEvents tabPage19 As System.Windows.Forms.TabPage
        Friend WithEvents chkBkrui As System.Windows.Forms.CheckBox
        Friend WithEvents btnMidDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtMidDirPath As System.Windows.Forms.TextBox
        Friend WithEvents lblMidDir As System.Windows.Forms.Label
        Friend WithEvents chkDuplicate As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkNkinkbn As System.Windows.Forms.CheckBox
        Friend WithEvents chkToritaiyo As System.Windows.Forms.CheckBox
        Friend WithEvents chkNkinkomk As System.Windows.Forms.CheckBox
        Friend WithEvents chkKozo As System.Windows.Forms.CheckBox
        Friend WithEvents chkKozasyu As System.Windows.Forms.CheckBox
        Friend WithEvents chkSetubi As System.Windows.Forms.CheckBox
        Friend WithEvents chkKagi As System.Windows.Forms.CheckBox
        Friend WithEvents dgvKagi As System.Windows.Forms.DataGridView
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents grpKagi As System.Windows.Forms.GroupBox
        Friend WithEvents optKyKagi As System.Windows.Forms.RadioButton
        Friend WithEvents optHyKagi As System.Windows.Forms.RadioButton
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents chkJisyakoza As System.Windows.Forms.CheckBox
        Friend WithEvents dgvJisyakoza As System.Windows.Forms.DataGridView
        Friend WithEvents chkLogTblDrop As System.Windows.Forms.CheckBox
        Friend WithEvents pgbRelItemRead As System.Windows.Forms.ProgressBar
        Friend WithEvents lblRelItemRead As System.Windows.Forms.Label
        Friend WithEvents btnBack As System.Windows.Forms.Button
        Friend WithEvents btnNext As System.Windows.Forms.Button
        Friend WithEvents btnEnd As System.Windows.Forms.Button
        Friend WithEvents grp10ConnectInfo As System.Windows.Forms.GroupBox
        Friend WithEvents grpV10Authent As System.Windows.Forms.GroupBox
        Friend WithEvents optV10Authent2 As System.Windows.Forms.RadioButton
        Friend WithEvents optV10Authent1 As System.Windows.Forms.RadioButton
        Friend WithEvents txtV10Pass As System.Windows.Forms.TextBox
        Friend WithEvents txtV10User As System.Windows.Forms.TextBox
        Friend WithEvents txtV10Networklib As System.Windows.Forms.TextBox
        Friend WithEvents txtV10Catalog As System.Windows.Forms.TextBox
        Friend WithEvents txtV10Server As System.Windows.Forms.TextBox
        Friend WithEvents lblV10 As System.Windows.Forms.Label
        Friend WithEvents grpV7ConnectInfo As System.Windows.Forms.GroupBox
        Friend WithEvents grpV7Authent As System.Windows.Forms.GroupBox
        Friend WithEvents optV7Authent2 As System.Windows.Forms.RadioButton
        Friend WithEvents optV7Authent1 As System.Windows.Forms.RadioButton
        Friend WithEvents txtV7Pass As System.Windows.Forms.TextBox
        Friend WithEvents txtV7User As System.Windows.Forms.TextBox
        Friend WithEvents txtV7Networklib As System.Windows.Forms.TextBox
        Friend WithEvents txtV7Catalog As System.Windows.Forms.TextBox
        Friend WithEvents txtV7Server As System.Windows.Forms.TextBox
        Friend WithEvents lblV7 As System.Windows.Forms.Label
        Friend WithEvents Label57 As System.Windows.Forms.Label
        Friend WithEvents txtTimeOut As System.Windows.Forms.TextBox
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents Label36 As System.Windows.Forms.Label
        Friend WithEvents Label35 As System.Windows.Forms.Label
        Friend WithEvents lblNetworklib As System.Windows.Forms.Label
        Friend WithEvents Label30 As System.Windows.Forms.Label
        Friend WithEvents Label13 As System.Windows.Forms.Label
        Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
        Friend WithEvents Label5 As System.Windows.Forms.Label
        Friend WithEvents lblCaution As System.Windows.Forms.Label
        Friend WithEvents Label6 As System.Windows.Forms.Label
        Friend WithEvents Label7 As System.Windows.Forms.Label
        Friend WithEvents Label8 As System.Windows.Forms.Label
        Friend WithEvents Label9 As System.Windows.Forms.Label
        Friend WithEvents Label10 As System.Windows.Forms.Label
        Friend WithEvents Label12 As System.Windows.Forms.Label
        Friend WithEvents Label14 As System.Windows.Forms.Label
        Friend WithEvents Label15 As System.Windows.Forms.Label
        Friend WithEvents chkGazo As System.Windows.Forms.CheckBox
        Friend WithEvents tabPage20 As System.Windows.Forms.TabPage
        Friend WithEvents dgvGazo As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage21 As System.Windows.Forms.TabPage
        Friend WithEvents dgvFBInfo As System.Windows.Forms.DataGridView
        Friend WithEvents chkFBFmt As System.Windows.Forms.CheckBox
        Friend WithEvents tabPage22 As System.Windows.Forms.TabPage
        Friend WithEvents chkKyrui As System.Windows.Forms.CheckBox
        Friend WithEvents dgvKyrui As System.Windows.Forms.DataGridView
        Friend WithEvents tabPage23 As System.Windows.Forms.TabPage
        Friend WithEvents dgvTosiYoto As System.Windows.Forms.DataGridView
        Friend WithEvents chkTosiYoto As System.Windows.Forms.CheckBox
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents ColFBInfo0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo3 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo4 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo5 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo6 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo7 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo8 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo9 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo10 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo11 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo12 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColFBInfo13 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents btnSetNkinnoBulk As System.Windows.Forms.Button
        Friend WithEvents Label18 As System.Windows.Forms.Label
        Friend WithEvents Label17 As System.Windows.Forms.Label
        Friend WithEvents txtNkinnoSta As System.Windows.Forms.TextBox
        Friend WithEvents Label16 As System.Windows.Forms.Label
        Friend WithEvents Label19 As System.Windows.Forms.Label
        Friend WithEvents Label20 As System.Windows.Forms.Label
        Friend WithEvents Label21 As System.Windows.Forms.Label
        Friend WithEvents Label22 As System.Windows.Forms.Label
        Friend WithEvents Label23 As System.Windows.Forms.Label
        Friend WithEvents Label24 As System.Windows.Forms.Label
        Friend WithEvents Label25 As System.Windows.Forms.Label
        Friend WithEvents Label26 As System.Windows.Forms.Label
        Friend WithEvents Label27 As System.Windows.Forms.Label
        Friend WithEvents Label28 As System.Windows.Forms.Label
        Friend WithEvents lbltest As System.Windows.Forms.Label
        Friend WithEvents Label31 As System.Windows.Forms.Label
        Friend WithEvents ColNkinkomk0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk3 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk4 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk5 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk6 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk7 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinkomk8 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents Label32 As System.Windows.Forms.Label
        Friend WithEvents Label33 As System.Windows.Forms.Label
        Friend WithEvents ColBkrui0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColBkrui1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColBkrui2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColHyrui0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColHyrui1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColHyrui2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinKbn0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinKbn1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinKbn2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinKbn3 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColNkinKbn4 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColToritaiyo0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColToritaiyo1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColToritaiyo2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKozo0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKozo1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKozo2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKozasyu0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKozasyu1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKozasyu2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi3 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi4 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi5 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi6 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColSetubi7 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKyrui0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKyrui1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColKyrui2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColTosiYoto0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColTosiYoto1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColTosiYoto2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColTosiYoto3 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents Label29 As System.Windows.Forms.Label
        Friend WithEvents pnlTimeOut As System.Windows.Forms.Panel
        Friend WithEvents lblHidden3 As System.Windows.Forms.Label
        Friend WithEvents lblLine0 As System.Windows.Forms.Label
        Friend WithEvents Label34 As System.Windows.Forms.Label
        Friend WithEvents Label38 As System.Windows.Forms.Label
        Friend WithEvents Label39 As System.Windows.Forms.Label
        Friend WithEvents Label37 As System.Windows.Forms.Label
        Friend WithEvents Label40 As System.Windows.Forms.Label
        Friend WithEvents Panel1 As System.Windows.Forms.Panel
        Friend WithEvents Label44 As System.Windows.Forms.Label
        Friend WithEvents Label43 As System.Windows.Forms.Label
        Friend WithEvents Label42 As System.Windows.Forms.Label
        Friend WithEvents Label45 As System.Windows.Forms.Label
        Friend WithEvents Label46 As System.Windows.Forms.Label
        Friend WithEvents Label47 As System.Windows.Forms.Label
        Friend WithEvents pnlHaisyoku As System.Windows.Forms.Panel
        Friend WithEvents pnlRelSelectPanel As System.Windows.Forms.Panel
        Friend WithEvents lblRelSelectTosiYoto As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectJisyaKoza As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectKyBunrui As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectGazo As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectFBFmt As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectNkin As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectHySetubi As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectKozo As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectKozaSyubetu As System.Windows.Forms.Label
        Friend WithEvents lblTitleRelSelect As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectBkBunrui As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectTaiyo As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectHyBunrui As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectNkKbn As System.Windows.Forms.Label
        Friend WithEvents lblHidden4 As System.Windows.Forms.Label
        Friend WithEvents lblLine2 As System.Windows.Forms.Label
        Friend WithEvents lblCntAllTosiYoto As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiTosiYoto As System.Windows.Forms.Label
        Friend WithEvents lblCntAllKyBunrui As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiKyBunrui As System.Windows.Forms.Label
        Friend WithEvents lblCntAllFbFmt As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiFbFmt As System.Windows.Forms.Label
        Friend WithEvents lblCntAllGazo As System.Windows.Forms.Label
        Friend WithEvents lblSumiZanGazo As System.Windows.Forms.Label
        Friend WithEvents lblCntAllJisyaKoza As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiJisyaKoza As System.Windows.Forms.Label
        Friend WithEvents lblCntAllSetubi As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiSetubi As System.Windows.Forms.Label
        Friend WithEvents lblCntAllKozaSyubetu As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiKozaSyubetu As System.Windows.Forms.Label
        Friend WithEvents lblCntAllKozo As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiKozo As System.Windows.Forms.Label
        Friend WithEvents lblCntAllNkin As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiNkin As System.Windows.Forms.Label
        Friend WithEvents lblCntAllTaiyo As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiTaiyo As System.Windows.Forms.Label
        Friend WithEvents lblCntAllNkKbn As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiNkKbn As System.Windows.Forms.Label
        Friend WithEvents lblCntAllHyBunrui As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiHyBunrui As System.Windows.Forms.Label
        Friend WithEvents lblCntAllBkBunrui As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiBkBunrui As System.Windows.Forms.Label
        Friend WithEvents lblTitleRelTotalCnt As System.Windows.Forms.Label
        Friend WithEvents lblTitleRelCnt As System.Windows.Forms.Label
        Friend WithEvents lblTitleRelKomk As System.Windows.Forms.Label
        Friend WithEvents lblCntBkBunruiBack As System.Windows.Forms.Label
        Friend WithEvents lblCntTosiYotoBack As System.Windows.Forms.Label
        Friend WithEvents lblCntKyBunruiBack As System.Windows.Forms.Label
        Friend WithEvents lblCntFbFmtBack As System.Windows.Forms.Label
        Friend WithEvents lblCntGazoBack As System.Windows.Forms.Label
        Friend WithEvents lblCntNkKbnBack As System.Windows.Forms.Label
        Friend WithEvents lblCntTaiyoBack As System.Windows.Forms.Label
        Friend WithEvents lblCntNkinBack As System.Windows.Forms.Label
        Friend WithEvents lblCntJisyaKozaBack As System.Windows.Forms.Label
        Friend WithEvents lblCntSetubiBack As System.Windows.Forms.Label
        Friend WithEvents lblCntKozaSyubetuBack As System.Windows.Forms.Label
        Friend WithEvents lblCntKozoBack As System.Windows.Forms.Label
        Friend WithEvents lblCntHyBunruiBack As System.Windows.Forms.Label
        Friend WithEvents ColGazo0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColGazo1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColGazo2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColGazo3 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColGazo4 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColGazo5 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza0 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza1 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza2 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza3 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza4 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza5 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza6 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza7 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza8 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza9 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza10 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza11 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza12 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza13 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza14 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza15 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza16 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents ColJisyakoza17 As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents lblCntAllKyoyoKagi As System.Windows.Forms.Label
        Friend WithEvents lblCntSumiKyoyoKagi As System.Windows.Forms.Label
        Friend WithEvents lblRelSelectKyoyoKagi As System.Windows.Forms.Label
        Friend WithEvents lblCntKyoyoKagiBack As System.Windows.Forms.Label
        Friend WithEvents grpKagicntgrp As System.Windows.Forms.GroupBox
    End Class

End Namespace