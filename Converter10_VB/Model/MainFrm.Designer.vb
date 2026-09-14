Namespace Njc.Frm

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class MainFrm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainFrm))
        Me.chkRelationFile = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblTitleKi = New System.Windows.Forms.Label()
        Me.pnlMenuDatacv = New System.Windows.Forms.Panel()
        Me.pnlPrgDCConv = New System.Windows.Forms.Panel()
        Me.picIcoDCConv2 = New System.Windows.Forms.PictureBox()
        Me.picIcoDCConv1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgDCConv = New System.Windows.Forms.Label()
        Me.pnlPrgDCRelation = New System.Windows.Forms.Panel()
        Me.picIcoDCRelation2 = New System.Windows.Forms.PictureBox()
        Me.picIcoDCRelation1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgDCRelation = New System.Windows.Forms.Label()
        Me.pnlPrgDCFileWrite = New System.Windows.Forms.Panel()
        Me.picIcoDCFileWrite2 = New System.Windows.Forms.PictureBox()
        Me.picIcoDCFileWrite1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgDCFileWrite = New System.Windows.Forms.Label()
        Me.picArwDCEnd2 = New System.Windows.Forms.PictureBox()
        Me.picArwDCJikko2 = New System.Windows.Forms.PictureBox()
        Me.picArwDCSelect2 = New System.Windows.Forms.PictureBox()
        Me.picArwDCEnd1 = New System.Windows.Forms.PictureBox()
        Me.picArwDCJikko1 = New System.Windows.Forms.PictureBox()
        Me.picArwDCSelect1 = New System.Windows.Forms.PictureBox()
        Me.pnlPrgDCEnd = New System.Windows.Forms.Panel()
        Me.picIcoDCEnd2 = New System.Windows.Forms.PictureBox()
        Me.picIcoDCEnd1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgDCEnd = New System.Windows.Forms.Label()
        Me.pnlPrgDCJikko = New System.Windows.Forms.Panel()
        Me.picIcoDCJikko2 = New System.Windows.Forms.PictureBox()
        Me.picIcoDCJikko1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgDCJikko = New System.Windows.Forms.Label()
        Me.pnlPrgDCSelect = New System.Windows.Forms.Panel()
        Me.picIcoDCSelect2 = New System.Windows.Forms.PictureBox()
        Me.picIcoDCSelect1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgDCSelect = New System.Windows.Forms.Label()
        Me.pnlPrgDCHajimeni = New System.Windows.Forms.Panel()
        Me.picIcoDCHajimeni2 = New System.Windows.Forms.PictureBox()
        Me.picIcoDCHajimeni1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgDCHajimeni = New System.Windows.Forms.Label()
        Me.lblHidden1 = New System.Windows.Forms.Label()
        Me.tabCtrlMain = New System.Windows.Forms.TabControl()
        Me.tabPageDev = New System.Windows.Forms.TabPage()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.Label361 = New System.Windows.Forms.Label()
        Me.txtExistMidToBaseMid = New System.Windows.Forms.TextBox()
        Me.btnExistMidToBaseMid = New System.Windows.Forms.Button()
        Me.btnExistMidToBaseMidDirSerach = New System.Windows.Forms.Button()
        Me.grpDevSettingX = New System.Windows.Forms.GroupBox()
        Me.chkChildItemControl = New System.Windows.Forms.CheckBox()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.txtLogOutputCnt = New System.Windows.Forms.TextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.grpDevSetting1 = New System.Windows.Forms.GroupBox()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.Label360 = New System.Windows.Forms.Label()
        Me.chkKyNewest = New System.Windows.Forms.CheckBox()
        Me.chkSyskanriinit = New System.Windows.Forms.CheckBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.chkRelTblDrop = New System.Windows.Forms.CheckBox()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.chkV7ViewDrop = New System.Windows.Forms.CheckBox()
        Me.chkLogTblDrop = New System.Windows.Forms.CheckBox()
        Me.chkOverWrite = New System.Windows.Forms.CheckBox()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.grpDevSetting2 = New System.Windows.Forms.GroupBox()
        Me.btnRelDirSeach = New System.Windows.Forms.Button()
        Me.txtRelationDirPath = New System.Windows.Forms.TextBox()
        Me.lblRelationDir = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.chkRelation = New System.Windows.Forms.CheckBox()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.chkMidNotStop = New System.Windows.Forms.CheckBox()
        Me.chkCVStart = New System.Windows.Forms.CheckBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.chkMiddleFile = New System.Windows.Forms.CheckBox()
        Me.grpDevSetting3 = New System.Windows.Forms.GroupBox()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.chkRommDuplicate = New System.Windows.Forms.CheckBox()
        Me.chkEmptyRoomNo = New System.Windows.Forms.CheckBox()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.grpRelation = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkMstKasyorui = New System.Windows.Forms.CheckBox()
        Me.chkMstGenjotokuyaku = New System.Windows.Forms.CheckBox()
        Me.lblTokuyakuInfo = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.chkMstKozasyubetu = New System.Windows.Forms.CheckBox()
        Me.chkMstSetubi = New System.Windows.Forms.CheckBox()
        Me.chkMstKozo = New System.Windows.Forms.CheckBox()
        Me.chkMstTorihikitaiyo = New System.Windows.Forms.CheckBox()
        Me.chkMstNkinkbn = New System.Windows.Forms.CheckBox()
        Me.chkMstHyrui = New System.Windows.Forms.CheckBox()
        Me.chkMstNkinkomok = New System.Windows.Forms.CheckBox()
        Me.chkMstBkrui = New System.Windows.Forms.CheckBox()
        Me.grptaihi = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnFileSeach = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.chkGyYatinhosyoKoza = New System.Windows.Forms.CheckBox()
        Me.CheckBox66 = New System.Windows.Forms.CheckBox()
        Me.lblGy = New System.Windows.Forms.Label()
        Me.lblMst = New System.Windows.Forms.Label()
        Me.lblKys = New System.Windows.Forms.Label()
        Me.lblOwner = New System.Windows.Forms.Label()
        Me.lblJisya = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.chkMs25 = New System.Windows.Forms.CheckBox()
        Me.chkMsKeiyakusyubetu = New System.Windows.Forms.CheckBox()
        Me.chkMs23 = New System.Windows.Forms.CheckBox()
        Me.chkMs24 = New System.Windows.Forms.CheckBox()
        Me.chkMsOwevent = New System.Windows.Forms.CheckBox()
        Me.chkMsYane = New System.Windows.Forms.CheckBox()
        Me.chkMsKeikaikakunin = New System.Windows.Forms.CheckBox()
        Me.chkMsKinyu = New System.Windows.Forms.CheckBox()
        Me.chkMsKinyuten = New System.Windows.Forms.CheckBox()
        Me.tabPageStart = New System.Windows.Forms.TabPage()
        Me.grpKFirstNaiyo = New System.Windows.Forms.GroupBox()
        Me.lblHajimeKi352 = New System.Windows.Forms.Label()
        Me.lblHajimeKi351 = New System.Windows.Forms.Label()
        Me.lblHajimeKi342 = New System.Windows.Forms.Label()
        Me.lblHajimeKi341 = New System.Windows.Forms.Label()
        Me.lblHajimeKi332 = New System.Windows.Forms.Label()
        Me.lblHajimeKi322 = New System.Windows.Forms.Label()
        Me.lblHajimeKi312 = New System.Windows.Forms.Label()
        Me.lblHajimeKi331 = New System.Windows.Forms.Label()
        Me.lblHajimeKi321 = New System.Windows.Forms.Label()
        Me.lblHajimeKi311 = New System.Windows.Forms.Label()
        Me.lblHajimeKi302 = New System.Windows.Forms.Label()
        Me.lblHajimeKi202 = New System.Windows.Forms.Label()
        Me.lblHajimeKi102 = New System.Windows.Forms.Label()
        Me.lblHajimeKi301 = New System.Windows.Forms.Label()
        Me.lblHajimeKi201 = New System.Windows.Forms.Label()
        Me.lblHajimeKi101 = New System.Windows.Forms.Label()
        Me.lblFirstDescription1 = New System.Windows.Forms.Label()
        Me.lblKFirstLabel = New System.Windows.Forms.Label()
        Me.lblHFirstLabel = New System.Windows.Forms.Label()
        Me.grpHFirstNaiyo = New System.Windows.Forms.GroupBox()
        Me.lblHajimeH312 = New System.Windows.Forms.Label()
        Me.lblHajimeH311 = New System.Windows.Forms.Label()
        Me.lblHajimeH342 = New System.Windows.Forms.Label()
        Me.lblHajimeH322 = New System.Windows.Forms.Label()
        Me.lblHajimeH341 = New System.Windows.Forms.Label()
        Me.lblHajimeH321 = New System.Windows.Forms.Label()
        Me.lblHajimeH302 = New System.Windows.Forms.Label()
        Me.lblHajimeH202 = New System.Windows.Forms.Label()
        Me.lblHajimeH102 = New System.Windows.Forms.Label()
        Me.lblHajimeH301 = New System.Windows.Forms.Label()
        Me.lblHajimeH201 = New System.Windows.Forms.Label()
        Me.lblHajimeH101 = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.tabPageSession = New System.Windows.Forms.TabPage()
        Me.grpTimeOut = New System.Windows.Forms.GroupBox()
        Me.txtTimeOut = New System.Windows.Forms.TextBox()
        Me.lblTimeOutSec = New System.Windows.Forms.Label()
        Me.lblSessionCaution = New System.Windows.Forms.Label()
        Me.lblKSessionDescription1 = New System.Windows.Forms.Label()
        Me.lblHSessionDescription1 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.lblNetworklib = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.grp10ConnectInfo = New System.Windows.Forms.GroupBox()
        Me.cmbV10Networklib = New System.Windows.Forms.ComboBox()
        Me.grpV10Authent = New System.Windows.Forms.GroupBox()
        Me.optV10Authent2 = New System.Windows.Forms.RadioButton()
        Me.optV10Authent1 = New System.Windows.Forms.RadioButton()
        Me.txtV10Pass = New System.Windows.Forms.TextBox()
        Me.txtV10User = New System.Windows.Forms.TextBox()
        Me.txtV10Catalog = New System.Windows.Forms.TextBox()
        Me.txtV10Server = New System.Windows.Forms.TextBox()
        Me.lblV10 = New System.Windows.Forms.Label()
        Me.grpV7ConnectInfo = New System.Windows.Forms.GroupBox()
        Me.cmbV7Networklib = New System.Windows.Forms.ComboBox()
        Me.grpV7Authent = New System.Windows.Forms.GroupBox()
        Me.optV7Authent2 = New System.Windows.Forms.RadioButton()
        Me.optV7Authent1 = New System.Windows.Forms.RadioButton()
        Me.txtV7Pass = New System.Windows.Forms.TextBox()
        Me.txtV7User = New System.Windows.Forms.TextBox()
        Me.txtV7Catalog = New System.Windows.Forms.TextBox()
        Me.txtV7Server = New System.Windows.Forms.TextBox()
        Me.lblV7 = New System.Windows.Forms.Label()
        Me.btnDefConInfoRead = New System.Windows.Forms.Button()
        Me.btnConnectTest = New System.Windows.Forms.Button()
        Me.tabPageSyoki = New System.Windows.Forms.TabPage()
        Me.pnlConvertType = New System.Windows.Forms.Panel()
        Me.Label106 = New System.Windows.Forms.Label()
        Me.pnlCVType = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.optCVNew = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.optCVAdd = New System.Windows.Forms.RadioButton()
        Me.Label107 = New System.Windows.Forms.Label()
        Me.pnlUnyoKaisi = New System.Windows.Forms.Panel()
        Me.txtUnyoYYYYMM = New System.Windows.Forms.MaskedTextBox()
        Me.Label91 = New System.Windows.Forms.Label()
        Me.Label98 = New System.Windows.Forms.Label()
        Me.Label137 = New System.Windows.Forms.Label()
        Me.pnlOptSelect = New System.Windows.Forms.Panel()
        Me.lblKOpSeleTitle = New System.Windows.Forms.Label()
        Me.pnlOptionSelect = New System.Windows.Forms.Panel()
        Me.chkOptionSelectKaikei = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectKy = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectReform = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectSq = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectClaim = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectNk = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectFB = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectSh = New System.Windows.Forms.CheckBox()
        Me.chkOptionSelectSyoki = New System.Windows.Forms.CheckBox()
        Me.lblHOpSeleTitle = New System.Windows.Forms.Label()
        Me.lblHOpSeleNaiyo = New System.Windows.Forms.Label()
        Me.lblSOpSeleNaiyo = New System.Windows.Forms.Label()
        Me.lblKOpSeleNaiyo = New System.Windows.Forms.Label()
        Me.pnlUserName = New System.Windows.Forms.Panel()
        Me.lblExecutor = New System.Windows.Forms.Label()
        Me.txtRecUser = New System.Windows.Forms.TextBox()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.pnlLogPath = New System.Windows.Forms.Panel()
        Me.lblLogDir = New System.Windows.Forms.Label()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.btnLogDirSeach = New System.Windows.Forms.Button()
        Me.txtLogDirPath = New System.Windows.Forms.TextBox()
        Me.lblSyokiDescription1 = New System.Windows.Forms.Label()
        Me.tabPageMenu = New System.Windows.Forms.TabPage()
        Me.grpMenuHojyo = New System.Windows.Forms.GroupBox()
        Me.btnMenuHojyo = New System.Windows.Forms.Button()
        Me.lblMenuHojyo = New System.Windows.Forms.Label()
        Me.lblMenuCaution = New System.Windows.Forms.Label()
        Me.grpMenuJizen = New System.Windows.Forms.GroupBox()
        Me.lblKMenuJizen = New System.Windows.Forms.Label()
        Me.lblHMenuJizen = New System.Windows.Forms.Label()
        Me.btnMenuJizen = New System.Windows.Forms.Button()
        Me.grpMenuDatacv = New System.Windows.Forms.GroupBox()
        Me.lblKMenuDatacv = New System.Windows.Forms.Label()
        Me.lblHMenuDatacv = New System.Windows.Forms.Label()
        Me.btnMenuDatacv = New System.Windows.Forms.Button()
        Me.grpMenuGazocv = New System.Windows.Forms.GroupBox()
        Me.btnMenuGazocv = New System.Windows.Forms.Button()
        Me.lblMenuGazocv = New System.Windows.Forms.Label()
        Me.grpMenuJigo = New System.Windows.Forms.GroupBox()
        Me.btnMenuJigo = New System.Windows.Forms.Button()
        Me.lblMenuJigo = New System.Windows.Forms.Label()
        Me.lblMenuDescription1 = New System.Windows.Forms.Label()
        Me.tabPageJizen = New System.Windows.Forms.TabPage()
        Me.pnlJizenListPath = New System.Windows.Forms.Panel()
        Me.Label254 = New System.Windows.Forms.Label()
        Me.btnJizenListDirSeach = New System.Windows.Forms.Button()
        Me.txtJizenListPath = New System.Windows.Forms.TextBox()
        Me.lblLine1 = New System.Windows.Forms.Label()
        Me.lblHidden3 = New System.Windows.Forms.Label()
        Me.tabCtrlJizen = New System.Windows.Forms.TabControl()
        Me.tabPageJizen1 = New System.Windows.Forms.TabPage()
        Me.grpJizen1 = New System.Windows.Forms.GroupBox()
        Me.pnlJizenHurikae = New System.Windows.Forms.Panel()
        Me.chkJizenHurikae = New System.Windows.Forms.CheckBox()
        Me.Label88 = New System.Windows.Forms.Label()
        Me.pnlJizenAzu = New System.Windows.Forms.Panel()
        Me.chkJizenAzu = New System.Windows.Forms.CheckBox()
        Me.Label85 = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.btnListJizenAzu = New System.Windows.Forms.Button()
        Me.Label166 = New System.Windows.Forms.Label()
        Me.Label167 = New System.Windows.Forms.Label()
        Me.Label168 = New System.Windows.Forms.Label()
        Me.lblCntJizenAzu = New System.Windows.Forms.Label()
        Me.pnlJizenKai = New System.Windows.Forms.Panel()
        Me.chkJizenKai = New System.Windows.Forms.CheckBox()
        Me.Label84 = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.btnListJizenKai = New System.Windows.Forms.Button()
        Me.Label161 = New System.Windows.Forms.Label()
        Me.Label164 = New System.Windows.Forms.Label()
        Me.Label185 = New System.Windows.Forms.Label()
        Me.lblCntJizenKai = New System.Windows.Forms.Label()
        Me.lblJizenPageCnt1 = New System.Windows.Forms.Label()
        Me.tabPageJizen2 = New System.Windows.Forms.TabPage()
        Me.grpJizen2 = New System.Windows.Forms.GroupBox()
        Me.Panel12 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.btnListJizenHasseiHen = New System.Windows.Forms.Button()
        Me.Label125 = New System.Windows.Forms.Label()
        Me.Label130 = New System.Windows.Forms.Label()
        Me.Label131 = New System.Windows.Forms.Label()
        Me.lblCntJizenHasseiHen = New System.Windows.Forms.Label()
        Me.chkJizenHasseiHen = New System.Windows.Forms.CheckBox()
        Me.Label165 = New System.Windows.Forms.Label()
        Me.pnlJizenHasseiOw = New System.Windows.Forms.Panel()
        Me.chkJizenHasseiOw = New System.Windows.Forms.CheckBox()
        Me.Label100 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnListJizenHasseiOw = New System.Windows.Forms.Button()
        Me.Label89 = New System.Windows.Forms.Label()
        Me.Label118 = New System.Windows.Forms.Label()
        Me.Label120 = New System.Windows.Forms.Label()
        Me.lblCntJizenHasseiOw = New System.Windows.Forms.Label()
        Me.pnlJizenYanuso = New System.Windows.Forms.Panel()
        Me.chkJizenYanuso = New System.Windows.Forms.CheckBox()
        Me.Label96 = New System.Windows.Forms.Label()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.btnListJizenYanuso = New System.Windows.Forms.Button()
        Me.Label188 = New System.Windows.Forms.Label()
        Me.Label189 = New System.Windows.Forms.Label()
        Me.Label190 = New System.Windows.Forms.Label()
        Me.lblCntJizenYanuso = New System.Windows.Forms.Label()
        Me.pnlJizenBkhourei = New System.Windows.Forms.Panel()
        Me.chkJizenBkhourei = New System.Windows.Forms.CheckBox()
        Me.Label94 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label179 = New System.Windows.Forms.Label()
        Me.Label198 = New System.Windows.Forms.Label()
        Me.lblCntJizenBkhoureiNotCV = New System.Windows.Forms.Label()
        Me.btnListJizenBkhourei = New System.Windows.Forms.Button()
        Me.Label103 = New System.Windows.Forms.Label()
        Me.Label117 = New System.Windows.Forms.Label()
        Me.Label171 = New System.Windows.Forms.Label()
        Me.lblCntJizenBkhoureiNotDef = New System.Windows.Forms.Label()
        Me.lblJizenPageCnt2 = New System.Windows.Forms.Label()
        Me.tabPageJizen3 = New System.Windows.Forms.TabPage()
        Me.grpJizen3 = New System.Windows.Forms.GroupBox()
        Me.pnlJizenKagi = New System.Windows.Forms.Panel()
        Me.chkJizenKagi = New System.Windows.Forms.CheckBox()
        Me.Label262 = New System.Windows.Forms.Label()
        Me.Panel19 = New System.Windows.Forms.Panel()
        Me.btnListJizenKyKagi = New System.Windows.Forms.Button()
        Me.Label266 = New System.Windows.Forms.Label()
        Me.Label269 = New System.Windows.Forms.Label()
        Me.lblCntJizenKyKagi = New System.Windows.Forms.Label()
        Me.btnListJizenHyKagi = New System.Windows.Forms.Button()
        Me.Label276 = New System.Windows.Forms.Label()
        Me.Label279 = New System.Windows.Forms.Label()
        Me.Label283 = New System.Windows.Forms.Label()
        Me.lblCntJizenHyKagi = New System.Windows.Forms.Label()
        Me.pnlJizenSzenKyshutan = New System.Windows.Forms.Panel()
        Me.chkJizenSzenKyshutan = New System.Windows.Forms.CheckBox()
        Me.Label173 = New System.Windows.Forms.Label()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.btnListJizenSzenKyshutan = New System.Windows.Forms.Button()
        Me.Label202 = New System.Windows.Forms.Label()
        Me.Label203 = New System.Windows.Forms.Label()
        Me.Label204 = New System.Windows.Forms.Label()
        Me.lblCntJizenSzenKyshutan = New System.Windows.Forms.Label()
        Me.Label224 = New System.Windows.Forms.Label()
        Me.tabPageJizen4 = New System.Windows.Forms.TabPage()
        Me.grpJizen4 = New System.Windows.Forms.GroupBox()
        Me.Panel15 = New System.Windows.Forms.Panel()
        Me.chkJizenKozameigikana = New System.Windows.Forms.CheckBox()
        Me.Label353 = New System.Windows.Forms.Label()
        Me.Panel20 = New System.Windows.Forms.Panel()
        Me.btnListJizenKozameigikana = New System.Windows.Forms.Button()
        Me.Label354 = New System.Windows.Forms.Label()
        Me.Label356 = New System.Windows.Forms.Label()
        Me.Label357 = New System.Windows.Forms.Label()
        Me.lblCntJizenKozameigikana = New System.Windows.Forms.Label()
        Me.Panel21 = New System.Windows.Forms.Panel()
        Me.chkJizenNonJisyaKoza = New System.Windows.Forms.CheckBox()
        Me.Label365 = New System.Windows.Forms.Label()
        Me.Panel27 = New System.Windows.Forms.Panel()
        Me.btnListJizenNonJisyaKoza = New System.Windows.Forms.Button()
        Me.Label366 = New System.Windows.Forms.Label()
        Me.Label367 = New System.Windows.Forms.Label()
        Me.Label368 = New System.Windows.Forms.Label()
        Me.lblCntJizenNonJisyaKoza = New System.Windows.Forms.Label()
        Me.Label370 = New System.Windows.Forms.Label()
        Me.tabPageJizen5 = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label374 = New System.Windows.Forms.Label()
        Me.Panel13 = New System.Windows.Forms.Panel()
        Me.chkJizenSimeSokin = New System.Windows.Forms.CheckBox()
        Me.Label149 = New System.Windows.Forms.Label()
        Me.Panel14 = New System.Windows.Forms.Panel()
        Me.btnListJizenSimeSokin = New System.Windows.Forms.Button()
        Me.Label310 = New System.Windows.Forms.Label()
        Me.Label348 = New System.Windows.Forms.Label()
        Me.Label351 = New System.Windows.Forms.Label()
        Me.lblCntJizenSimeSokin = New System.Windows.Forms.Label()
        Me.tabPageHJizen1 = New System.Windows.Forms.TabPage()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.pnlHJizenTyukan = New System.Windows.Forms.Panel()
        Me.Label352 = New System.Windows.Forms.Label()
        Me.chkHJizenTyukan = New System.Windows.Forms.CheckBox()
        Me.Label350 = New System.Windows.Forms.Label()
        Me.Panel26 = New System.Windows.Forms.Panel()
        Me.Label322 = New System.Windows.Forms.Label()
        Me.btnHJizenTyukanOpen = New System.Windows.Forms.Button()
        Me.lblMiddleFile = New System.Windows.Forms.Label()
        Me.txtMidDirPath = New System.Windows.Forms.TextBox()
        Me.btnMidDirSeach = New System.Windows.Forms.Button()
        Me.Label355 = New System.Windows.Forms.Label()
        Me.lblCautionDescription = New System.Windows.Forms.Label()
        Me.lblJizenDescription1 = New System.Windows.Forms.Label()
        Me.tabPageJigo = New System.Windows.Forms.TabPage()
        Me.pnlJigoListPath = New System.Windows.Forms.Panel()
        Me.btnJigoListDirSeach = New System.Windows.Forms.Button()
        Me.Label248 = New System.Windows.Forms.Label()
        Me.txtJigoListPath = New System.Windows.Forms.TextBox()
        Me.Label245 = New System.Windows.Forms.Label()
        Me.lblLine3 = New System.Windows.Forms.Label()
        Me.lblHidden4 = New System.Windows.Forms.Label()
        Me.tabCtrlJigo = New System.Windows.Forms.TabControl()
        Me.tabPageJigo1 = New System.Windows.Forms.TabPage()
        Me.grpJigoUserSagyo = New System.Windows.Forms.GroupBox()
        Me.pnlKJigoCmtSyudo = New System.Windows.Forms.Panel()
        Me.Label306 = New System.Windows.Forms.Label()
        Me.Label305 = New System.Windows.Forms.Label()
        Me.Label143 = New System.Windows.Forms.Label()
        Me.pnlHJigoCmtSyudo = New System.Windows.Forms.Panel()
        Me.Label136 = New System.Windows.Forms.Label()
        Me.Label317 = New System.Windows.Forms.Label()
        Me.Label318 = New System.Windows.Forms.Label()
        Me.pnlJigoCmtSyusi = New System.Windows.Forms.Panel()
        Me.Label304 = New System.Windows.Forms.Label()
        Me.Label303 = New System.Windows.Forms.Label()
        Me.Label146 = New System.Windows.Forms.Label()
        Me.pnlJizenMinus = New System.Windows.Forms.Panel()
        Me.Label358 = New System.Windows.Forms.Label()
        Me.chkJizenMinus = New System.Windows.Forms.CheckBox()
        Me.Label102 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label93 = New System.Windows.Forms.Label()
        Me.Label182 = New System.Windows.Forms.Label()
        Me.Label183 = New System.Windows.Forms.Label()
        Me.lblCntJizenMinus = New System.Windows.Forms.Label()
        Me.btnListJizenMinus = New System.Windows.Forms.Button()
        Me.pnlJigoCmtCsvSyuturyoku = New System.Windows.Forms.Panel()
        Me.Label147 = New System.Windows.Forms.Label()
        Me.Label334 = New System.Windows.Forms.Label()
        Me.Label333 = New System.Windows.Forms.Label()
        Me.pnlJigoCmtSyosiki = New System.Windows.Forms.Panel()
        Me.Label148 = New System.Windows.Forms.Label()
        Me.Label337 = New System.Windows.Forms.Label()
        Me.Label336 = New System.Windows.Forms.Label()
        Me.lblJizenPageNum1 = New System.Windows.Forms.Label()
        Me.tabPageJigo2 = New System.Windows.Forms.TabPage()
        Me.grpJigoCmtTyuui = New System.Windows.Forms.GroupBox()
        Me.pnlJigoCmtHeiko = New System.Windows.Forms.Panel()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label115 = New System.Windows.Forms.Label()
        Me.lblJizenPageNum2 = New System.Windows.Forms.Label()
        Me.grpJigoDonyuji = New System.Windows.Forms.GroupBox()
        Me.pnlJigoCmtSoKotiku = New System.Windows.Forms.Panel()
        Me.Label109 = New System.Windows.Forms.Label()
        Me.Label108 = New System.Windows.Forms.Label()
        Me.Label142 = New System.Windows.Forms.Label()
        Me.pnlJigoCmtNkNyuryoku = New System.Windows.Forms.Panel()
        Me.Label99 = New System.Windows.Forms.Label()
        Me.Label111 = New System.Windows.Forms.Label()
        Me.Label110 = New System.Windows.Forms.Label()
        Me.pnlJigoCmtSqKotiku = New System.Windows.Forms.Panel()
        Me.Label113 = New System.Windows.Forms.Label()
        Me.Label112 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.lblJigoDescription1 = New System.Windows.Forms.Label()
        Me.tabPageHojyo = New System.Windows.Forms.TabPage()
        Me.pnlKensyoListPath = New System.Windows.Forms.Panel()
        Me.btnKensyoListDirSeach = New System.Windows.Forms.Button()
        Me.Label257 = New System.Windows.Forms.Label()
        Me.txtKensyoListPath = New System.Windows.Forms.TextBox()
        Me.lblLine4 = New System.Windows.Forms.Label()
        Me.lblHidden5 = New System.Windows.Forms.Label()
        Me.tabCtrlHojyo = New System.Windows.Forms.TabControl()
        Me.tabPageHojyo1 = New System.Windows.Forms.TabPage()
        Me.GroupBox27 = New System.Windows.Forms.GroupBox()
        Me.pnlKiOpKys = New System.Windows.Forms.Panel()
        Me.Label291 = New System.Windows.Forms.Label()
        Me.chkKiOpKys = New System.Windows.Forms.CheckBox()
        Me.Label184 = New System.Windows.Forms.Label()
        Me.Panel18 = New System.Windows.Forms.Panel()
        Me.Label206 = New System.Windows.Forms.Label()
        Me.btnListKiOpKys = New System.Windows.Forms.Button()
        Me.btnKiOpKys = New System.Windows.Forms.Button()
        Me.Label187 = New System.Windows.Forms.Label()
        Me.Label191 = New System.Windows.Forms.Label()
        Me.Label193 = New System.Windows.Forms.Label()
        Me.lblCntKiOpKys = New System.Windows.Forms.Label()
        Me.pnlKiOpOw = New System.Windows.Forms.Panel()
        Me.Label287 = New System.Windows.Forms.Label()
        Me.chkKiOpOw = New System.Windows.Forms.CheckBox()
        Me.Label169 = New System.Windows.Forms.Label()
        Me.Panel16 = New System.Windows.Forms.Panel()
        Me.btnKiOpOw = New System.Windows.Forms.Button()
        Me.btnListKiOpOw = New System.Windows.Forms.Button()
        Me.Label205 = New System.Windows.Forms.Label()
        Me.Label172 = New System.Windows.Forms.Label()
        Me.Label177 = New System.Windows.Forms.Label()
        Me.Label178 = New System.Windows.Forms.Label()
        Me.lblCntKiOpOw = New System.Windows.Forms.Label()
        Me.pnlRelRename = New System.Windows.Forms.Panel()
        Me.Label274 = New System.Windows.Forms.Label()
        Me.chkRelRename = New System.Windows.Forms.CheckBox()
        Me.Panel25 = New System.Windows.Forms.Panel()
        Me.Label242 = New System.Windows.Forms.Label()
        Me.Label231 = New System.Windows.Forms.Label()
        Me.Label221 = New System.Windows.Forms.Label()
        Me.lblRelRenameSzen = New System.Windows.Forms.Label()
        Me.Label229 = New System.Windows.Forms.Label()
        Me.Label217 = New System.Windows.Forms.Label()
        Me.Label215 = New System.Windows.Forms.Label()
        Me.Label212 = New System.Windows.Forms.Label()
        Me.chkRelRenameSzen = New System.Windows.Forms.CheckBox()
        Me.chkRelRenameClaim = New System.Windows.Forms.CheckBox()
        Me.btnRNRelRename = New System.Windows.Forms.Button()
        Me.Label394 = New System.Windows.Forms.Label()
        Me.lblRelRenameClaim = New System.Windows.Forms.Label()
        Me.Label397 = New System.Windows.Forms.Label()
        Me.txtAfRelRename = New System.Windows.Forms.TextBox()
        Me.Label398 = New System.Windows.Forms.Label()
        Me.txtBfRelRename = New System.Windows.Forms.TextBox()
        Me.Label399 = New System.Windows.Forms.Label()
        Me.Label400 = New System.Windows.Forms.Label()
        Me.Label401 = New System.Windows.Forms.Label()
        Me.Label402 = New System.Windows.Forms.Label()
        Me.tabPageHojyo2 = New System.Windows.Forms.TabPage()
        Me.GroupBox25 = New System.Windows.Forms.GroupBox()
        Me.pnlRelationSet = New System.Windows.Forms.Panel()
        Me.Label307 = New System.Windows.Forms.Label()
        Me.chkRelationSet = New System.Windows.Forms.CheckBox()
        Me.Label225 = New System.Windows.Forms.Label()
        Me.Panel17 = New System.Windows.Forms.Panel()
        Me.btnListRelationSet = New System.Windows.Forms.Button()
        Me.Label234 = New System.Windows.Forms.Label()
        Me.Label237 = New System.Windows.Forms.Label()
        Me.Label239 = New System.Windows.Forms.Label()
        Me.lblCntRelationSet = New System.Windows.Forms.Label()
        Me.pnlSzenKysSorit = New System.Windows.Forms.Panel()
        Me.Label297 = New System.Windows.Forms.Label()
        Me.chkSzenKysSorit = New System.Windows.Forms.CheckBox()
        Me.Label194 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnListSzenKysSorit = New System.Windows.Forms.Button()
        Me.Label92 = New System.Windows.Forms.Label()
        Me.Label95 = New System.Windows.Forms.Label()
        Me.Label175 = New System.Windows.Forms.Label()
        Me.lblCntSzenKysSorit = New System.Windows.Forms.Label()
        Me.pnlKojyosh = New System.Windows.Forms.Panel()
        Me.Label296 = New System.Windows.Forms.Label()
        Me.chkKojyosh = New System.Windows.Forms.CheckBox()
        Me.Label209 = New System.Windows.Forms.Label()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.btnListKojyosh = New System.Windows.Forms.Button()
        Me.Label199 = New System.Windows.Forms.Label()
        Me.Label200 = New System.Windows.Forms.Label()
        Me.Label201 = New System.Windows.Forms.Label()
        Me.lblCntKojyosh = New System.Windows.Forms.Label()
        Me.pnlBunkatumisyu = New System.Windows.Forms.Panel()
        Me.Label293 = New System.Windows.Forms.Label()
        Me.chkBunkatumisyu = New System.Windows.Forms.CheckBox()
        Me.Label207 = New System.Windows.Forms.Label()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.btnListBunkatumisyu = New System.Windows.Forms.Button()
        Me.Label195 = New System.Windows.Forms.Label()
        Me.Label196 = New System.Windows.Forms.Label()
        Me.Label197 = New System.Windows.Forms.Label()
        Me.lblCntBunkatumisyu = New System.Windows.Forms.Label()
        Me.Label208 = New System.Windows.Forms.Label()
        Me.lblHojyoDescription1 = New System.Windows.Forms.Label()
        Me.tabPageHajimeni = New System.Windows.Forms.TabPage()
        Me.lblDatacvHajimeniDescription1 = New System.Windows.Forms.Label()
        Me.lblKDatacvHajimeniLabel = New System.Windows.Forms.Label()
        Me.lblHDatacvHajimeniLabel = New System.Windows.Forms.Label()
        Me.btnDoui = New System.Windows.Forms.Button()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.pnlDcFstCmtH99 = New System.Windows.Forms.Panel()
        Me.Label301 = New System.Windows.Forms.Label()
        Me.Label302 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmtK99 = New System.Windows.Forms.Panel()
        Me.Label299 = New System.Windows.Forms.Label()
        Me.Label300 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmtH01 = New System.Windows.Forms.Panel()
        Me.Label158 = New System.Windows.Forms.Label()
        Me.Label159 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmtH02 = New System.Windows.Forms.Panel()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt11 = New System.Windows.Forms.Panel()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmtK01 = New System.Windows.Forms.Panel()
        Me.Label371 = New System.Windows.Forms.Label()
        Me.Label372 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt06 = New System.Windows.Forms.Panel()
        Me.Label364 = New System.Windows.Forms.Label()
        Me.Label369 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt05 = New System.Windows.Forms.Panel()
        Me.Label362 = New System.Windows.Forms.Label()
        Me.Label363 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt07 = New System.Windows.Forms.Panel()
        Me.Label122 = New System.Windows.Forms.Label()
        Me.Label104 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt10 = New System.Windows.Forms.Panel()
        Me.Label123 = New System.Windows.Forms.Label()
        Me.Label105 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt02 = New System.Windows.Forms.Panel()
        Me.Label127 = New System.Windows.Forms.Label()
        Me.Label133 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt08 = New System.Windows.Forms.Panel()
        Me.Label126 = New System.Windows.Forms.Label()
        Me.Label132 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt04 = New System.Windows.Forms.Panel()
        Me.Label119 = New System.Windows.Forms.Label()
        Me.Label116 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt03 = New System.Windows.Forms.Panel()
        Me.Label97 = New System.Windows.Forms.Label()
        Me.Label101 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt01 = New System.Windows.Forms.Panel()
        Me.Label129 = New System.Windows.Forms.Label()
        Me.Label135 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmtK02 = New System.Windows.Forms.Panel()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.pnlDcFstCmt09 = New System.Windows.Forms.Panel()
        Me.Label326 = New System.Windows.Forms.Label()
        Me.Label325 = New System.Windows.Forms.Label()
        Me.PictureBox14 = New System.Windows.Forms.PictureBox()
        Me.tabPageSelect = New System.Windows.Forms.TabPage()
        Me.grpExistMidToBaseMid = New System.Windows.Forms.GroupBox()
        Me.btnAllChk = New System.Windows.Forms.Button()
        Me.lblDatacvSelectCaution = New System.Windows.Forms.Label()
        Me.lblLine2 = New System.Windows.Forms.Label()
        Me.lblHidden2 = New System.Windows.Forms.Label()
        Me.grpMiddleFile = New System.Windows.Forms.GroupBox()
        Me.Label349 = New System.Windows.Forms.Label()
        Me.lblMiddleFile2 = New System.Windows.Forms.Label()
        Me.txtMidDirPath2 = New System.Windows.Forms.TextBox()
        Me.btnMidDirSeach2 = New System.Windows.Forms.Button()
        Me.btnHJizenTyukanOpen2 = New System.Windows.Forms.Button()
        Me.Label328 = New System.Windows.Forms.Label()
        Me.Label124 = New System.Windows.Forms.Label()
        Me.btnMidFileCheck = New System.Windows.Forms.Button()
        Me.btnMidDirLogSeach = New System.Windows.Forms.Button()
        Me.txtMidDirLogPath = New System.Windows.Forms.TextBox()
        Me.lblMiddleFileLog = New System.Windows.Forms.Label()
        Me.tabCtrlCVItem = New System.Windows.Forms.TabControl()
        Me.tabPageKizon110 = New System.Windows.Forms.TabPage()
        Me.grpKizon1 = New System.Windows.Forms.GroupBox()
        Me.pnlKiMstHendo = New System.Windows.Forms.Panel()
        Me.chkKiMstHendo = New System.Windows.Forms.CheckBox()
        Me.lblKiMstHendoCnt = New System.Windows.Forms.Label()
        Me.Label230 = New System.Windows.Forms.Label()
        Me.Label82 = New System.Windows.Forms.Label()
        Me.Label81 = New System.Windows.Forms.Label()
        Me.pnlKiMstTokuyaku = New System.Windows.Forms.Panel()
        Me.chkKiMstTokuyaku = New System.Windows.Forms.CheckBox()
        Me.lblKiMstTokuyakuCnt = New System.Windows.Forms.Label()
        Me.Label223 = New System.Windows.Forms.Label()
        Me.Label222 = New System.Windows.Forms.Label()
        Me.Label78 = New System.Windows.Forms.Label()
        Me.pnlKiMstKasyoClaimrui = New System.Windows.Forms.Panel()
        Me.chkKiMstKasyoClaimrui = New System.Windows.Forms.CheckBox()
        Me.lblKiMstKasyoClaimruiCnt = New System.Windows.Forms.Label()
        Me.Label228 = New System.Windows.Forms.Label()
        Me.Label227 = New System.Windows.Forms.Label()
        Me.Label226 = New System.Windows.Forms.Label()
        Me.pnlKiMstTitle = New System.Windows.Forms.Panel()
        Me.Label346 = New System.Windows.Forms.Label()
        Me.Label236 = New System.Windows.Forms.Label()
        Me.Label233 = New System.Windows.Forms.Label()
        Me.lblKiMstGazotitleCnt = New System.Windows.Forms.Label()
        Me.Label347 = New System.Windows.Forms.Label()
        Me.lblKiMstBikotitleCnt = New System.Windows.Forms.Label()
        Me.Label235 = New System.Windows.Forms.Label()
        Me.chkKiMstTitle = New System.Windows.Forms.CheckBox()
        Me.lblKiMstKagititleCnt = New System.Windows.Forms.Label()
        Me.Label232 = New System.Windows.Forms.Label()
        Me.Label80 = New System.Windows.Forms.Label()
        Me.pnlKiMstArea = New System.Windows.Forms.Panel()
        Me.chkKiMstArea = New System.Windows.Forms.CheckBox()
        Me.lblKiMstAreaCnt = New System.Windows.Forms.Label()
        Me.Label216 = New System.Windows.Forms.Label()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.Label76 = New System.Windows.Forms.Label()
        Me.pnlKiMstSchool = New System.Windows.Forms.Panel()
        Me.chkKiMstSchool = New System.Windows.Forms.CheckBox()
        Me.lblKiMstSchoolCnt = New System.Windows.Forms.Label()
        Me.Label214 = New System.Windows.Forms.Label()
        Me.Label213 = New System.Windows.Forms.Label()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.pnlKiMstHokenrui = New System.Windows.Forms.Panel()
        Me.chkKiMstHokenrui = New System.Windows.Forms.CheckBox()
        Me.lblKiMstHokenruiCnt = New System.Windows.Forms.Label()
        Me.Label220 = New System.Windows.Forms.Label()
        Me.Label219 = New System.Windows.Forms.Label()
        Me.Label218 = New System.Windows.Forms.Label()
        Me.pnlKiMstBus = New System.Windows.Forms.Panel()
        Me.chkKiMstBus = New System.Windows.Forms.CheckBox()
        Me.lblKiMstBusCnt = New System.Windows.Forms.Label()
        Me.Label211 = New System.Windows.Forms.Label()
        Me.Label181 = New System.Windows.Forms.Label()
        Me.Label210 = New System.Windows.Forms.Label()
        Me.lblSelectPageCnt1 = New System.Windows.Forms.Label()
        Me.tabPageKizon120 = New System.Windows.Forms.TabPage()
        Me.grpKizon3 = New System.Windows.Forms.GroupBox()
        Me.pnlKiGySyuzenBase = New System.Windows.Forms.Panel()
        Me.chkKiGySyuzenBase = New System.Windows.Forms.CheckBox()
        Me.lblKiGySyuzenBaseCnt = New System.Windows.Forms.Label()
        Me.Label241 = New System.Windows.Forms.Label()
        Me.Label240 = New System.Windows.Forms.Label()
        Me.pnlKiGyYatinhosyoBase = New System.Windows.Forms.Panel()
        Me.chkKiGyYatinhosyoBase = New System.Windows.Forms.CheckBox()
        Me.lblKiGyYatinhosyoBaseCnt = New System.Windows.Forms.Label()
        Me.Label250 = New System.Windows.Forms.Label()
        Me.Label249 = New System.Windows.Forms.Label()
        Me.pnlKiGyLifelineBase = New System.Windows.Forms.Panel()
        Me.chkKiGyLifelineBase = New System.Windows.Forms.CheckBox()
        Me.lblKiGyLifelineBaseCnt = New System.Windows.Forms.Label()
        Me.Label244 = New System.Windows.Forms.Label()
        Me.Label243 = New System.Windows.Forms.Label()
        Me.pnlKiGyHokenBase = New System.Windows.Forms.Panel()
        Me.chkKiGyHokenBase = New System.Windows.Forms.CheckBox()
        Me.lblKiGyHokenBaseCnt = New System.Windows.Forms.Label()
        Me.Label247 = New System.Windows.Forms.Label()
        Me.Label246 = New System.Windows.Forms.Label()
        Me.pnlKiGySisetuBase = New System.Windows.Forms.Panel()
        Me.chkKiGySisetuBase = New System.Windows.Forms.CheckBox()
        Me.lblKiGySisetuBaseCnt = New System.Windows.Forms.Label()
        Me.Label253 = New System.Windows.Forms.Label()
        Me.Label252 = New System.Windows.Forms.Label()
        Me.pnlKiGyCyukaiBase = New System.Windows.Forms.Panel()
        Me.chkKiGyCyukaiBase = New System.Windows.Forms.CheckBox()
        Me.lblKiGyCyukaiBaseCnt = New System.Windows.Forms.Label()
        Me.Label238 = New System.Windows.Forms.Label()
        Me.Label87 = New System.Windows.Forms.Label()
        Me.pnlKiGySekoBase = New System.Windows.Forms.Panel()
        Me.chkKiGySekoBase = New System.Windows.Forms.CheckBox()
        Me.lblKiGySekoBaseCnt = New System.Windows.Forms.Label()
        Me.Label256 = New System.Windows.Forms.Label()
        Me.Label255 = New System.Windows.Forms.Label()
        Me.grpKizon5 = New System.Windows.Forms.GroupBox()
        Me.pnlKiOw = New System.Windows.Forms.Panel()
        Me.chkKiOwBase = New System.Windows.Forms.CheckBox()
        Me.lblKiOwBaseCnt = New System.Windows.Forms.Label()
        Me.Label261 = New System.Windows.Forms.Label()
        Me.Label268 = New System.Windows.Forms.Label()
        Me.Label267 = New System.Windows.Forms.Label()
        Me.pnlKiJisya = New System.Windows.Forms.Panel()
        Me.chkKiJisyaBase = New System.Windows.Forms.CheckBox()
        Me.lblKiJisyaBaseCnt = New System.Windows.Forms.Label()
        Me.Label265 = New System.Windows.Forms.Label()
        Me.Label264 = New System.Windows.Forms.Label()
        Me.Label263 = New System.Windows.Forms.Label()
        Me.pnlKiSyskanriBase = New System.Windows.Forms.Panel()
        Me.chkKiSyskanriBase = New System.Windows.Forms.CheckBox()
        Me.lblKiSyskanriBaseCnt = New System.Windows.Forms.Label()
        Me.Label260 = New System.Windows.Forms.Label()
        Me.Label259 = New System.Windows.Forms.Label()
        Me.Label258 = New System.Windows.Forms.Label()
        Me.lblSelectPageCnt2 = New System.Windows.Forms.Label()
        Me.tabPageKizon130 = New System.Windows.Forms.TabPage()
        Me.grpKizon2 = New System.Windows.Forms.GroupBox()
        Me.Label359 = New System.Windows.Forms.Label()
        Me.pnlRendo = New System.Windows.Forms.Panel()
        Me.Label83 = New System.Windows.Forms.Label()
        Me.txtRendoID = New System.Windows.Forms.TextBox()
        Me.Label121 = New System.Windows.Forms.Label()
        Me.chkKiRendoBase = New System.Windows.Forms.CheckBox()
        Me.lblKiRendoBaseCnt = New System.Windows.Forms.Label()
        Me.Label314 = New System.Windows.Forms.Label()
        Me.Label316 = New System.Windows.Forms.Label()
        Me.pnlKiSq = New System.Windows.Forms.Panel()
        Me.Label308 = New System.Windows.Forms.Label()
        Me.lblKiSqOwKojoBaseCnt = New System.Windows.Forms.Label()
        Me.Label311 = New System.Windows.Forms.Label()
        Me.Label290 = New System.Windows.Forms.Label()
        Me.Label162 = New System.Windows.Forms.Label()
        Me.chkKiSqBase = New System.Windows.Forms.CheckBox()
        Me.lblKiSqMiBaseCnt = New System.Windows.Forms.Label()
        Me.Label289 = New System.Windows.Forms.Label()
        Me.Label288 = New System.Windows.Forms.Label()
        Me.lblKiSqAzBaseCnt = New System.Windows.Forms.Label()
        Me.Label160 = New System.Windows.Forms.Label()
        Me.pnlSzen = New System.Windows.Forms.Panel()
        Me.chkKiSzenBase = New System.Windows.Forms.CheckBox()
        Me.lblKiSzenBaseCnt = New System.Windows.Forms.Label()
        Me.Label295 = New System.Windows.Forms.Label()
        Me.Label294 = New System.Windows.Forms.Label()
        Me.Label174 = New System.Windows.Forms.Label()
        Me.pnlKiClaim = New System.Windows.Forms.Panel()
        Me.chkKiClaimBase = New System.Windows.Forms.CheckBox()
        Me.lblKiClaimBaseCnt = New System.Windows.Forms.Label()
        Me.Label292 = New System.Windows.Forms.Label()
        Me.Label180 = New System.Windows.Forms.Label()
        Me.pnlKiKy = New System.Windows.Forms.Panel()
        Me.chkKiKyBase = New System.Windows.Forms.CheckBox()
        Me.lblKiKyBaseCnt = New System.Windows.Forms.Label()
        Me.Label286 = New System.Windows.Forms.Label()
        Me.Label285 = New System.Windows.Forms.Label()
        Me.Label284 = New System.Windows.Forms.Label()
        Me.pnlKiKys = New System.Windows.Forms.Panel()
        Me.chkKiKysBase = New System.Windows.Forms.CheckBox()
        Me.lblKiKysBaseCnt = New System.Windows.Forms.Label()
        Me.Label282 = New System.Windows.Forms.Label()
        Me.Label281 = New System.Windows.Forms.Label()
        Me.Label280 = New System.Windows.Forms.Label()
        Me.pnlKiHy = New System.Windows.Forms.Panel()
        Me.chkKiHyBase = New System.Windows.Forms.CheckBox()
        Me.lblKiHyBaseCnt = New System.Windows.Forms.Label()
        Me.Label275 = New System.Windows.Forms.Label()
        Me.Label270 = New System.Windows.Forms.Label()
        Me.chkKiHySetubi = New System.Windows.Forms.CheckBox()
        Me.lblKiHySetubiCnt = New System.Windows.Forms.Label()
        Me.Label278 = New System.Windows.Forms.Label()
        Me.Label277 = New System.Windows.Forms.Label()
        Me.pnlKiBk = New System.Windows.Forms.Panel()
        Me.chkKiBkBase = New System.Windows.Forms.CheckBox()
        Me.lblKiBkBaseCnt = New System.Windows.Forms.Label()
        Me.Label273 = New System.Windows.Forms.Label()
        Me.Label271 = New System.Windows.Forms.Label()
        Me.grpKizonKagi = New System.Windows.Forms.GroupBox()
        Me.optKiKyKagi = New System.Windows.Forms.RadioButton()
        Me.optKiHyKagi = New System.Windows.Forms.RadioButton()
        Me.Label86 = New System.Windows.Forms.Label()
        Me.Label272 = New System.Windows.Forms.Label()
        Me.tabPageBase110 = New System.Windows.Forms.TabPage()
        Me.lblRelItemInfo = New System.Windows.Forms.Label()
        Me.grpMst = New System.Windows.Forms.GroupBox()
        Me.chkMstGazotitle = New System.Windows.Forms.CheckBox()
        Me.chkMstBikolst = New System.Windows.Forms.CheckBox()
        Me.chkMstBikotitle = New System.Windows.Forms.CheckBox()
        Me.chkMstHendoitiran = New System.Windows.Forms.CheckBox()
        Me.chkMstHendo = New System.Windows.Forms.CheckBox()
        Me.chkMstKasyoClaimrui = New System.Windows.Forms.CheckBox()
        Me.chkMstBus = New System.Windows.Forms.CheckBox()
        Me.chkMstSchool = New System.Windows.Forms.CheckBox()
        Me.chkMstTokuyaku = New System.Windows.Forms.CheckBox()
        Me.chkMstHokenrui = New System.Windows.Forms.CheckBox()
        Me.chkMstBusKotu = New System.Windows.Forms.CheckBox()
        Me.chkMstKagititle = New System.Windows.Forms.CheckBox()
        Me.chkMstArea = New System.Windows.Forms.CheckBox()
        Me.tabPageBase120 = New System.Windows.Forms.TabPage()
        Me.grpGy = New System.Windows.Forms.GroupBox()
        Me.chkGySyuzenMemo = New System.Windows.Forms.CheckBox()
        Me.chkGyHokenMemo = New System.Windows.Forms.CheckBox()
        Me.chkGySyuzenKoza = New System.Windows.Forms.CheckBox()
        Me.chkGySekoBase = New System.Windows.Forms.CheckBox()
        Me.chkGySisetuBase = New System.Windows.Forms.CheckBox()
        Me.chkGySyuzenBase = New System.Windows.Forms.CheckBox()
        Me.chkGyYatinhosyoMemo = New System.Windows.Forms.CheckBox()
        Me.chkGyYatinhosyoBase = New System.Windows.Forms.CheckBox()
        Me.chkGyCyukaiKoza = New System.Windows.Forms.CheckBox()
        Me.chkGyCyukaiMemo = New System.Windows.Forms.CheckBox()
        Me.chkGyHokenBase = New System.Windows.Forms.CheckBox()
        Me.chkGyLifelineBase = New System.Windows.Forms.CheckBox()
        Me.chkGyHokenKoza = New System.Windows.Forms.CheckBox()
        Me.chkGyCyukaiBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase130 = New System.Windows.Forms.TabPage()
        Me.grpKys = New System.Windows.Forms.GroupBox()
        Me.chkKysHosyonin = New System.Windows.Forms.CheckBox()
        Me.chkKysSyogoKana = New System.Windows.Forms.CheckBox()
        Me.chkKysKoza = New System.Windows.Forms.CheckBox()
        Me.chkKysMemo = New System.Windows.Forms.CheckBox()
        Me.chkKysBase = New System.Windows.Forms.CheckBox()
        Me.grpOw = New System.Windows.Forms.GroupBox()
        Me.chkOwEvent = New System.Windows.Forms.CheckBox()
        Me.chkOwKoza = New System.Windows.Forms.CheckBox()
        Me.chkOwMemo = New System.Windows.Forms.CheckBox()
        Me.chkOwBase = New System.Windows.Forms.CheckBox()
        Me.grpJisya = New System.Windows.Forms.GroupBox()
        Me.chkFBANSERSetuzoku = New System.Windows.Forms.CheckBox()
        Me.chkMstANSERArea = New System.Windows.Forms.CheckBox()
        Me.chkMstANSERAccpoint = New System.Windows.Forms.CheckBox()
        Me.chkJisyaTanto = New System.Windows.Forms.CheckBox()
        Me.chkMstYatinKoza = New System.Windows.Forms.CheckBox()
        Me.chkJisyaKoza = New System.Windows.Forms.CheckBox()
        Me.chkJisyaBase = New System.Windows.Forms.CheckBox()
        Me.chkFBFuriirai = New System.Windows.Forms.CheckBox()
        Me.chkFBNsSyutoku = New System.Windows.Forms.CheckBox()
        Me.chkJisyaMemo = New System.Windows.Forms.CheckBox()
        Me.chkFBKozafurikae = New System.Windows.Forms.CheckBox()
        Me.chkFBFuritesuryo = New System.Windows.Forms.CheckBox()
        Me.tabPageBase140 = New System.Windows.Forms.TabPage()
        Me.grpBk = New System.Windows.Forms.GroupBox()
        Me.chkOpKys = New System.Windows.Forms.CheckBox()
        Me.chkOpOw = New System.Windows.Forms.CheckBox()
        Me.grpKagiSelect = New System.Windows.Forms.GroupBox()
        Me.optKyKagi = New System.Windows.Forms.RadioButton()
        Me.optHyKagi = New System.Windows.Forms.RadioButton()
        Me.chkBkSyo = New System.Windows.Forms.CheckBox()
        Me.chkBkHendo = New System.Windows.Forms.CheckBox()
        Me.chkBkKinrincyusyajo = New System.Windows.Forms.CheckBox()
        Me.chkBkSansyofile = New System.Windows.Forms.CheckBox()
        Me.chkBkSzeniji = New System.Windows.Forms.CheckBox()
        Me.chkBkSyuhen = New System.Windows.Forms.CheckBox()
        Me.chkBkSetudo = New System.Windows.Forms.CheckBox()
        Me.chkBkKotu = New System.Windows.Forms.CheckBox()
        Me.chkBkKenri = New System.Windows.Forms.CheckBox()
        Me.chkBkGomi = New System.Windows.Forms.CheckBox()
        Me.chkBkSyosai = New System.Windows.Forms.CheckBox()
        Me.chkBkKagi = New System.Windows.Forms.CheckBox()
        Me.chkBkMemo = New System.Windows.Forms.CheckBox()
        Me.chkBkBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase150 = New System.Windows.Forms.TabPage()
        Me.grpHy = New System.Windows.Forms.GroupBox()
        Me.chkHySyo = New System.Windows.Forms.CheckBox()
        Me.chkHySansyofile = New System.Windows.Forms.CheckBox()
        Me.chkHyGenjotanka = New System.Windows.Forms.CheckBox()
        Me.chkHyKenri = New System.Windows.Forms.CheckBox()
        Me.chkHyConfirm = New System.Windows.Forms.CheckBox()
        Me.chkHyHendo = New System.Windows.Forms.CheckBox()
        Me.chkHyCommonsalespoint = New System.Windows.Forms.CheckBox()
        Me.chkHyMenseki = New System.Windows.Forms.CheckBox()
        Me.chkHyNkinkomk = New System.Windows.Forms.CheckBox()
        Me.chkHyMadoriutiwake = New System.Windows.Forms.CheckBox()
        Me.chkHySzeniji = New System.Windows.Forms.CheckBox()
        Me.chkHyTokuyaku = New System.Windows.Forms.CheckBox()
        Me.chkHyParking = New System.Windows.Forms.CheckBox()
        Me.chkHySyosai = New System.Windows.Forms.CheckBox()
        Me.chkHyMemo = New System.Windows.Forms.CheckBox()
        Me.chkHyKagi = New System.Windows.Forms.CheckBox()
        Me.chkHySetubi = New System.Windows.Forms.CheckBox()
        Me.chkHyBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase160 = New System.Windows.Forms.TabPage()
        Me.grpSorule = New System.Windows.Forms.GroupBox()
        Me.chkSoruleSosaki = New System.Windows.Forms.CheckBox()
        Me.chkSoruleKojo = New System.Windows.Forms.CheckBox()
        Me.chkSoruleNkin = New System.Windows.Forms.CheckBox()
        Me.chkSoruleBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase170 = New System.Windows.Forms.TabPage()
        Me.grpKy = New System.Windows.Forms.GroupBox()
        Me.chkKySzenmeisai = New System.Windows.Forms.CheckBox()
        Me.chkKySzen = New System.Windows.Forms.CheckBox()
        Me.chkKyKai = New System.Windows.Forms.CheckBox()
        Me.CheckBox14 = New System.Windows.Forms.CheckBox()
        Me.CheckBox24 = New System.Windows.Forms.CheckBox()
        Me.CheckBox23 = New System.Windows.Forms.CheckBox()
        Me.chkKyHosyonin = New System.Windows.Forms.CheckBox()
        Me.chkKyMemo = New System.Windows.Forms.CheckBox()
        Me.chkKyTokuyaku = New System.Windows.Forms.CheckBox()
        Me.CheckBox18 = New System.Windows.Forms.CheckBox()
        Me.chkKySorule = New System.Windows.Forms.CheckBox()
        Me.CheckBox16 = New System.Windows.Forms.CheckBox()
        Me.CheckBox15 = New System.Windows.Forms.CheckBox()
        Me.chkKyNyukyo = New System.Windows.Forms.CheckBox()
        Me.chkKyNkinkomkNx = New System.Windows.Forms.CheckBox()
        Me.chkKyNkinkomk = New System.Windows.Forms.CheckBox()
        Me.chkKyHoken = New System.Windows.Forms.CheckBox()
        Me.chkKyKojoRule = New System.Windows.Forms.CheckBox()
        Me.chkKyKys = New System.Windows.Forms.CheckBox()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.chkKyRireki = New System.Windows.Forms.CheckBox()
        Me.chkKyHendo = New System.Windows.Forms.CheckBox()
        Me.chkKyCar = New System.Windows.Forms.CheckBox()
        Me.chkKyBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase180 = New System.Windows.Forms.TabPage()
        Me.grpSq = New System.Windows.Forms.GroupBox()
        Me.chkSqSqKojo = New System.Windows.Forms.CheckBox()
        Me.chkSqKoteiKojo = New System.Windows.Forms.CheckBox()
        Me.chkSqHendokensin = New System.Windows.Forms.CheckBox()
        Me.chkSqUnyotaino = New System.Windows.Forms.CheckBox()
        Me.chkSqSq = New System.Windows.Forms.CheckBox()
        Me.chkSqKajyo = New System.Windows.Forms.CheckBox()
        Me.tabPageBase190 = New System.Windows.Forms.TabPage()
        Me.grpSzen = New System.Windows.Forms.GroupBox()
        Me.chkSzenRelfile = New System.Windows.Forms.CheckBox()
        Me.chkSzenClaim = New System.Windows.Forms.CheckBox()
        Me.chkSzenMemo = New System.Windows.Forms.CheckBox()
        Me.chkSzenSzen = New System.Windows.Forms.CheckBox()
        Me.chkSzenSzenmeisai = New System.Windows.Forms.CheckBox()
        Me.chkSzenBase = New System.Windows.Forms.CheckBox()
        Me.grpClaim = New System.Windows.Forms.GroupBox()
        Me.chkClaimTaiorireki = New System.Windows.Forms.CheckBox()
        Me.chkClaimRelfile = New System.Windows.Forms.CheckBox()
        Me.chkClaimBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase200 = New System.Windows.Forms.TabPage()
        Me.grpSyskanri = New System.Windows.Forms.GroupBox()
        Me.chkSyskanriNkinkomkmerge = New System.Windows.Forms.CheckBox()
        Me.chkSyskanriHenkanmoji = New System.Windows.Forms.CheckBox()
        Me.chkSyskanriZei = New System.Windows.Forms.CheckBox()
        Me.chkSyskanriBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase210 = New System.Windows.Forms.TabPage()
        Me.grpRendo = New System.Windows.Forms.GroupBox()
        Me.chkRendoMapdisp = New System.Windows.Forms.CheckBox()
        Me.chkRendoBtoBgroup = New System.Windows.Forms.CheckBox()
        Me.chkRendoHysosin = New System.Windows.Forms.CheckBox()
        Me.chkRendoHyrui = New System.Windows.Forms.CheckBox()
        Me.chkRendoKokokuSuumo = New System.Windows.Forms.CheckBox()
        Me.chkRendoKokokuAthome = New System.Windows.Forms.CheckBox()
        Me.chkRendoKokokuHomes = New System.Windows.Forms.CheckBox()
        Me.chkRendoKokokuJisyaweb = New System.Windows.Forms.CheckBox()
        Me.chkRendoSosinSuumo = New System.Windows.Forms.CheckBox()
        Me.chkRendoSosinAthome = New System.Windows.Forms.CheckBox()
        Me.chkRendoSosinHomes = New System.Windows.Forms.CheckBox()
        Me.chkRendoSosinJisyaweb = New System.Windows.Forms.CheckBox()
        Me.chkRendoSosinBase = New System.Windows.Forms.CheckBox()
        Me.tabPageBase900 = New System.Windows.Forms.TabPage()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.tabPageHanyo110 = New System.Windows.Forms.TabPage()
        Me.grpHMst = New System.Windows.Forms.GroupBox()
        Me.Label138 = New System.Windows.Forms.Label()
        Me.Label150 = New System.Windows.Forms.Label()
        Me.Label151 = New System.Windows.Forms.Label()
        Me.Label152 = New System.Windows.Forms.Label()
        Me.Label153 = New System.Windows.Forms.Label()
        Me.Label154 = New System.Windows.Forms.Label()
        Me.Label155 = New System.Windows.Forms.Label()
        Me.Label156 = New System.Windows.Forms.Label()
        Me.Label157 = New System.Windows.Forms.Label()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
        Me.CheckBox9 = New System.Windows.Forms.CheckBox()
        Me.CheckBox10 = New System.Windows.Forms.CheckBox()
        Me.CheckBox11 = New System.Windows.Forms.CheckBox()
        Me.CheckBox12 = New System.Windows.Forms.CheckBox()
        Me.CheckBox13 = New System.Windows.Forms.CheckBox()
        Me.CheckBox17 = New System.Windows.Forms.CheckBox()
        Me.tabPageHanyo120 = New System.Windows.Forms.TabPage()
        Me.grpHGy = New System.Windows.Forms.GroupBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.CheckBox30 = New System.Windows.Forms.CheckBox()
        Me.CheckBox31 = New System.Windows.Forms.CheckBox()
        Me.CheckBox32 = New System.Windows.Forms.CheckBox()
        Me.CheckBox33 = New System.Windows.Forms.CheckBox()
        Me.CheckBox34 = New System.Windows.Forms.CheckBox()
        Me.CheckBox35 = New System.Windows.Forms.CheckBox()
        Me.grpHOw = New System.Windows.Forms.GroupBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox26 = New System.Windows.Forms.CheckBox()
        Me.CheckBox27 = New System.Windows.Forms.CheckBox()
        Me.CheckBox28 = New System.Windows.Forms.CheckBox()
        Me.CheckBox29 = New System.Windows.Forms.CheckBox()
        Me.grpHJisya = New System.Windows.Forms.GroupBox()
        Me.CheckBox19 = New System.Windows.Forms.CheckBox()
        Me.CheckBox20 = New System.Windows.Forms.CheckBox()
        Me.CheckBox21 = New System.Windows.Forms.CheckBox()
        Me.CheckBox22 = New System.Windows.Forms.CheckBox()
        Me.CheckBox25 = New System.Windows.Forms.CheckBox()
        Me.tabPageHanyo130 = New System.Windows.Forms.TabPage()
        Me.tabPageHanyo140 = New System.Windows.Forms.TabPage()
        Me.tabPageHanyo150 = New System.Windows.Forms.TabPage()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.grpKizon4 = New System.Windows.Forms.GroupBox()
        Me.lblSelectPageCnt3 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.Label70 = New System.Windows.Forms.Label()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.chkHSongai = New System.Windows.Forms.CheckBox()
        Me.chkHJisyaKoza = New System.Windows.Forms.CheckBox()
        Me.chkHSetubi = New System.Windows.Forms.CheckBox()
        Me.chkHDataFmt = New System.Windows.Forms.CheckBox()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.Label68 = New System.Windows.Forms.Label()
        Me.Label67 = New System.Windows.Forms.Label()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.Label61 = New System.Windows.Forms.Label()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.chkHNkinKomk = New System.Windows.Forms.CheckBox()
        Me.chkHTaiyo = New System.Windows.Forms.CheckBox()
        Me.chkHKozo = New System.Windows.Forms.CheckBox()
        Me.chkHKozaSyu = New System.Windows.Forms.CheckBox()
        Me.chkHEki = New System.Windows.Forms.CheckBox()
        Me.chkHEnsen = New System.Windows.Forms.CheckBox()
        Me.chkHKinyuSiten = New System.Windows.Forms.CheckBox()
        Me.chkHYouto = New System.Windows.Forms.CheckBox()
        Me.chkHBkBunrui = New System.Windows.Forms.CheckBox()
        Me.chkHKinyu = New System.Windows.Forms.CheckBox()
        Me.chkHKyBunrui = New System.Windows.Forms.CheckBox()
        Me.chkHKagi = New System.Windows.Forms.CheckBox()
        Me.chkHHouKenri = New System.Windows.Forms.CheckBox()
        Me.chkHNkinKbn = New System.Windows.Forms.CheckBox()
        Me.chkHHyBunrui = New System.Windows.Forms.CheckBox()
        Me.pnlRekiClear = New System.Windows.Forms.Panel()
        Me.btnRekiClear = New System.Windows.Forms.Button()
        Me.Label79 = New System.Windows.Forms.Label()
        Me.Label298 = New System.Windows.Forms.Label()
        Me.lblDatacvSelectDescription1 = New System.Windows.Forms.Label()
        Me.tabPageJikko = New System.Windows.Forms.TabPage()
        Me.lblDatacvJikkoCaution = New System.Windows.Forms.Label()
        Me.lblDatacvJikkoDescription1 = New System.Windows.Forms.Label()
        Me.grpTotalProcess = New System.Windows.Forms.GroupBox()
        Me.lblCVItem = New System.Windows.Forms.Label()
        Me.txtPartialSituation = New System.Windows.Forms.TextBox()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.lblPgbPartial = New System.Windows.Forms.Label()
        Me.lblPgbTotal = New System.Windows.Forms.Label()
        Me.txtTotalSituation = New System.Windows.Forms.TextBox()
        Me.pgbpartial = New System.Windows.Forms.ProgressBar()
        Me.pgbTotal = New System.Windows.Forms.ProgressBar()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.lblTotalSituation = New System.Windows.Forms.Label()
        Me.tabPageEndOK = New System.Windows.Forms.TabPage()
        Me.lblDatacvEndDescription1 = New System.Windows.Forms.Label()
        Me.lblKDatacvEndLabel = New System.Windows.Forms.Label()
        Me.lblHDatacvEndLabel = New System.Windows.Forms.Label()
        Me.PictureBox15 = New System.Windows.Forms.PictureBox()
        Me.tabPageEndError = New System.Windows.Forms.TabPage()
        Me.GroupBox23 = New System.Windows.Forms.GroupBox()
        Me.Label163 = New System.Windows.Forms.Label()
        Me.lblDatacvErrorDescription1 = New System.Windows.Forms.Label()
        Me.GroupBox19 = New System.Windows.Forms.GroupBox()
        Me.Label170 = New System.Windows.Forms.Label()
        Me.GroupBox20 = New System.Windows.Forms.GroupBox()
        Me.Label176 = New System.Windows.Forms.Label()
        Me.lblDatacvErrorLabel = New System.Windows.Forms.Label()
        Me.PictureBox23 = New System.Windows.Forms.PictureBox()
        Me.tabPageEndCancel = New System.Windows.Forms.TabPage()
        Me.lblDatacvCancelDescription1 = New System.Windows.Forms.Label()
        Me.GroupBox21 = New System.Windows.Forms.GroupBox()
        Me.Label186 = New System.Windows.Forms.Label()
        Me.GroupBox22 = New System.Windows.Forms.GroupBox()
        Me.Label192 = New System.Windows.Forms.Label()
        Me.lblDatacvCancelLabel = New System.Windows.Forms.Label()
        Me.PictureBox25 = New System.Windows.Forms.PictureBox()
        Me.tabPageIkkatu = New System.Windows.Forms.TabPage()
        Me.Label114 = New System.Windows.Forms.Label()
        Me.Label144 = New System.Windows.Forms.Label()
        Me.tabPageHanyoJizen = New System.Windows.Forms.TabPage()
        Me.Label139 = New System.Windows.Forms.Label()
        Me.Label145 = New System.Windows.Forms.Label()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.Label140 = New System.Windows.Forms.Label()
        Me.Label141 = New System.Windows.Forms.Label()
        Me.Label128 = New System.Windows.Forms.Label()
        Me.Label134 = New System.Windows.Forms.Label()
        Me.pnlHJizenGazoKeisiki = New System.Windows.Forms.Panel()
        Me.Label343 = New System.Windows.Forms.Label()
        Me.chkHJizenGazoKeisiki = New System.Windows.Forms.CheckBox()
        Me.Label344 = New System.Windows.Forms.Label()
        Me.pnlJizenSo = New System.Windows.Forms.Panel()
        Me.chkJizenSo = New System.Windows.Forms.CheckBox()
        Me.Label90 = New System.Windows.Forms.Label()
        Me.btnDevTabChange = New System.Windows.Forms.Button()
        Me.Label393 = New System.Windows.Forms.Label()
        Me.Label392 = New System.Windows.Forms.Label()
        Me.CheckBox51 = New System.Windows.Forms.CheckBox()
        Me.Label391 = New System.Windows.Forms.Label()
        Me.CheckBox50 = New System.Windows.Forms.CheckBox()
        Me.Panel24 = New System.Windows.Forms.Panel()
        Me.Label390 = New System.Windows.Forms.Label()
        Me.Label389 = New System.Windows.Forms.Label()
        Me.Label388 = New System.Windows.Forms.Label()
        Me.Label387 = New System.Windows.Forms.Label()
        Me.Panel23 = New System.Windows.Forms.Panel()
        Me.Label386 = New System.Windows.Forms.Label()
        Me.Label385 = New System.Windows.Forms.Label()
        Me.Label384 = New System.Windows.Forms.Label()
        Me.Label383 = New System.Windows.Forms.Label()
        Me.Label382 = New System.Windows.Forms.Label()
        Me.CheckBox49 = New System.Windows.Forms.CheckBox()
        Me.Panel22 = New System.Windows.Forms.Panel()
        Me.Label381 = New System.Windows.Forms.Label()
        Me.Label380 = New System.Windows.Forms.Label()
        Me.Label379 = New System.Windows.Forms.Label()
        Me.Label378 = New System.Windows.Forms.Label()
        Me.pnlMenuMain = New System.Windows.Forms.Panel()
        Me.pnlPrgMStart = New System.Windows.Forms.Panel()
        Me.picIcoMStart2 = New System.Windows.Forms.PictureBox()
        Me.picIcoMStart1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.lblPrgMStart = New System.Windows.Forms.Label()
        Me.picArwMSession2 = New System.Windows.Forms.PictureBox()
        Me.picArwMSession1 = New System.Windows.Forms.PictureBox()
        Me.pnlPrgMEnd = New System.Windows.Forms.Panel()
        Me.picIcoMEnd2 = New System.Windows.Forms.PictureBox()
        Me.picIcoMEnd1 = New System.Windows.Forms.PictureBox()
        Me.picIcoJizen1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgMEnd = New System.Windows.Forms.Label()
        Me.pnlPrgMMenu = New System.Windows.Forms.Panel()
        Me.picIcoMMenu2 = New System.Windows.Forms.PictureBox()
        Me.picIcoMMenu1 = New System.Windows.Forms.PictureBox()
        Me.picIcoSession1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgMMenu = New System.Windows.Forms.Label()
        Me.pnlPrgMSyoki = New System.Windows.Forms.Panel()
        Me.picIcoMSyoki2 = New System.Windows.Forms.PictureBox()
        Me.picIcoMSyoki1 = New System.Windows.Forms.PictureBox()
        Me.picIcoSyoki1 = New System.Windows.Forms.PictureBox()
        Me.lblPrgMSyoki = New System.Windows.Forms.Label()
        Me.pnlPrgMSession = New System.Windows.Forms.Panel()
        Me.picIcoMSession2 = New System.Windows.Forms.PictureBox()
        Me.picIcoMSession1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.lblPrgMSession = New System.Windows.Forms.Label()
        Me.picArwMEnd2 = New System.Windows.Forms.PictureBox()
        Me.picArwMMenu2 = New System.Windows.Forms.PictureBox()
        Me.picArwMSyoki2 = New System.Windows.Forms.PictureBox()
        Me.picArwMEnd1 = New System.Windows.Forms.PictureBox()
        Me.picArwMMenu1 = New System.Windows.Forms.PictureBox()
        Me.picArwMSyoki1 = New System.Windows.Forms.PictureBox()
        Me.lblLine0 = New System.Windows.Forms.Label()
        Me.pnlRefresh = New System.Windows.Forms.Panel()
        Me.Label251 = New System.Windows.Forms.Label()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.PictureBox24 = New System.Windows.Forms.PictureBox()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.Button21 = New System.Windows.Forms.Button()
        Me.Button20 = New System.Windows.Forms.Button()
        Me.Button19 = New System.Windows.Forms.Button()
        Me.lblTitleH = New System.Windows.Forms.Label()
        Me.pgbCheck = New System.Windows.Forms.ProgressBar()
        Me.lblPgbCheck = New System.Windows.Forms.Label()
        Me.lblCheckSituation = New System.Windows.Forms.Label()
        Me.pnlPrgChk = New System.Windows.Forms.Panel()
        Me.pnlMenuDatacv.SuspendLayout
        Me.pnlPrgDCConv.SuspendLayout
        CType(Me.picIcoDCConv2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoDCConv1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgDCRelation.SuspendLayout
        CType(Me.picIcoDCRelation2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoDCRelation1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgDCFileWrite.SuspendLayout
        CType(Me.picIcoDCFileWrite2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoDCFileWrite1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwDCEnd2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwDCJikko2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwDCSelect2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwDCEnd1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwDCJikko1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwDCSelect1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgDCEnd.SuspendLayout
        CType(Me.picIcoDCEnd2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoDCEnd1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgDCJikko.SuspendLayout
        CType(Me.picIcoDCJikko2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoDCJikko1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgDCSelect.SuspendLayout
        CType(Me.picIcoDCSelect2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoDCSelect1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgDCHajimeni.SuspendLayout
        CType(Me.picIcoDCHajimeni2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoDCHajimeni1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.tabCtrlMain.SuspendLayout
        Me.tabPageDev.SuspendLayout
        Me.GroupBox6.SuspendLayout
        Me.grpDevSettingX.SuspendLayout
        CType(Me.PictureBox1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.grpDevSetting1.SuspendLayout
        Me.grpDevSetting2.SuspendLayout
        Me.grpDevSetting3.SuspendLayout
        Me.GroupBox3.SuspendLayout
        Me.GroupBox2.SuspendLayout
        Me.GroupBox1.SuspendLayout
        Me.grptaihi.SuspendLayout
        Me.tabPageStart.SuspendLayout
        Me.grpKFirstNaiyo.SuspendLayout
        Me.grpHFirstNaiyo.SuspendLayout
        CType(Me.PictureBox2,System.ComponentModel.ISupportInitialize).BeginInit
        Me.tabPageSession.SuspendLayout
        Me.grpTimeOut.SuspendLayout
        Me.grp10ConnectInfo.SuspendLayout
        Me.grpV10Authent.SuspendLayout
        Me.grpV7ConnectInfo.SuspendLayout
        Me.grpV7Authent.SuspendLayout
        Me.tabPageSyoki.SuspendLayout
        Me.pnlConvertType.SuspendLayout
        Me.pnlCVType.SuspendLayout
        Me.pnlUnyoKaisi.SuspendLayout
        Me.pnlOptSelect.SuspendLayout
        Me.pnlOptionSelect.SuspendLayout
        Me.pnlUserName.SuspendLayout
        Me.pnlLogPath.SuspendLayout
        Me.tabPageMenu.SuspendLayout
        Me.grpMenuHojyo.SuspendLayout
        Me.grpMenuJizen.SuspendLayout
        Me.grpMenuDatacv.SuspendLayout
        Me.grpMenuGazocv.SuspendLayout
        Me.grpMenuJigo.SuspendLayout
        Me.tabPageJizen.SuspendLayout
        Me.pnlJizenListPath.SuspendLayout
        Me.tabCtrlJizen.SuspendLayout
        Me.tabPageJizen1.SuspendLayout
        Me.grpJizen1.SuspendLayout
        Me.pnlJizenHurikae.SuspendLayout
        Me.pnlJizenAzu.SuspendLayout
        Me.Panel7.SuspendLayout
        Me.pnlJizenKai.SuspendLayout
        Me.Panel6.SuspendLayout
        Me.tabPageJizen2.SuspendLayout
        Me.grpJizen2.SuspendLayout
        Me.Panel12.SuspendLayout
        Me.Panel5.SuspendLayout
        Me.pnlJizenHasseiOw.SuspendLayout
        Me.Panel3.SuspendLayout
        Me.pnlJizenYanuso.SuspendLayout
        Me.Panel8.SuspendLayout
        Me.pnlJizenBkhourei.SuspendLayout
        Me.Panel4.SuspendLayout
        Me.tabPageJizen3.SuspendLayout
        Me.grpJizen3.SuspendLayout
        Me.pnlJizenKagi.SuspendLayout
        Me.Panel19.SuspendLayout
        Me.pnlJizenSzenKyshutan.SuspendLayout
        Me.Panel11.SuspendLayout
        Me.tabPageJizen4.SuspendLayout
        Me.grpJizen4.SuspendLayout
        Me.Panel15.SuspendLayout
        Me.Panel20.SuspendLayout
        Me.Panel21.SuspendLayout
        Me.Panel27.SuspendLayout
        Me.tabPageJizen5.SuspendLayout
        Me.GroupBox4.SuspendLayout
        Me.Panel13.SuspendLayout
        Me.Panel14.SuspendLayout
        Me.tabPageHJizen1.SuspendLayout
        Me.GroupBox5.SuspendLayout
        Me.pnlHJizenTyukan.SuspendLayout
        Me.Panel26.SuspendLayout
        Me.tabPageJigo.SuspendLayout
        Me.pnlJigoListPath.SuspendLayout
        Me.tabCtrlJigo.SuspendLayout
        Me.tabPageJigo1.SuspendLayout
        Me.grpJigoUserSagyo.SuspendLayout
        Me.pnlKJigoCmtSyudo.SuspendLayout
        Me.pnlHJigoCmtSyudo.SuspendLayout
        Me.pnlJigoCmtSyusi.SuspendLayout
        Me.pnlJizenMinus.SuspendLayout
        Me.Panel2.SuspendLayout
        Me.pnlJigoCmtCsvSyuturyoku.SuspendLayout
        Me.pnlJigoCmtSyosiki.SuspendLayout
        Me.tabPageJigo2.SuspendLayout
        Me.grpJigoCmtTyuui.SuspendLayout
        Me.pnlJigoCmtHeiko.SuspendLayout
        Me.grpJigoDonyuji.SuspendLayout
        Me.pnlJigoCmtSoKotiku.SuspendLayout
        Me.pnlJigoCmtNkNyuryoku.SuspendLayout
        Me.pnlJigoCmtSqKotiku.SuspendLayout
        Me.tabPageHojyo.SuspendLayout
        Me.pnlKensyoListPath.SuspendLayout
        Me.tabCtrlHojyo.SuspendLayout
        Me.tabPageHojyo1.SuspendLayout
        Me.GroupBox27.SuspendLayout
        Me.pnlKiOpKys.SuspendLayout
        Me.Panel18.SuspendLayout
        Me.pnlKiOpOw.SuspendLayout
        Me.Panel16.SuspendLayout
        Me.pnlRelRename.SuspendLayout
        Me.Panel25.SuspendLayout
        Me.tabPageHojyo2.SuspendLayout
        Me.GroupBox25.SuspendLayout
        Me.pnlRelationSet.SuspendLayout
        Me.Panel17.SuspendLayout
        Me.pnlSzenKysSorit.SuspendLayout
        Me.Panel1.SuspendLayout
        Me.pnlKojyosh.SuspendLayout
        Me.Panel10.SuspendLayout
        Me.pnlBunkatumisyu.SuspendLayout
        Me.Panel9.SuspendLayout
        Me.tabPageHajimeni.SuspendLayout
        Me.GroupBox10.SuspendLayout
        Me.pnlDcFstCmtH99.SuspendLayout
        Me.pnlDcFstCmtK99.SuspendLayout
        Me.pnlDcFstCmtH01.SuspendLayout
        Me.pnlDcFstCmtH02.SuspendLayout
        Me.pnlDcFstCmt11.SuspendLayout
        Me.pnlDcFstCmtK01.SuspendLayout
        Me.pnlDcFstCmt06.SuspendLayout
        Me.pnlDcFstCmt05.SuspendLayout
        Me.pnlDcFstCmt07.SuspendLayout
        Me.pnlDcFstCmt10.SuspendLayout
        Me.pnlDcFstCmt02.SuspendLayout
        Me.pnlDcFstCmt08.SuspendLayout
        Me.pnlDcFstCmt04.SuspendLayout
        Me.pnlDcFstCmt03.SuspendLayout
        Me.pnlDcFstCmt01.SuspendLayout
        Me.pnlDcFstCmtK02.SuspendLayout
        Me.pnlDcFstCmt09.SuspendLayout
        CType(Me.PictureBox14,System.ComponentModel.ISupportInitialize).BeginInit
        Me.tabPageSelect.SuspendLayout
        Me.grpMiddleFile.SuspendLayout
        Me.tabCtrlCVItem.SuspendLayout
        Me.tabPageKizon110.SuspendLayout
        Me.grpKizon1.SuspendLayout
        Me.pnlKiMstHendo.SuspendLayout
        Me.pnlKiMstTokuyaku.SuspendLayout
        Me.pnlKiMstKasyoClaimrui.SuspendLayout
        Me.pnlKiMstTitle.SuspendLayout
        Me.pnlKiMstArea.SuspendLayout
        Me.pnlKiMstSchool.SuspendLayout
        Me.pnlKiMstHokenrui.SuspendLayout
        Me.pnlKiMstBus.SuspendLayout
        Me.tabPageKizon120.SuspendLayout
        Me.grpKizon3.SuspendLayout
        Me.pnlKiGySyuzenBase.SuspendLayout
        Me.pnlKiGyYatinhosyoBase.SuspendLayout
        Me.pnlKiGyLifelineBase.SuspendLayout
        Me.pnlKiGyHokenBase.SuspendLayout
        Me.pnlKiGySisetuBase.SuspendLayout
        Me.pnlKiGyCyukaiBase.SuspendLayout
        Me.pnlKiGySekoBase.SuspendLayout
        Me.grpKizon5.SuspendLayout
        Me.pnlKiOw.SuspendLayout
        Me.pnlKiJisya.SuspendLayout
        Me.pnlKiSyskanriBase.SuspendLayout
        Me.tabPageKizon130.SuspendLayout
        Me.grpKizon2.SuspendLayout
        Me.pnlRendo.SuspendLayout
        Me.pnlKiSq.SuspendLayout
        Me.pnlSzen.SuspendLayout
        Me.pnlKiClaim.SuspendLayout
        Me.pnlKiKy.SuspendLayout
        Me.pnlKiKys.SuspendLayout
        Me.pnlKiHy.SuspendLayout
        Me.pnlKiBk.SuspendLayout
        Me.grpKizonKagi.SuspendLayout
        Me.tabPageBase110.SuspendLayout
        Me.grpMst.SuspendLayout
        Me.tabPageBase120.SuspendLayout
        Me.grpGy.SuspendLayout
        Me.tabPageBase130.SuspendLayout
        Me.grpKys.SuspendLayout
        Me.grpOw.SuspendLayout
        Me.grpJisya.SuspendLayout
        Me.tabPageBase140.SuspendLayout
        Me.grpBk.SuspendLayout
        Me.grpKagiSelect.SuspendLayout
        Me.tabPageBase150.SuspendLayout
        Me.grpHy.SuspendLayout
        Me.tabPageBase160.SuspendLayout
        Me.grpSorule.SuspendLayout
        Me.tabPageBase170.SuspendLayout
        Me.grpKy.SuspendLayout
        Me.tabPageBase180.SuspendLayout
        Me.grpSq.SuspendLayout
        Me.tabPageBase190.SuspendLayout
        Me.grpSzen.SuspendLayout
        Me.grpClaim.SuspendLayout
        Me.tabPageBase200.SuspendLayout
        Me.grpSyskanri.SuspendLayout
        Me.tabPageBase210.SuspendLayout
        Me.grpRendo.SuspendLayout
        Me.tabPageBase900.SuspendLayout
        Me.tabPageHanyo110.SuspendLayout
        Me.grpHMst.SuspendLayout
        Me.tabPageHanyo120.SuspendLayout
        Me.grpHGy.SuspendLayout
        Me.grpHOw.SuspendLayout
        Me.grpHJisya.SuspendLayout
        Me.TabPage1.SuspendLayout
        Me.grpKizon4.SuspendLayout
        Me.pnlRekiClear.SuspendLayout
        Me.tabPageJikko.SuspendLayout
        Me.grpTotalProcess.SuspendLayout
        Me.tabPageEndOK.SuspendLayout
        CType(Me.PictureBox15,System.ComponentModel.ISupportInitialize).BeginInit
        Me.tabPageEndError.SuspendLayout
        Me.GroupBox23.SuspendLayout
        Me.GroupBox19.SuspendLayout
        Me.GroupBox20.SuspendLayout
        CType(Me.PictureBox23,System.ComponentModel.ISupportInitialize).BeginInit
        Me.tabPageEndCancel.SuspendLayout
        Me.GroupBox21.SuspendLayout
        Me.GroupBox22.SuspendLayout
        CType(Me.PictureBox25,System.ComponentModel.ISupportInitialize).BeginInit
        Me.tabPageIkkatu.SuspendLayout
        Me.tabPageHanyoJizen.SuspendLayout
        Me.GroupBox11.SuspendLayout
        Me.pnlHJizenGazoKeisiki.SuspendLayout
        Me.pnlJizenSo.SuspendLayout
        Me.pnlMenuMain.SuspendLayout
        Me.pnlPrgMStart.SuspendLayout
        CType(Me.picIcoMStart2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoMStart1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.PictureBox3,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMSession2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMSession1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgMEnd.SuspendLayout
        CType(Me.picIcoMEnd2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoMEnd1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoJizen1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgMMenu.SuspendLayout
        CType(Me.picIcoMMenu2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoMMenu1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoSession1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgMSyoki.SuspendLayout
        CType(Me.picIcoMSyoki2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoMSyoki1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoSyoki1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgMSession.SuspendLayout
        CType(Me.picIcoMSession2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picIcoMSession1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.PictureBox4,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMEnd2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMMenu2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMSyoki2,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMEnd1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMMenu1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.picArwMSyoki1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlRefresh.SuspendLayout
        CType(Me.PictureBox24,System.ComponentModel.ISupportInitialize).BeginInit
        Me.pnlPrgChk.SuspendLayout
        Me.SuspendLayout
        '
        'chkRelationFile
        '
        Me.chkRelationFile.AutoSize = true
        Me.chkRelationFile.Font = New System.Drawing.Font("メイリオ", 6.75!, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkRelationFile.Location = New System.Drawing.Point(8, 126)
        Me.chkRelationFile.Name = "chkRelationFile"
        Me.chkRelationFile.Size = New System.Drawing.Size(125, 18)
        Me.chkRelationFile.TabIndex = 110
        Me.chkRelationFile.Text = "紐付けファイル作成済み"
        Me.chkRelationFile.UseVisualStyleBackColor = true
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("メイリオ", 6.75!, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label9.Location = New System.Drawing.Point(8, 147)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(211, 31)
        Me.Label9.TabIndex = 116
        Me.Label9.Text = "紐付け設定を仮TBLへ登録済みの場合にチェックを付けて下さい。"
        Me.Label9.UseCompatibleTextRendering = true
        '
        'lblTitleKi
        '
        Me.lblTitleKi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTitleKi.Font = New System.Drawing.Font("メイリオ", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic),System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblTitleKi.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.lblTitleKi.Location = New System.Drawing.Point(33, 18)
        Me.lblTitleKi.Name = "lblTitleKi"
        Me.lblTitleKi.Size = New System.Drawing.Size(250, 26)
        Me.lblTitleKi.TabIndex = 0
        Me.lblTitleKi.Text = "[賃貸革命V7 → 10]"
        Me.lblTitleKi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlMenuDatacv
        '
        Me.pnlMenuDatacv.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlMenuDatacv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlMenuDatacv.Controls.Add(Me.pnlPrgDCConv)
        Me.pnlMenuDatacv.Controls.Add(Me.pnlPrgDCRelation)
        Me.pnlMenuDatacv.Controls.Add(Me.pnlPrgDCFileWrite)
        Me.pnlMenuDatacv.Controls.Add(Me.picArwDCEnd2)
        Me.pnlMenuDatacv.Controls.Add(Me.picArwDCJikko2)
        Me.pnlMenuDatacv.Controls.Add(Me.picArwDCSelect2)
        Me.pnlMenuDatacv.Controls.Add(Me.picArwDCEnd1)
        Me.pnlMenuDatacv.Controls.Add(Me.picArwDCJikko1)
        Me.pnlMenuDatacv.Controls.Add(Me.picArwDCSelect1)
        Me.pnlMenuDatacv.Controls.Add(Me.pnlPrgDCEnd)
        Me.pnlMenuDatacv.Controls.Add(Me.pnlPrgDCJikko)
        Me.pnlMenuDatacv.Controls.Add(Me.pnlPrgDCSelect)
        Me.pnlMenuDatacv.Controls.Add(Me.pnlPrgDCHajimeni)
        Me.pnlMenuDatacv.Location = New System.Drawing.Point(33, 60)
        Me.pnlMenuDatacv.Name = "pnlMenuDatacv"
        Me.pnlMenuDatacv.Size = New System.Drawing.Size(250, 561)
        Me.pnlMenuDatacv.TabIndex = 1
        '
        'pnlPrgDCConv
        '
        Me.pnlPrgDCConv.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgDCConv.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgDCConv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgDCConv.CausesValidation = false
        Me.pnlPrgDCConv.Controls.Add(Me.picIcoDCConv2)
        Me.pnlPrgDCConv.Controls.Add(Me.picIcoDCConv1)
        Me.pnlPrgDCConv.Controls.Add(Me.lblPrgDCConv)
        Me.pnlPrgDCConv.Location = New System.Drawing.Point(31, 336)
        Me.pnlPrgDCConv.Name = "pnlPrgDCConv"
        Me.pnlPrgDCConv.Size = New System.Drawing.Size(184, 33)
        Me.pnlPrgDCConv.TabIndex = 9
        '
        'picIcoDCConv2
        '
        Me.picIcoDCConv2.Image = CType(resources.GetObject("picIcoDCConv2.Image"),System.Drawing.Image)
        Me.picIcoDCConv2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCConv2.Name = "picIcoDCConv2"
        Me.picIcoDCConv2.Size = New System.Drawing.Size(25, 25)
        Me.picIcoDCConv2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCConv2.TabIndex = 4
        Me.picIcoDCConv2.TabStop = false
        '
        'picIcoDCConv1
        '
        Me.picIcoDCConv1.Image = CType(resources.GetObject("picIcoDCConv1.Image"),System.Drawing.Image)
        Me.picIcoDCConv1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCConv1.Name = "picIcoDCConv1"
        Me.picIcoDCConv1.Size = New System.Drawing.Size(25, 26)
        Me.picIcoDCConv1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCConv1.TabIndex = 3
        Me.picIcoDCConv1.TabStop = false
        '
        'lblPrgDCConv
        '
        Me.lblPrgDCConv.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgDCConv.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrgDCConv.Location = New System.Drawing.Point(33, 5)
        Me.lblPrgDCConv.Name = "lblPrgDCConv"
        Me.lblPrgDCConv.Size = New System.Drawing.Size(149, 20)
        Me.lblPrgDCConv.TabIndex = 0
        Me.lblPrgDCConv.Text = " データコンバート"
        '
        'pnlPrgDCRelation
        '
        Me.pnlPrgDCRelation.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgDCRelation.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgDCRelation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgDCRelation.CausesValidation = false
        Me.pnlPrgDCRelation.Controls.Add(Me.picIcoDCRelation2)
        Me.pnlPrgDCRelation.Controls.Add(Me.picIcoDCRelation1)
        Me.pnlPrgDCRelation.Controls.Add(Me.lblPrgDCRelation)
        Me.pnlPrgDCRelation.Location = New System.Drawing.Point(31, 297)
        Me.pnlPrgDCRelation.Name = "pnlPrgDCRelation"
        Me.pnlPrgDCRelation.Size = New System.Drawing.Size(184, 33)
        Me.pnlPrgDCRelation.TabIndex = 8
        '
        'picIcoDCRelation2
        '
        Me.picIcoDCRelation2.Image = CType(resources.GetObject("picIcoDCRelation2.Image"),System.Drawing.Image)
        Me.picIcoDCRelation2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCRelation2.Name = "picIcoDCRelation2"
        Me.picIcoDCRelation2.Size = New System.Drawing.Size(25, 25)
        Me.picIcoDCRelation2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCRelation2.TabIndex = 4
        Me.picIcoDCRelation2.TabStop = false
        '
        'picIcoDCRelation1
        '
        Me.picIcoDCRelation1.Image = CType(resources.GetObject("picIcoDCRelation1.Image"),System.Drawing.Image)
        Me.picIcoDCRelation1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCRelation1.Name = "picIcoDCRelation1"
        Me.picIcoDCRelation1.Size = New System.Drawing.Size(25, 26)
        Me.picIcoDCRelation1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCRelation1.TabIndex = 3
        Me.picIcoDCRelation1.TabStop = false
        '
        'lblPrgDCRelation
        '
        Me.lblPrgDCRelation.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgDCRelation.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrgDCRelation.Location = New System.Drawing.Point(33, 5)
        Me.lblPrgDCRelation.Name = "lblPrgDCRelation"
        Me.lblPrgDCRelation.Size = New System.Drawing.Size(149, 20)
        Me.lblPrgDCRelation.TabIndex = 0
        Me.lblPrgDCRelation.Text = " 紐 付 設 定 作 業"
        '
        'pnlPrgDCFileWrite
        '
        Me.pnlPrgDCFileWrite.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgDCFileWrite.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgDCFileWrite.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgDCFileWrite.CausesValidation = false
        Me.pnlPrgDCFileWrite.Controls.Add(Me.picIcoDCFileWrite2)
        Me.pnlPrgDCFileWrite.Controls.Add(Me.picIcoDCFileWrite1)
        Me.pnlPrgDCFileWrite.Controls.Add(Me.lblPrgDCFileWrite)
        Me.pnlPrgDCFileWrite.Location = New System.Drawing.Point(31, 258)
        Me.pnlPrgDCFileWrite.Name = "pnlPrgDCFileWrite"
        Me.pnlPrgDCFileWrite.Size = New System.Drawing.Size(184, 33)
        Me.pnlPrgDCFileWrite.TabIndex = 7
        '
        'picIcoDCFileWrite2
        '
        Me.picIcoDCFileWrite2.Image = CType(resources.GetObject("picIcoDCFileWrite2.Image"),System.Drawing.Image)
        Me.picIcoDCFileWrite2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCFileWrite2.Name = "picIcoDCFileWrite2"
        Me.picIcoDCFileWrite2.Size = New System.Drawing.Size(25, 25)
        Me.picIcoDCFileWrite2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCFileWrite2.TabIndex = 4
        Me.picIcoDCFileWrite2.TabStop = false
        '
        'picIcoDCFileWrite1
        '
        Me.picIcoDCFileWrite1.Image = CType(resources.GetObject("picIcoDCFileWrite1.Image"),System.Drawing.Image)
        Me.picIcoDCFileWrite1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCFileWrite1.Name = "picIcoDCFileWrite1"
        Me.picIcoDCFileWrite1.Size = New System.Drawing.Size(25, 26)
        Me.picIcoDCFileWrite1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCFileWrite1.TabIndex = 3
        Me.picIcoDCFileWrite1.TabStop = false
        '
        'lblPrgDCFileWrite
        '
        Me.lblPrgDCFileWrite.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgDCFileWrite.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrgDCFileWrite.Location = New System.Drawing.Point(33, 5)
        Me.lblPrgDCFileWrite.Name = "lblPrgDCFileWrite"
        Me.lblPrgDCFileWrite.Size = New System.Drawing.Size(149, 20)
        Me.lblPrgDCFileWrite.TabIndex = 0
        Me.lblPrgDCFileWrite.Text = " 中間ファイル書込"
        '
        'picArwDCEnd2
        '
        Me.picArwDCEnd2.Image = CType(resources.GetObject("picArwDCEnd2.Image"),System.Drawing.Image)
        Me.picArwDCEnd2.Location = New System.Drawing.Point(103, 375)
        Me.picArwDCEnd2.Name = "picArwDCEnd2"
        Me.picArwDCEnd2.Size = New System.Drawing.Size(40, 35)
        Me.picArwDCEnd2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwDCEnd2.TabIndex = 12
        Me.picArwDCEnd2.TabStop = false
        '
        'picArwDCJikko2
        '
        Me.picArwDCJikko2.Image = CType(resources.GetObject("picArwDCJikko2.Image"),System.Drawing.Image)
        Me.picArwDCJikko2.Location = New System.Drawing.Point(103, 157)
        Me.picArwDCJikko2.Name = "picArwDCJikko2"
        Me.picArwDCJikko2.Size = New System.Drawing.Size(40, 35)
        Me.picArwDCJikko2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwDCJikko2.TabIndex = 11
        Me.picArwDCJikko2.TabStop = false
        '
        'picArwDCSelect2
        '
        Me.picArwDCSelect2.Image = CType(resources.GetObject("picArwDCSelect2.Image"),System.Drawing.Image)
        Me.picArwDCSelect2.Location = New System.Drawing.Point(103, 62)
        Me.picArwDCSelect2.Name = "picArwDCSelect2"
        Me.picArwDCSelect2.Size = New System.Drawing.Size(40, 35)
        Me.picArwDCSelect2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwDCSelect2.TabIndex = 10
        Me.picArwDCSelect2.TabStop = false
        '
        'picArwDCEnd1
        '
        Me.picArwDCEnd1.Image = CType(resources.GetObject("picArwDCEnd1.Image"),System.Drawing.Image)
        Me.picArwDCEnd1.Location = New System.Drawing.Point(103, 375)
        Me.picArwDCEnd1.Name = "picArwDCEnd1"
        Me.picArwDCEnd1.Size = New System.Drawing.Size(40, 35)
        Me.picArwDCEnd1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwDCEnd1.TabIndex = 9
        Me.picArwDCEnd1.TabStop = false
        '
        'picArwDCJikko1
        '
        Me.picArwDCJikko1.Image = CType(resources.GetObject("picArwDCJikko1.Image"),System.Drawing.Image)
        Me.picArwDCJikko1.Location = New System.Drawing.Point(103, 157)
        Me.picArwDCJikko1.Name = "picArwDCJikko1"
        Me.picArwDCJikko1.Size = New System.Drawing.Size(40, 35)
        Me.picArwDCJikko1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwDCJikko1.TabIndex = 8
        Me.picArwDCJikko1.TabStop = false
        '
        'picArwDCSelect1
        '
        Me.picArwDCSelect1.Image = CType(resources.GetObject("picArwDCSelect1.Image"),System.Drawing.Image)
        Me.picArwDCSelect1.Location = New System.Drawing.Point(103, 62)
        Me.picArwDCSelect1.Name = "picArwDCSelect1"
        Me.picArwDCSelect1.Size = New System.Drawing.Size(40, 35)
        Me.picArwDCSelect1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwDCSelect1.TabIndex = 7
        Me.picArwDCSelect1.TabStop = false
        '
        'pnlPrgDCEnd
        '
        Me.pnlPrgDCEnd.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgDCEnd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgDCEnd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgDCEnd.CausesValidation = false
        Me.pnlPrgDCEnd.Controls.Add(Me.picIcoDCEnd2)
        Me.pnlPrgDCEnd.Controls.Add(Me.picIcoDCEnd1)
        Me.pnlPrgDCEnd.Controls.Add(Me.lblPrgDCEnd)
        Me.pnlPrgDCEnd.Location = New System.Drawing.Point(12, 416)
        Me.pnlPrgDCEnd.Name = "pnlPrgDCEnd"
        Me.pnlPrgDCEnd.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgDCEnd.TabIndex = 6
        '
        'picIcoDCEnd2
        '
        Me.picIcoDCEnd2.Image = CType(resources.GetObject("picIcoDCEnd2.Image"),System.Drawing.Image)
        Me.picIcoDCEnd2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCEnd2.Name = "picIcoDCEnd2"
        Me.picIcoDCEnd2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCEnd2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCEnd2.TabIndex = 4
        Me.picIcoDCEnd2.TabStop = false
        '
        'picIcoDCEnd1
        '
        Me.picIcoDCEnd1.Image = CType(resources.GetObject("picIcoDCEnd1.Image"),System.Drawing.Image)
        Me.picIcoDCEnd1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCEnd1.Name = "picIcoDCEnd1"
        Me.picIcoDCEnd1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCEnd1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcoDCEnd1.TabIndex = 3
        Me.picIcoDCEnd1.TabStop = false
        '
        'lblPrgDCEnd
        '
        Me.lblPrgDCEnd.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgDCEnd.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrgDCEnd.Location = New System.Drawing.Point(60, 8)
        Me.lblPrgDCEnd.Name = "lblPrgDCEnd"
        Me.lblPrgDCEnd.Size = New System.Drawing.Size(140, 30)
        Me.lblPrgDCEnd.TabIndex = 0
        Me.lblPrgDCEnd.Text = "  終      了  "
        Me.lblPrgDCEnd.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlPrgDCJikko
        '
        Me.pnlPrgDCJikko.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgDCJikko.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgDCJikko.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgDCJikko.CausesValidation = false
        Me.pnlPrgDCJikko.Controls.Add(Me.picIcoDCJikko2)
        Me.pnlPrgDCJikko.Controls.Add(Me.picIcoDCJikko1)
        Me.pnlPrgDCJikko.Controls.Add(Me.lblPrgDCJikko)
        Me.pnlPrgDCJikko.Location = New System.Drawing.Point(12, 200)
        Me.pnlPrgDCJikko.Name = "pnlPrgDCJikko"
        Me.pnlPrgDCJikko.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgDCJikko.TabIndex = 5
        '
        'picIcoDCJikko2
        '
        Me.picIcoDCJikko2.Image = CType(resources.GetObject("picIcoDCJikko2.Image"),System.Drawing.Image)
        Me.picIcoDCJikko2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCJikko2.Name = "picIcoDCJikko2"
        Me.picIcoDCJikko2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCJikko2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoDCJikko2.TabIndex = 4
        Me.picIcoDCJikko2.TabStop = false
        '
        'picIcoDCJikko1
        '
        Me.picIcoDCJikko1.Image = CType(resources.GetObject("picIcoDCJikko1.Image"),System.Drawing.Image)
        Me.picIcoDCJikko1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCJikko1.Name = "picIcoDCJikko1"
        Me.picIcoDCJikko1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCJikko1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoDCJikko1.TabIndex = 3
        Me.picIcoDCJikko1.TabStop = false
        '
        'lblPrgDCJikko
        '
        Me.lblPrgDCJikko.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgDCJikko.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrgDCJikko.Location = New System.Drawing.Point(60, 8)
        Me.lblPrgDCJikko.Name = "lblPrgDCJikko"
        Me.lblPrgDCJikko.Size = New System.Drawing.Size(140, 30)
        Me.lblPrgDCJikko.TabIndex = 0
        Me.lblPrgDCJikko.Text = " 移 行 処 理"
        Me.lblPrgDCJikko.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlPrgDCSelect
        '
        Me.pnlPrgDCSelect.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgDCSelect.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgDCSelect.CausesValidation = false
        Me.pnlPrgDCSelect.Controls.Add(Me.picIcoDCSelect2)
        Me.pnlPrgDCSelect.Controls.Add(Me.picIcoDCSelect1)
        Me.pnlPrgDCSelect.Controls.Add(Me.lblPrgDCSelect)
        Me.pnlPrgDCSelect.Location = New System.Drawing.Point(12, 103)
        Me.pnlPrgDCSelect.Name = "pnlPrgDCSelect"
        Me.pnlPrgDCSelect.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgDCSelect.TabIndex = 4
        '
        'picIcoDCSelect2
        '
        Me.picIcoDCSelect2.Image = CType(resources.GetObject("picIcoDCSelect2.Image"),System.Drawing.Image)
        Me.picIcoDCSelect2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCSelect2.Name = "picIcoDCSelect2"
        Me.picIcoDCSelect2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCSelect2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoDCSelect2.TabIndex = 4
        Me.picIcoDCSelect2.TabStop = false
        '
        'picIcoDCSelect1
        '
        Me.picIcoDCSelect1.Image = CType(resources.GetObject("picIcoDCSelect1.Image"),System.Drawing.Image)
        Me.picIcoDCSelect1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCSelect1.Name = "picIcoDCSelect1"
        Me.picIcoDCSelect1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCSelect1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoDCSelect1.TabIndex = 3
        Me.picIcoDCSelect1.TabStop = false
        '
        'lblPrgDCSelect
        '
        Me.lblPrgDCSelect.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgDCSelect.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrgDCSelect.Location = New System.Drawing.Point(60, 8)
        Me.lblPrgDCSelect.Name = "lblPrgDCSelect"
        Me.lblPrgDCSelect.Size = New System.Drawing.Size(140, 30)
        Me.lblPrgDCSelect.TabIndex = 0
        Me.lblPrgDCSelect.Text = "対象項目選択"
        Me.lblPrgDCSelect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlPrgDCHajimeni
        '
        Me.pnlPrgDCHajimeni.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgDCHajimeni.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgDCHajimeni.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgDCHajimeni.CausesValidation = false
        Me.pnlPrgDCHajimeni.Controls.Add(Me.picIcoDCHajimeni2)
        Me.pnlPrgDCHajimeni.Controls.Add(Me.picIcoDCHajimeni1)
        Me.pnlPrgDCHajimeni.Controls.Add(Me.lblPrgDCHajimeni)
        Me.pnlPrgDCHajimeni.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.pnlPrgDCHajimeni.Location = New System.Drawing.Point(12, 8)
        Me.pnlPrgDCHajimeni.Name = "pnlPrgDCHajimeni"
        Me.pnlPrgDCHajimeni.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgDCHajimeni.TabIndex = 0
        '
        'picIcoDCHajimeni2
        '
        Me.picIcoDCHajimeni2.Image = CType(resources.GetObject("picIcoDCHajimeni2.Image"),System.Drawing.Image)
        Me.picIcoDCHajimeni2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCHajimeni2.Name = "picIcoDCHajimeni2"
        Me.picIcoDCHajimeni2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCHajimeni2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoDCHajimeni2.TabIndex = 3
        Me.picIcoDCHajimeni2.TabStop = false
        '
        'picIcoDCHajimeni1
        '
        Me.picIcoDCHajimeni1.Image = CType(resources.GetObject("picIcoDCHajimeni1.Image"),System.Drawing.Image)
        Me.picIcoDCHajimeni1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoDCHajimeni1.Name = "picIcoDCHajimeni1"
        Me.picIcoDCHajimeni1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoDCHajimeni1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoDCHajimeni1.TabIndex = 2
        Me.picIcoDCHajimeni1.TabStop = false
        '
        'lblPrgDCHajimeni
        '
        Me.lblPrgDCHajimeni.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgDCHajimeni.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrgDCHajimeni.Location = New System.Drawing.Point(60, 8)
        Me.lblPrgDCHajimeni.Name = "lblPrgDCHajimeni"
        Me.lblPrgDCHajimeni.Size = New System.Drawing.Size(140, 30)
        Me.lblPrgDCHajimeni.TabIndex = 0
        Me.lblPrgDCHajimeni.Text = " は じ め に"
        Me.lblPrgDCHajimeni.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblHidden1
        '
        Me.lblHidden1.Location = New System.Drawing.Point(289, 11)
        Me.lblHidden1.Name = "lblHidden1"
        Me.lblHidden1.Size = New System.Drawing.Size(21, 55)
        Me.lblHidden1.TabIndex = 2
        Me.lblHidden1.Text = "　"
        '
        'tabCtrlMain
        '
        Me.tabCtrlMain.Controls.Add(Me.tabPageDev)
        Me.tabCtrlMain.Controls.Add(Me.tabPageStart)
        Me.tabCtrlMain.Controls.Add(Me.tabPageSession)
        Me.tabCtrlMain.Controls.Add(Me.tabPageSyoki)
        Me.tabCtrlMain.Controls.Add(Me.tabPageMenu)
        Me.tabCtrlMain.Controls.Add(Me.tabPageJizen)
        Me.tabCtrlMain.Controls.Add(Me.tabPageJigo)
        Me.tabCtrlMain.Controls.Add(Me.tabPageHojyo)
        Me.tabCtrlMain.Controls.Add(Me.tabPageHajimeni)
        Me.tabCtrlMain.Controls.Add(Me.tabPageSelect)
        Me.tabCtrlMain.Controls.Add(Me.tabPageJikko)
        Me.tabCtrlMain.Controls.Add(Me.tabPageEndOK)
        Me.tabCtrlMain.Controls.Add(Me.tabPageEndError)
        Me.tabCtrlMain.Controls.Add(Me.tabPageEndCancel)
        Me.tabCtrlMain.Controls.Add(Me.tabPageIkkatu)
        Me.tabCtrlMain.Controls.Add(Me.tabPageHanyoJizen)
        Me.tabCtrlMain.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
        Me.tabCtrlMain.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.tabCtrlMain.Location = New System.Drawing.Point(316, 40)
        Me.tabCtrlMain.Name = "tabCtrlMain"
        Me.tabCtrlMain.SelectedIndex = 0
        Me.tabCtrlMain.Size = New System.Drawing.Size(920, 580)
        Me.tabCtrlMain.TabIndex = 3
        '
        'tabPageDev
        '
        Me.tabPageDev.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageDev.Controls.Add(Me.GroupBox6)
        Me.tabPageDev.Controls.Add(Me.grpDevSettingX)
        Me.tabPageDev.Controls.Add(Me.PictureBox1)
        Me.tabPageDev.Controls.Add(Me.grpDevSetting1)
        Me.tabPageDev.Controls.Add(Me.grpDevSetting2)
        Me.tabPageDev.Controls.Add(Me.grpDevSetting3)
        Me.tabPageDev.Controls.Add(Me.grpRelation)
        Me.tabPageDev.Controls.Add(Me.GroupBox3)
        Me.tabPageDev.Controls.Add(Me.GroupBox2)
        Me.tabPageDev.Controls.Add(Me.GroupBox1)
        Me.tabPageDev.Controls.Add(Me.grptaihi)
        Me.tabPageDev.Location = New System.Drawing.Point(4, 27)
        Me.tabPageDev.Name = "tabPageDev"
        Me.tabPageDev.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageDev.Size = New System.Drawing.Size(912, 549)
        Me.tabPageDev.TabIndex = 5
        Me.tabPageDev.Text = " 開発用"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Label361)
        Me.GroupBox6.Controls.Add(Me.txtExistMidToBaseMid)
        Me.GroupBox6.Controls.Add(Me.btnExistMidToBaseMid)
        Me.GroupBox6.Controls.Add(Me.btnExistMidToBaseMidDirSerach)
        Me.GroupBox6.Location = New System.Drawing.Point(670, 226)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(228, 108)
        Me.GroupBox6.TabIndex = 168
        Me.GroupBox6.TabStop = false
        Me.GroupBox6.Text = "【テスト用汎用中間ファイル作成】"
        '
        'Label361
        '
        Me.Label361.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label361.Location = New System.Drawing.Point(18, 21)
        Me.Label361.Name = "Label361"
        Me.Label361.Size = New System.Drawing.Size(186, 20)
        Me.Label361.TabIndex = 145
        Me.Label361.Text = "中間ファイル作成元格納先"
        Me.Label361.UseCompatibleTextRendering = true
        '
        'txtExistMidToBaseMid
        '
        Me.txtExistMidToBaseMid.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtExistMidToBaseMid.Location = New System.Drawing.Point(18, 42)
        Me.txtExistMidToBaseMid.Name = "txtExistMidToBaseMid"
        Me.txtExistMidToBaseMid.Size = New System.Drawing.Size(164, 24)
        Me.txtExistMidToBaseMid.TabIndex = 143
        '
        'btnExistMidToBaseMid
        '
        Me.btnExistMidToBaseMid.BackColor = System.Drawing.SystemColors.Menu
        Me.btnExistMidToBaseMid.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnExistMidToBaseMid.Image = CType(resources.GetObject("btnExistMidToBaseMid.Image"),System.Drawing.Image)
        Me.btnExistMidToBaseMid.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExistMidToBaseMid.Location = New System.Drawing.Point(59, 72)
        Me.btnExistMidToBaseMid.Name = "btnExistMidToBaseMid"
        Me.btnExistMidToBaseMid.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnExistMidToBaseMid.Size = New System.Drawing.Size(100, 25)
        Me.btnExistMidToBaseMid.TabIndex = 46
        Me.btnExistMidToBaseMid.Text = "既存→汎用"
        Me.btnExistMidToBaseMid.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnExistMidToBaseMid.UseVisualStyleBackColor = true
        '
        'btnExistMidToBaseMidDirSerach
        '
        Me.btnExistMidToBaseMidDirSerach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnExistMidToBaseMidDirSerach.Location = New System.Drawing.Point(192, 42)
        Me.btnExistMidToBaseMidDirSerach.Name = "btnExistMidToBaseMidDirSerach"
        Me.btnExistMidToBaseMidDirSerach.Size = New System.Drawing.Size(30, 25)
        Me.btnExistMidToBaseMidDirSerach.TabIndex = 144
        Me.btnExistMidToBaseMidDirSerach.Text = "..."
        Me.btnExistMidToBaseMidDirSerach.UseVisualStyleBackColor = true
        '
        'grpDevSettingX
        '
        Me.grpDevSettingX.Controls.Add(Me.chkRelationFile)
        Me.grpDevSettingX.Controls.Add(Me.Label9)
        Me.grpDevSettingX.Controls.Add(Me.chkChildItemControl)
        Me.grpDevSettingX.Controls.Add(Me.Label48)
        Me.grpDevSettingX.Controls.Add(Me.Label41)
        Me.grpDevSettingX.Controls.Add(Me.Label50)
        Me.grpDevSettingX.Controls.Add(Me.Label49)
        Me.grpDevSettingX.Controls.Add(Me.txtLogOutputCnt)
        Me.grpDevSettingX.Location = New System.Drawing.Point(670, 354)
        Me.grpDevSettingX.Name = "grpDevSettingX"
        Me.grpDevSettingX.Size = New System.Drawing.Size(228, 188)
        Me.grpDevSettingX.TabIndex = 167
        Me.grpDevSettingX.TabStop = false
        Me.grpDevSettingX.Text = "【没】"
        '
        'chkChildItemControl
        '
        Me.chkChildItemControl.AutoSize = true
        Me.chkChildItemControl.Font = New System.Drawing.Font("メイリオ", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkChildItemControl.Location = New System.Drawing.Point(7, 19)
        Me.chkChildItemControl.Name = "chkChildItemControl"
        Me.chkChildItemControl.Size = New System.Drawing.Size(152, 18)
        Me.chkChildItemControl.TabIndex = 161
        Me.chkChildItemControl.Text = "子項目のチェックボックス制御"
        Me.chkChildItemControl.UseVisualStyleBackColor = true
        '
        'Label48
        '
        Me.Label48.Font = New System.Drawing.Font("メイリオ", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label48.Location = New System.Drawing.Point(8, 39)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(211, 30)
        Me.Label48.TabIndex = 162
        Me.Label48.Text = "対象項目選択画面内のチェックボックスの親子関係によるチェックを制御します。"
        Me.Label48.UseCompatibleTextRendering = true
        '
        'Label41
        '
        Me.Label41.Font = New System.Drawing.Font("メイリオ", 7!)
        Me.Label41.Location = New System.Drawing.Point(7, 71)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(73, 16)
        Me.Label41.TabIndex = 147
        Me.Label41.Text = "● ログ出力件数"
        Me.Label41.UseCompatibleTextRendering = true
        '
        'Label50
        '
        Me.Label50.Font = New System.Drawing.Font("メイリオ", 7!)
        Me.Label50.Location = New System.Drawing.Point(188, 86)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(32, 16)
        Me.Label50.TabIndex = 149
        Me.Label50.Text = "件毎"
        Me.Label50.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label50.UseCompatibleTextRendering = true
        '
        'Label49
        '
        Me.Label49.Font = New System.Drawing.Font("メイリオ", 7!)
        Me.Label49.Location = New System.Drawing.Point(8, 87)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(115, 29)
        Me.Label49.TabIndex = 148
        Me.Label49.Text = "右記に設定された件数毎にログを出力します。"
        Me.Label49.UseCompatibleTextRendering = true
        '
        'txtLogOutputCnt
        '
        Me.txtLogOutputCnt.Font = New System.Drawing.Font("メイリオ", 7!)
        Me.txtLogOutputCnt.Location = New System.Drawing.Point(129, 84)
        Me.txtLogOutputCnt.Name = "txtLogOutputCnt"
        Me.txtLogOutputCnt.Size = New System.Drawing.Size(53, 21)
        Me.txtLogOutputCnt.TabIndex = 146
        Me.txtLogOutputCnt.Text = "1"
        Me.txtLogOutputCnt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"),System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(770, 27)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(128, 100)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 165
        Me.PictureBox1.TabStop = false
        '
        'grpDevSetting1
        '
        Me.grpDevSetting1.Controls.Add(Me.Label51)
        Me.grpDevSetting1.Controls.Add(Me.Label360)
        Me.grpDevSetting1.Controls.Add(Me.chkKyNewest)
        Me.grpDevSetting1.Controls.Add(Me.chkSyskanriinit)
        Me.grpDevSetting1.Controls.Add(Me.Label46)
        Me.grpDevSetting1.Controls.Add(Me.Label45)
        Me.grpDevSetting1.Controls.Add(Me.Label44)
        Me.grpDevSetting1.Controls.Add(Me.chkRelTblDrop)
        Me.grpDevSetting1.Controls.Add(Me.Label43)
        Me.grpDevSetting1.Controls.Add(Me.Label42)
        Me.grpDevSetting1.Controls.Add(Me.chkV7ViewDrop)
        Me.grpDevSetting1.Controls.Add(Me.chkLogTblDrop)
        Me.grpDevSetting1.Controls.Add(Me.chkOverWrite)
        Me.grpDevSetting1.Controls.Add(Me.Label52)
        Me.grpDevSetting1.Location = New System.Drawing.Point(23, 27)
        Me.grpDevSetting1.Name = "grpDevSetting1"
        Me.grpDevSetting1.Size = New System.Drawing.Size(726, 191)
        Me.grpDevSetting1.TabIndex = 164
        Me.grpDevSetting1.TabStop = false
        Me.grpDevSetting1.Text = "【開発用設定1】"
        '
        'Label51
        '
        Me.Label51.Location = New System.Drawing.Point(240, 156)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(465, 22)
        Me.Label51.TabIndex = 164
        Me.Label51.Text = "解約されていない最新の契約情報のみをコンバートします。(V7受託CVと同等)"
        Me.Label51.UseCompatibleTextRendering = true
        '
        'Label360
        '
        Me.Label360.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label360.Location = New System.Drawing.Point(240, 135)
        Me.Label360.Name = "Label360"
        Me.Label360.Size = New System.Drawing.Size(444, 22)
        Me.Label360.TabIndex = 166
        Me.Label360.Text = "賃貸革命10の初期設定(XML)値を初期化します。"
        Me.Label360.UseCompatibleTextRendering = true
        '
        'chkKyNewest
        '
        Me.chkKyNewest.Location = New System.Drawing.Point(25, 155)
        Me.chkKyNewest.Name = "chkKyNewest"
        Me.chkKyNewest.Size = New System.Drawing.Size(195, 22)
        Me.chkKyNewest.TabIndex = 163
        Me.chkKyNewest.Text = "最新契約情報のみ対象"
        Me.chkKyNewest.UseVisualStyleBackColor = true
        '
        'chkSyskanriinit
        '
        Me.chkSyskanriinit.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkSyskanriinit.ForeColor = System.Drawing.Color.Black
        Me.chkSyskanriinit.Location = New System.Drawing.Point(25, 133)
        Me.chkSyskanriinit.Name = "chkSyskanriinit"
        Me.chkSyskanriinit.Size = New System.Drawing.Size(195, 22)
        Me.chkSyskanriinit.TabIndex = 161
        Me.chkSyskanriinit.Text = "初期設定初期化"
        Me.chkSyskanriinit.UseVisualStyleBackColor = true
        '
        'Label46
        '
        Me.Label46.Location = New System.Drawing.Point(25, 69)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(195, 20)
        Me.Label46.TabIndex = 160
        Me.Label46.Text = "● 仮TBL削除 (10DB作成)"
        Me.Label46.UseCompatibleTextRendering = true
        '
        'Label45
        '
        Me.Label45.Location = New System.Drawing.Point(240, 69)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(444, 22)
        Me.Label45.TabIndex = 159
        Me.Label45.Text = "賃貸革命10へコンバートする為に作成した以下の作業用TBLを削除します。"
        Me.Label45.UseCompatibleTextRendering = true
        '
        'Label44
        '
        Me.Label44.Location = New System.Drawing.Point(240, 92)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(444, 24)
        Me.Label44.TabIndex = 158
        Me.Label44.Text = "作業用として作成した紐付設定TBLを削除(ドロップ)します。"
        Me.Label44.UseCompatibleTextRendering = true
        '
        'chkRelTblDrop
        '
        Me.chkRelTblDrop.AutoSize = true
        Me.chkRelTblDrop.Location = New System.Drawing.Point(45, 91)
        Me.chkRelTblDrop.Name = "chkRelTblDrop"
        Me.chkRelTblDrop.Size = New System.Drawing.Size(111, 22)
        Me.chkRelTblDrop.TabIndex = 157
        Me.chkRelTblDrop.Text = "紐付設定データ"
        Me.chkRelTblDrop.UseVisualStyleBackColor = true
        '
        'Label43
        '
        Me.Label43.Location = New System.Drawing.Point(240, 46)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(444, 24)
        Me.Label43.TabIndex = 156
        Me.Label43.Text = "賃貸革命V7から必要項目を抽出する為に作成したVIEWを削除します。"
        Me.Label43.UseCompatibleTextRendering = true
        '
        'Label42
        '
        Me.Label42.Location = New System.Drawing.Point(240, 113)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(444, 22)
        Me.Label42.TabIndex = 155
        Me.Label42.Text = "作業用として作成したコンバートログTBLを削除(ドロップ)します。"
        Me.Label42.UseCompatibleTextRendering = true
        '
        'chkV7ViewDrop
        '
        Me.chkV7ViewDrop.AutoSize = true
        Me.chkV7ViewDrop.Location = New System.Drawing.Point(25, 46)
        Me.chkV7ViewDrop.Name = "chkV7ViewDrop"
        Me.chkV7ViewDrop.Size = New System.Drawing.Size(196, 22)
        Me.chkV7ViewDrop.TabIndex = 154
        Me.chkV7ViewDrop.Text = "仮作成VIEWを削除 (V7側作成)"
        Me.chkV7ViewDrop.UseVisualStyleBackColor = true
        '
        'chkLogTblDrop
        '
        Me.chkLogTblDrop.AutoSize = true
        Me.chkLogTblDrop.Location = New System.Drawing.Point(45, 112)
        Me.chkLogTblDrop.Name = "chkLogTblDrop"
        Me.chkLogTblDrop.Size = New System.Drawing.Size(87, 22)
        Me.chkLogTblDrop.TabIndex = 153
        Me.chkLogTblDrop.Text = "ログデータ"
        Me.chkLogTblDrop.UseVisualStyleBackColor = true
        '
        'chkOverWrite
        '
        Me.chkOverWrite.AutoSize = true
        Me.chkOverWrite.Location = New System.Drawing.Point(25, 25)
        Me.chkOverWrite.Name = "chkOverWrite"
        Me.chkOverWrite.Size = New System.Drawing.Size(123, 22)
        Me.chkOverWrite.TabIndex = 152
        Me.chkOverWrite.Text = "上書きコンバート"
        Me.chkOverWrite.UseVisualStyleBackColor = true
        '
        'Label52
        '
        Me.Label52.Location = New System.Drawing.Point(240, 25)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(448, 22)
        Me.Label52.TabIndex = 151
        Me.Label52.Text = "該当するデータが既に登録されている場合、上書きでコンバートを行います。"
        Me.Label52.UseCompatibleTextRendering = true
        '
        'grpDevSetting2
        '
        Me.grpDevSetting2.Controls.Add(Me.btnRelDirSeach)
        Me.grpDevSetting2.Controls.Add(Me.txtRelationDirPath)
        Me.grpDevSetting2.Controls.Add(Me.lblRelationDir)
        Me.grpDevSetting2.Controls.Add(Me.Label11)
        Me.grpDevSetting2.Controls.Add(Me.chkRelation)
        Me.grpDevSetting2.Controls.Add(Me.Label47)
        Me.grpDevSetting2.Controls.Add(Me.chkMidNotStop)
        Me.grpDevSetting2.Controls.Add(Me.chkCVStart)
        Me.grpDevSetting2.Controls.Add(Me.Label8)
        Me.grpDevSetting2.Controls.Add(Me.Label10)
        Me.grpDevSetting2.Controls.Add(Me.chkMiddleFile)
        Me.grpDevSetting2.Location = New System.Drawing.Point(23, 224)
        Me.grpDevSetting2.Name = "grpDevSetting2"
        Me.grpDevSetting2.Size = New System.Drawing.Size(611, 165)
        Me.grpDevSetting2.TabIndex = 163
        Me.grpDevSetting2.TabStop = false
        Me.grpDevSetting2.Text = "【開発用設定2】"
        '
        'btnRelDirSeach
        '
        Me.btnRelDirSeach.Location = New System.Drawing.Point(526, 130)
        Me.btnRelDirSeach.Name = "btnRelDirSeach"
        Me.btnRelDirSeach.Size = New System.Drawing.Size(30, 25)
        Me.btnRelDirSeach.TabIndex = 168
        Me.btnRelDirSeach.Text = "..."
        Me.btnRelDirSeach.UseVisualStyleBackColor = true
        '
        'txtRelationDirPath
        '
        Me.txtRelationDirPath.Location = New System.Drawing.Point(175, 130)
        Me.txtRelationDirPath.Name = "txtRelationDirPath"
        Me.txtRelationDirPath.Size = New System.Drawing.Size(345, 25)
        Me.txtRelationDirPath.TabIndex = 167
        '
        'lblRelationDir
        '
        Me.lblRelationDir.Location = New System.Drawing.Point(25, 133)
        Me.lblRelationDir.Name = "lblRelationDir"
        Me.lblRelationDir.Size = New System.Drawing.Size(144, 21)
        Me.lblRelationDir.TabIndex = 166
        Me.lblRelationDir.Text = "● 紐付ファイル格納先"
        Me.lblRelationDir.UseCompatibleTextRendering = true
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(175, 89)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(386, 38)
        Me.Label11.TabIndex = 165
        Me.Label11.Text = "入金項目などの移行元と移行先のデータの紐付けを行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"既に紐付け処理を行っている場合はチェックをOFFにして下さい。"
        Me.Label11.UseCompatibleTextRendering = true
        '
        'chkRelation
        '
        Me.chkRelation.AutoSize = true
        Me.chkRelation.Location = New System.Drawing.Point(25, 88)
        Me.chkRelation.Name = "chkRelation"
        Me.chkRelation.Size = New System.Drawing.Size(111, 22)
        Me.chkRelation.TabIndex = 164
        Me.chkRelation.Text = "項目紐付け設定"
        Me.chkRelation.UseVisualStyleBackColor = true
        '
        'Label47
        '
        Me.Label47.Location = New System.Drawing.Point(175, 68)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(376, 21)
        Me.Label47.TabIndex = 144
        Me.Label47.Text = "移行元→中間→移行先を連続で行います。"
        Me.Label47.UseCompatibleTextRendering = true
        '
        'chkMidNotStop
        '
        Me.chkMidNotStop.AutoSize = true
        Me.chkMidNotStop.Location = New System.Drawing.Point(25, 67)
        Me.chkMidNotStop.Name = "chkMidNotStop"
        Me.chkMidNotStop.Size = New System.Drawing.Size(75, 22)
        Me.chkMidNotStop.TabIndex = 143
        Me.chkMidNotStop.Text = "連続実行"
        Me.chkMidNotStop.UseVisualStyleBackColor = true
        '
        'chkCVStart
        '
        Me.chkCVStart.AutoSize = true
        Me.chkCVStart.Location = New System.Drawing.Point(25, 46)
        Me.chkCVStart.Name = "chkCVStart"
        Me.chkCVStart.Size = New System.Drawing.Size(111, 22)
        Me.chkCVStart.TabIndex = 133
        Me.chkCVStart.Text = "コンバート実施"
        Me.chkCVStart.UseVisualStyleBackColor = true
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(175, 47)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(345, 21)
        Me.Label8.TabIndex = 132
        Me.Label8.Text = "コンバートを実行します。"
        Me.Label8.UseCompatibleTextRendering = true
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(175, 27)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(365, 20)
        Me.Label10.TabIndex = 131
        Me.Label10.Text = "移行元のデータを中間ファイルに出力します。"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label10.UseCompatibleTextRendering = true
        '
        'chkMiddleFile
        '
        Me.chkMiddleFile.AutoSize = true
        Me.chkMiddleFile.Location = New System.Drawing.Point(25, 25)
        Me.chkMiddleFile.Name = "chkMiddleFile"
        Me.chkMiddleFile.Size = New System.Drawing.Size(123, 22)
        Me.chkMiddleFile.TabIndex = 130
        Me.chkMiddleFile.Text = "中間ファイル出力"
        Me.chkMiddleFile.UseVisualStyleBackColor = true
        '
        'grpDevSetting3
        '
        Me.grpDevSetting3.Controls.Add(Me.Label55)
        Me.grpDevSetting3.Controls.Add(Me.Label54)
        Me.grpDevSetting3.Controls.Add(Me.chkRommDuplicate)
        Me.grpDevSetting3.Controls.Add(Me.chkEmptyRoomNo)
        Me.grpDevSetting3.Controls.Add(Me.Label53)
        Me.grpDevSetting3.Location = New System.Drawing.Point(23, 400)
        Me.grpDevSetting3.Name = "grpDevSetting3"
        Me.grpDevSetting3.Size = New System.Drawing.Size(632, 144)
        Me.grpDevSetting3.TabIndex = 145
        Me.grpDevSetting3.TabStop = false
        Me.grpDevSetting3.Text = "【開発用設定3】"
        '
        'Label55
        '
        Me.Label55.Location = New System.Drawing.Point(240, 71)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(386, 22)
        Me.Label55.TabIndex = 160
        Me.Label55.Text = "重複部屋番号が存在する場合、任意の部屋番号に変更します。"
        Me.Label55.UseCompatibleTextRendering = true
        '
        'Label54
        '
        Me.Label54.Location = New System.Drawing.Point(240, 50)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(386, 22)
        Me.Label54.TabIndex = 159
        Me.Label54.Text = "移行元の部屋番号が未設定の場合、任意の部屋番号を設定します。"
        Me.Label54.UseCompatibleTextRendering = true
        '
        'chkRommDuplicate
        '
        Me.chkRommDuplicate.Location = New System.Drawing.Point(45, 71)
        Me.chkRommDuplicate.Name = "chkRommDuplicate"
        Me.chkRommDuplicate.Size = New System.Drawing.Size(171, 22)
        Me.chkRommDuplicate.TabIndex = 158
        Me.chkRommDuplicate.Text = "重複部屋番号変更設定"
        Me.chkRommDuplicate.UseVisualStyleBackColor = true
        '
        'chkEmptyRoomNo
        '
        Me.chkEmptyRoomNo.AutoSize = true
        Me.chkEmptyRoomNo.Location = New System.Drawing.Point(45, 49)
        Me.chkEmptyRoomNo.Name = "chkEmptyRoomNo"
        Me.chkEmptyRoomNo.Size = New System.Drawing.Size(171, 22)
        Me.chkEmptyRoomNo.TabIndex = 157
        Me.chkEmptyRoomNo.Text = "未設定部屋番号初期値設定"
        Me.chkEmptyRoomNo.UseVisualStyleBackColor = true
        '
        'Label53
        '
        Me.Label53.Location = New System.Drawing.Point(25, 26)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(131, 20)
        Me.Label53.TabIndex = 156
        Me.Label53.Text = "● 部屋番号設定"
        Me.Label53.UseCompatibleTextRendering = true
        '
        'grpRelation
        '
        Me.grpRelation.Location = New System.Drawing.Point(391, 423)
        Me.grpRelation.Name = "grpRelation"
        Me.grpRelation.Size = New System.Drawing.Size(64, 72)
        Me.grpRelation.TabIndex = 158
        Me.grpRelation.TabStop = false
        Me.grpRelation.Text = "紐付設定"
        Me.grpRelation.Visible = false
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label39)
        Me.GroupBox3.Controls.Add(Me.Label33)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label56)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Location = New System.Drawing.Point(301, 406)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(77, 68)
        Me.GroupBox3.TabIndex = 157
        Me.GroupBox3.TabStop = false
        Me.GroupBox3.Text = "退避4"
        Me.GroupBox3.Visible = false
        '
        'Label39
        '
        Me.Label39.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label39.Location = New System.Drawing.Point(40, -6)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(32, 96)
        Me.Label39.TabIndex = 136
        Me.Label39.Text = "├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"└"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label33
        '
        Me.Label33.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label33.Location = New System.Drawing.Point(44, -14)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(32, 64)
        Me.Label33.TabIndex = 135
        Me.Label33.Text = "├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"└"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label6.Location = New System.Drawing.Point(28, -14)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 40)
        Me.Label6.TabIndex = 134
        Me.Label6.Text = "├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"└"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label56
        '
        Me.Label56.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label56.Location = New System.Drawing.Point(4, -26)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(32, 16)
        Me.Label56.TabIndex = 133
        Me.Label56.Text = "└"
        Me.Label56.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label16.Location = New System.Drawing.Point(-180, -26)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(32, 16)
        Me.Label16.TabIndex = 132
        Me.Label16.Text = "└"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label15.Location = New System.Drawing.Point(112, 32)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(32, 40)
        Me.Label15.TabIndex = 131
        Me.Label15.Text = "├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"└"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label14.Location = New System.Drawing.Point(16, 32)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(32, 40)
        Me.Label14.TabIndex = 130
        Me.Label14.Text = "├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"└"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label12.Location = New System.Drawing.Point(64, 32)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(32, 40)
        Me.Label12.TabIndex = 129
        Me.Label12.Text = "├"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"|"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"└"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label18.Location = New System.Drawing.Point(16, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(32, 16)
        Me.Label18.TabIndex = 128
        Me.Label18.Text = "└"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkMstKasyorui)
        Me.GroupBox2.Controls.Add(Me.chkMstGenjotokuyaku)
        Me.GroupBox2.Controls.Add(Me.lblTokuyakuInfo)
        Me.GroupBox2.Location = New System.Drawing.Point(183, 401)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(96, 103)
        Me.GroupBox2.TabIndex = 156
        Me.GroupBox2.TabStop = false
        Me.GroupBox2.Text = "退避3"
        Me.GroupBox2.Visible = false
        '
        'chkMstKasyorui
        '
        Me.chkMstKasyorui.AutoSize = true
        Me.chkMstKasyorui.Location = New System.Drawing.Point(8, 112)
        Me.chkMstKasyorui.Name = "chkMstKasyorui"
        Me.chkMstKasyorui.Size = New System.Drawing.Size(111, 22)
        Me.chkMstKasyorui.TabIndex = 127
        Me.chkMstKasyorui.Text = "箇所分類マスタ"
        Me.chkMstKasyorui.UseVisualStyleBackColor = true
        '
        'chkMstGenjotokuyaku
        '
        Me.chkMstGenjotokuyaku.AutoSize = true
        Me.chkMstGenjotokuyaku.Location = New System.Drawing.Point(8, 88)
        Me.chkMstGenjotokuyaku.Name = "chkMstGenjotokuyaku"
        Me.chkMstGenjotokuyaku.Size = New System.Drawing.Size(135, 22)
        Me.chkMstGenjotokuyaku.TabIndex = 126
        Me.chkMstGenjotokuyaku.Text = "原状回復特約マスタ"
        Me.chkMstGenjotokuyaku.UseVisualStyleBackColor = true
        '
        'lblTokuyakuInfo
        '
        Me.lblTokuyakuInfo.Location = New System.Drawing.Point(-72, 22)
        Me.lblTokuyakuInfo.Name = "lblTokuyakuInfo"
        Me.lblTokuyakuInfo.Size = New System.Drawing.Size(344, 56)
        Me.lblTokuyakuInfo.TabIndex = 124
        Me.lblTokuyakuInfo.Text = "※1.特約事項マスタおよび原状回復特約マスタは特約事項内容設定へ"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"      統合されます。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※2.箇所分類マスタおよびクレーム分類マスタはクレーム管理設定へ"& _ 
    ""&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"      統合されます。"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label34)
        Me.GroupBox1.Controls.Add(Me.Label29)
        Me.GroupBox1.Controls.Add(Me.chkMstKozasyubetu)
        Me.GroupBox1.Controls.Add(Me.chkMstSetubi)
        Me.GroupBox1.Controls.Add(Me.chkMstKozo)
        Me.GroupBox1.Controls.Add(Me.chkMstTorihikitaiyo)
        Me.GroupBox1.Controls.Add(Me.chkMstNkinkbn)
        Me.GroupBox1.Controls.Add(Me.chkMstHyrui)
        Me.GroupBox1.Controls.Add(Me.chkMstNkinkomok)
        Me.GroupBox1.Controls.Add(Me.chkMstBkrui)
        Me.GroupBox1.Location = New System.Drawing.Point(71, 407)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(88, 40)
        Me.GroupBox1.TabIndex = 124
        Me.GroupBox1.TabStop = false
        Me.GroupBox1.Text = "退避2"
        Me.GroupBox1.Visible = false
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(48, 216)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(160, 16)
        Me.Label34.TabIndex = 128
        Me.Label34.Text = "項目別進捗"
        Me.Label34.UseCompatibleTextRendering = true
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(40, 176)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(160, 16)
        Me.Label29.TabIndex = 127
        Me.Label29.Text = "全体進捗"
        Me.Label29.UseCompatibleTextRendering = true
        '
        'chkMstKozasyubetu
        '
        Me.chkMstKozasyubetu.AutoSize = true
        Me.chkMstKozasyubetu.Location = New System.Drawing.Point(167, 45)
        Me.chkMstKozasyubetu.Name = "chkMstKozasyubetu"
        Me.chkMstKozasyubetu.Size = New System.Drawing.Size(111, 22)
        Me.chkMstKozasyubetu.TabIndex = 126
        Me.chkMstKozasyubetu.Text = "口座種別マスタ"
        Me.chkMstKozasyubetu.UseVisualStyleBackColor = true
        '
        'chkMstSetubi
        '
        Me.chkMstSetubi.AutoSize = true
        Me.chkMstSetubi.Location = New System.Drawing.Point(167, 132)
        Me.chkMstSetubi.Name = "chkMstSetubi"
        Me.chkMstSetubi.Size = New System.Drawing.Size(87, 22)
        Me.chkMstSetubi.TabIndex = 125
        Me.chkMstSetubi.Text = "設備マスタ"
        Me.chkMstSetubi.UseVisualStyleBackColor = true
        '
        'chkMstKozo
        '
        Me.chkMstKozo.AutoSize = true
        Me.chkMstKozo.Location = New System.Drawing.Point(39, 103)
        Me.chkMstKozo.Name = "chkMstKozo"
        Me.chkMstKozo.Size = New System.Drawing.Size(87, 22)
        Me.chkMstKozo.TabIndex = 124
        Me.chkMstKozo.Text = "構造マスタ"
        Me.chkMstKozo.UseVisualStyleBackColor = true
        '
        'chkMstTorihikitaiyo
        '
        Me.chkMstTorihikitaiyo.AutoSize = true
        Me.chkMstTorihikitaiyo.Location = New System.Drawing.Point(39, 132)
        Me.chkMstTorihikitaiyo.Name = "chkMstTorihikitaiyo"
        Me.chkMstTorihikitaiyo.Size = New System.Drawing.Size(111, 22)
        Me.chkMstTorihikitaiyo.TabIndex = 123
        Me.chkMstTorihikitaiyo.Text = "取引態様マスタ"
        Me.chkMstTorihikitaiyo.UseVisualStyleBackColor = true
        '
        'chkMstNkinkbn
        '
        Me.chkMstNkinkbn.AutoSize = true
        Me.chkMstNkinkbn.Location = New System.Drawing.Point(167, 74)
        Me.chkMstNkinkbn.Name = "chkMstNkinkbn"
        Me.chkMstNkinkbn.Size = New System.Drawing.Size(111, 22)
        Me.chkMstNkinkbn.TabIndex = 122
        Me.chkMstNkinkbn.Text = "入金区分マスタ"
        Me.chkMstNkinkbn.UseVisualStyleBackColor = true
        '
        'chkMstHyrui
        '
        Me.chkMstHyrui.AutoSize = true
        Me.chkMstHyrui.Location = New System.Drawing.Point(39, 74)
        Me.chkMstHyrui.Name = "chkMstHyrui"
        Me.chkMstHyrui.Size = New System.Drawing.Size(111, 22)
        Me.chkMstHyrui.TabIndex = 121
        Me.chkMstHyrui.Text = "部屋分類マスタ"
        Me.chkMstHyrui.UseVisualStyleBackColor = true
        '
        'chkMstNkinkomok
        '
        Me.chkMstNkinkomok.AutoSize = true
        Me.chkMstNkinkomok.Location = New System.Drawing.Point(167, 103)
        Me.chkMstNkinkomok.Name = "chkMstNkinkomok"
        Me.chkMstNkinkomok.Size = New System.Drawing.Size(111, 22)
        Me.chkMstNkinkomok.TabIndex = 120
        Me.chkMstNkinkomok.Text = "入金項目マスタ"
        Me.chkMstNkinkomok.UseVisualStyleBackColor = true
        '
        'chkMstBkrui
        '
        Me.chkMstBkrui.AutoSize = true
        Me.chkMstBkrui.Location = New System.Drawing.Point(39, 45)
        Me.chkMstBkrui.Name = "chkMstBkrui"
        Me.chkMstBkrui.Size = New System.Drawing.Size(111, 22)
        Me.chkMstBkrui.TabIndex = 119
        Me.chkMstBkrui.Text = "物件分類マスタ"
        Me.chkMstBkrui.UseVisualStyleBackColor = true
        '
        'grptaihi
        '
        Me.grptaihi.Controls.Add(Me.Label7)
        Me.grptaihi.Controls.Add(Me.btnFileSeach)
        Me.grptaihi.Controls.Add(Me.Label5)
        Me.grptaihi.Controls.Add(Me.chkGyYatinhosyoKoza)
        Me.grptaihi.Controls.Add(Me.CheckBox66)
        Me.grptaihi.Controls.Add(Me.lblGy)
        Me.grptaihi.Controls.Add(Me.lblMst)
        Me.grptaihi.Controls.Add(Me.lblKys)
        Me.grptaihi.Controls.Add(Me.lblOwner)
        Me.grptaihi.Controls.Add(Me.lblJisya)
        Me.grptaihi.Controls.Add(Me.Label17)
        Me.grptaihi.Controls.Add(Me.Label19)
        Me.grptaihi.Controls.Add(Me.Label20)
        Me.grptaihi.Controls.Add(Me.chkMs25)
        Me.grptaihi.Controls.Add(Me.chkMsKeiyakusyubetu)
        Me.grptaihi.Controls.Add(Me.chkMs23)
        Me.grptaihi.Controls.Add(Me.chkMs24)
        Me.grptaihi.Controls.Add(Me.chkMsOwevent)
        Me.grptaihi.Controls.Add(Me.chkMsYane)
        Me.grptaihi.Controls.Add(Me.chkMsKeikaikakunin)
        Me.grptaihi.Controls.Add(Me.chkMsKinyu)
        Me.grptaihi.Controls.Add(Me.chkMsKinyuten)
        Me.grptaihi.Location = New System.Drawing.Point(71, 452)
        Me.grptaihi.Name = "grptaihi"
        Me.grptaihi.Size = New System.Drawing.Size(64, 48)
        Me.grptaihi.TabIndex = 115
        Me.grptaihi.TabStop = false
        Me.grptaihi.Text = "退避"
        Me.grptaihi.Visible = false
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 136)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(150, 25)
        Me.Label7.TabIndex = 126
        Me.Label7.Text = "中間ファイル登録データチェック"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label7.UseCompatibleTextRendering = true
        '
        'btnFileSeach
        '
        Me.btnFileSeach.Location = New System.Drawing.Point(368, 128)
        Me.btnFileSeach.Name = "btnFileSeach"
        Me.btnFileSeach.Size = New System.Drawing.Size(27, 21)
        Me.btnFileSeach.TabIndex = 125
        Me.btnFileSeach.Text = "..."
        Me.btnFileSeach.UseVisualStyleBackColor = true
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(304, 96)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(374, 16)
        Me.Label5.TabIndex = 117
        Me.Label5.Text = "既存データを上書きして移行する場合はチェックをONにして下さい。"
        Me.Label5.UseCompatibleTextRendering = true
        '
        'chkGyYatinhosyoKoza
        '
        Me.chkGyYatinhosyoKoza.AutoSize = true
        Me.chkGyYatinhosyoKoza.Location = New System.Drawing.Point(129, 86)
        Me.chkGyYatinhosyoKoza.Name = "chkGyYatinhosyoKoza"
        Me.chkGyYatinhosyoKoza.Size = New System.Drawing.Size(147, 22)
        Me.chkGyYatinhosyoKoza.TabIndex = 116
        Me.chkGyYatinhosyoKoza.Text = "家賃保証業者口座情報"
        Me.chkGyYatinhosyoKoza.UseVisualStyleBackColor = true
        '
        'CheckBox66
        '
        Me.CheckBox66.AutoSize = true
        Me.CheckBox66.Location = New System.Drawing.Point(8, 16)
        Me.CheckBox66.Name = "CheckBox66"
        Me.CheckBox66.Size = New System.Drawing.Size(104, 22)
        Me.CheckBox66.TabIndex = 115
        Me.CheckBox66.Text = "src内部に記載"
        Me.CheckBox66.UseVisualStyleBackColor = true
        '
        'lblGy
        '
        Me.lblGy.AutoSize = true
        Me.lblGy.Location = New System.Drawing.Point(195, 171)
        Me.lblGy.Name = "lblGy"
        Me.lblGy.Size = New System.Drawing.Size(79, 18)
        Me.lblGy.TabIndex = 113
        Me.lblGy.Text = "2.業者マスタ"
        '
        'lblMst
        '
        Me.lblMst.AutoSize = true
        Me.lblMst.Location = New System.Drawing.Point(195, 178)
        Me.lblMst.Name = "lblMst"
        Me.lblMst.Size = New System.Drawing.Size(67, 18)
        Me.lblMst.TabIndex = 114
        Me.lblMst.Text = "1.マスタ系"
        '
        'lblKys
        '
        Me.lblKys.AutoSize = true
        Me.lblKys.Location = New System.Drawing.Point(185, 123)
        Me.lblKys.Name = "lblKys"
        Me.lblKys.Size = New System.Drawing.Size(79, 18)
        Me.lblKys.TabIndex = 107
        Me.lblKys.Text = "5.契約者情報"
        '
        'lblOwner
        '
        Me.lblOwner.AutoSize = true
        Me.lblOwner.Location = New System.Drawing.Point(185, 123)
        Me.lblOwner.Name = "lblOwner"
        Me.lblOwner.Size = New System.Drawing.Size(67, 18)
        Me.lblOwner.TabIndex = 108
        Me.lblOwner.Text = "4.家主情報"
        '
        'lblJisya
        '
        Me.lblJisya.AutoSize = true
        Me.lblJisya.Location = New System.Drawing.Point(195, 146)
        Me.lblJisya.Name = "lblJisya"
        Me.lblJisya.Size = New System.Drawing.Size(67, 18)
        Me.lblJisya.TabIndex = 109
        Me.lblJisya.Text = "3.自社情報"
        '
        'Label17
        '
        Me.Label17.AutoSize = true
        Me.Label17.Location = New System.Drawing.Point(185, 123)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 18)
        Me.Label17.TabIndex = 110
        Me.Label17.Text = "1.マスタ系"
        '
        'Label19
        '
        Me.Label19.AutoSize = true
        Me.Label19.Location = New System.Drawing.Point(185, 123)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(67, 18)
        Me.Label19.TabIndex = 111
        Me.Label19.Text = "1.マスタ系"
        '
        'Label20
        '
        Me.Label20.AutoSize = true
        Me.Label20.Location = New System.Drawing.Point(185, 123)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(67, 18)
        Me.Label20.TabIndex = 112
        Me.Label20.Text = "1.マスタ系"
        '
        'chkMs25
        '
        Me.chkMs25.AutoSize = true
        Me.chkMs25.Location = New System.Drawing.Point(183, 56)
        Me.chkMs25.Name = "chkMs25"
        Me.chkMs25.Size = New System.Drawing.Size(104, 22)
        Me.chkMs25.TabIndex = 106
        Me.chkMs25.Text = "src内部に記載"
        Me.chkMs25.UseVisualStyleBackColor = true
        '
        'chkMsKeiyakusyubetu
        '
        Me.chkMsKeiyakusyubetu.AutoSize = true
        Me.chkMsKeiyakusyubetu.Location = New System.Drawing.Point(153, 104)
        Me.chkMsKeiyakusyubetu.Name = "chkMsKeiyakusyubetu"
        Me.chkMsKeiyakusyubetu.Size = New System.Drawing.Size(104, 22)
        Me.chkMsKeiyakusyubetu.TabIndex = 104
        Me.chkMsKeiyakusyubetu.Text = "src内部に記載"
        Me.chkMsKeiyakusyubetu.UseVisualStyleBackColor = true
        '
        'chkMs23
        '
        Me.chkMs23.AutoSize = true
        Me.chkMs23.Location = New System.Drawing.Point(183, 10)
        Me.chkMs23.Name = "chkMs23"
        Me.chkMs23.Size = New System.Drawing.Size(104, 22)
        Me.chkMs23.TabIndex = 105
        Me.chkMs23.Text = "src内部に記載"
        Me.chkMs23.UseVisualStyleBackColor = true
        '
        'chkMs24
        '
        Me.chkMs24.AutoSize = true
        Me.chkMs24.Location = New System.Drawing.Point(183, 34)
        Me.chkMs24.Name = "chkMs24"
        Me.chkMs24.Size = New System.Drawing.Size(104, 22)
        Me.chkMs24.TabIndex = 104
        Me.chkMs24.Text = "src内部に記載"
        Me.chkMs24.UseVisualStyleBackColor = true
        '
        'chkMsOwevent
        '
        Me.chkMsOwevent.AutoSize = true
        Me.chkMsOwevent.Location = New System.Drawing.Point(50, 18)
        Me.chkMsOwevent.Name = "chkMsOwevent"
        Me.chkMsOwevent.Size = New System.Drawing.Size(104, 22)
        Me.chkMsOwevent.TabIndex = 82
        Me.chkMsOwevent.Text = "src内部に記載"
        Me.chkMsOwevent.UseVisualStyleBackColor = true
        '
        'chkMsYane
        '
        Me.chkMsYane.AutoSize = true
        Me.chkMsYane.Location = New System.Drawing.Point(50, 40)
        Me.chkMsYane.Name = "chkMsYane"
        Me.chkMsYane.Size = New System.Drawing.Size(104, 22)
        Me.chkMsYane.TabIndex = 93
        Me.chkMsYane.Text = "src内部に記載"
        Me.chkMsYane.UseVisualStyleBackColor = true
        '
        'chkMsKeikaikakunin
        '
        Me.chkMsKeikaikakunin.AutoSize = true
        Me.chkMsKeikaikakunin.Location = New System.Drawing.Point(50, 62)
        Me.chkMsKeikaikakunin.Name = "chkMsKeikaikakunin"
        Me.chkMsKeikaikakunin.Size = New System.Drawing.Size(104, 22)
        Me.chkMsKeikaikakunin.TabIndex = 84
        Me.chkMsKeikaikakunin.Text = "src内部に記載"
        Me.chkMsKeikaikakunin.UseVisualStyleBackColor = true
        '
        'chkMsKinyu
        '
        Me.chkMsKinyu.AutoSize = true
        Me.chkMsKinyu.Location = New System.Drawing.Point(50, 104)
        Me.chkMsKinyu.Name = "chkMsKinyu"
        Me.chkMsKinyu.Size = New System.Drawing.Size(104, 22)
        Me.chkMsKinyu.TabIndex = 95
        Me.chkMsKinyu.Text = "src内部に記載"
        Me.chkMsKinyu.UseVisualStyleBackColor = true
        '
        'chkMsKinyuten
        '
        Me.chkMsKinyuten.AutoSize = true
        Me.chkMsKinyuten.Location = New System.Drawing.Point(50, 82)
        Me.chkMsKinyuten.Name = "chkMsKinyuten"
        Me.chkMsKinyuten.Size = New System.Drawing.Size(104, 22)
        Me.chkMsKinyuten.TabIndex = 96
        Me.chkMsKinyuten.Text = "src内部に記載"
        Me.chkMsKinyuten.UseVisualStyleBackColor = true
        '
        'tabPageStart
        '
        Me.tabPageStart.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageStart.Controls.Add(Me.grpKFirstNaiyo)
        Me.tabPageStart.Controls.Add(Me.lblFirstDescription1)
        Me.tabPageStart.Controls.Add(Me.lblKFirstLabel)
        Me.tabPageStart.Controls.Add(Me.lblHFirstLabel)
        Me.tabPageStart.Controls.Add(Me.grpHFirstNaiyo)
        Me.tabPageStart.Controls.Add(Me.PictureBox2)
        Me.tabPageStart.Location = New System.Drawing.Point(4, 27)
        Me.tabPageStart.Name = "tabPageStart"
        Me.tabPageStart.Size = New System.Drawing.Size(912, 549)
        Me.tabPageStart.TabIndex = 13
        Me.tabPageStart.Text = "開始"
        '
        'grpKFirstNaiyo
        '
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi352)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi351)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi342)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi341)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi332)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi322)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi312)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi331)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi321)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi311)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi302)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi202)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi102)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi301)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi201)
        Me.grpKFirstNaiyo.Controls.Add(Me.lblHajimeKi101)
        Me.grpKFirstNaiyo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpKFirstNaiyo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.grpKFirstNaiyo.Location = New System.Drawing.Point(65, 115)
        Me.grpKFirstNaiyo.Name = "grpKFirstNaiyo"
        Me.grpKFirstNaiyo.Size = New System.Drawing.Size(767, 259)
        Me.grpKFirstNaiyo.TabIndex = 148
        Me.grpKFirstNaiyo.TabStop = false
        Me.grpKFirstNaiyo.Text = "【コンバート作業内容】"
        '
        'lblHajimeKi352
        '
        Me.lblHajimeKi352.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi352.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi352.Location = New System.Drawing.Point(290, 221)
        Me.lblHajimeKi352.Name = "lblHajimeKi352"
        Me.lblHajimeKi352.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeKi352.TabIndex = 64
        Me.lblHajimeKi352.Text = "… 検証用リスト出力などを行います。"
        Me.lblHajimeKi352.UseCompatibleTextRendering = true
        '
        'lblHajimeKi351
        '
        Me.lblHajimeKi351.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi351.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi351.Location = New System.Drawing.Point(124, 222)
        Me.lblHajimeKi351.Name = "lblHajimeKi351"
        Me.lblHajimeKi351.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeKi351.TabIndex = 63
        Me.lblHajimeKi351.Text = "　　※補助機能"
        Me.lblHajimeKi351.UseCompatibleTextRendering = true
        '
        'lblHajimeKi342
        '
        Me.lblHajimeKi342.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi342.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi342.Location = New System.Drawing.Point(290, 194)
        Me.lblHajimeKi342.Name = "lblHajimeKi342"
        Me.lblHajimeKi342.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeKi342.TabIndex = 62
        Me.lblHajimeKi342.Text = "… コンバートを実行した後の「賃貸革命10」データの調整を行います。"
        Me.lblHajimeKi342.UseCompatibleTextRendering = true
        '
        'lblHajimeKi341
        '
        Me.lblHajimeKi341.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi341.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi341.Location = New System.Drawing.Point(124, 195)
        Me.lblHajimeKi341.Name = "lblHajimeKi341"
        Me.lblHajimeKi341.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeKi341.TabIndex = 58
        Me.lblHajimeKi341.Text = "　　・事後作業"
        Me.lblHajimeKi341.UseCompatibleTextRendering = true
        '
        'lblHajimeKi332
        '
        Me.lblHajimeKi332.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi332.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi332.Location = New System.Drawing.Point(290, 168)
        Me.lblHajimeKi332.Name = "lblHajimeKi332"
        Me.lblHajimeKi332.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeKi332.TabIndex = 61
        Me.lblHajimeKi332.Text = "… 画像コンバートを行います。"
        Me.lblHajimeKi332.UseCompatibleTextRendering = true
        '
        'lblHajimeKi322
        '
        Me.lblHajimeKi322.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi322.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi322.Location = New System.Drawing.Point(290, 141)
        Me.lblHajimeKi322.Name = "lblHajimeKi322"
        Me.lblHajimeKi322.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeKi322.TabIndex = 60
        Me.lblHajimeKi322.Text = "… データコンバートを行います。"
        Me.lblHajimeKi322.UseCompatibleTextRendering = true
        '
        'lblHajimeKi312
        '
        Me.lblHajimeKi312.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi312.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi312.Location = New System.Drawing.Point(290, 114)
        Me.lblHajimeKi312.Name = "lblHajimeKi312"
        Me.lblHajimeKi312.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeKi312.TabIndex = 59
        Me.lblHajimeKi312.Text = "… コンバートを実行する前の「賃貸革命V7」データの調整を行います。"
        Me.lblHajimeKi312.UseCompatibleTextRendering = true
        '
        'lblHajimeKi331
        '
        Me.lblHajimeKi331.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi331.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi331.Location = New System.Drawing.Point(124, 168)
        Me.lblHajimeKi331.Name = "lblHajimeKi331"
        Me.lblHajimeKi331.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeKi331.TabIndex = 57
        Me.lblHajimeKi331.Text = "　　・画像コンバート"
        Me.lblHajimeKi331.UseCompatibleTextRendering = true
        '
        'lblHajimeKi321
        '
        Me.lblHajimeKi321.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi321.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi321.Location = New System.Drawing.Point(124, 141)
        Me.lblHajimeKi321.Name = "lblHajimeKi321"
        Me.lblHajimeKi321.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeKi321.TabIndex = 56
        Me.lblHajimeKi321.Text = "　　・データコンバート"
        Me.lblHajimeKi321.UseCompatibleTextRendering = true
        '
        'lblHajimeKi311
        '
        Me.lblHajimeKi311.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi311.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi311.Location = New System.Drawing.Point(124, 114)
        Me.lblHajimeKi311.Name = "lblHajimeKi311"
        Me.lblHajimeKi311.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeKi311.TabIndex = 55
        Me.lblHajimeKi311.Text = "　　・事前作業"
        Me.lblHajimeKi311.UseCompatibleTextRendering = true
        '
        'lblHajimeKi302
        '
        Me.lblHajimeKi302.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi302.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi302.Location = New System.Drawing.Point(124, 85)
        Me.lblHajimeKi302.Name = "lblHajimeKi302"
        Me.lblHajimeKi302.Size = New System.Drawing.Size(620, 20)
        Me.lblHajimeKi302.TabIndex = 54
        Me.lblHajimeKi302.Text = "… コンバート作業を行います。(以下の内容から選択)"
        Me.lblHajimeKi302.UseCompatibleTextRendering = true
        '
        'lblHajimeKi202
        '
        Me.lblHajimeKi202.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi202.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi202.Location = New System.Drawing.Point(124, 57)
        Me.lblHajimeKi202.Name = "lblHajimeKi202"
        Me.lblHajimeKi202.Size = New System.Drawing.Size(620, 20)
        Me.lblHajimeKi202.TabIndex = 53
        Me.lblHajimeKi202.Text = "… 「ログファイル格納パス」の設定など、簡単な初期設定を行います。"
        Me.lblHajimeKi202.UseCompatibleTextRendering = true
        '
        'lblHajimeKi102
        '
        Me.lblHajimeKi102.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi102.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi102.Location = New System.Drawing.Point(124, 28)
        Me.lblHajimeKi102.Name = "lblHajimeKi102"
        Me.lblHajimeKi102.Size = New System.Drawing.Size(620, 20)
        Me.lblHajimeKi102.TabIndex = 52
        Me.lblHajimeKi102.Text = "… 「賃貸革命V7」と「賃貸革命10」のデータベースの接続設定を行います。"
        Me.lblHajimeKi102.UseCompatibleTextRendering = true
        '
        'lblHajimeKi301
        '
        Me.lblHajimeKi301.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi301.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi301.Location = New System.Drawing.Point(18, 85)
        Me.lblHajimeKi301.Name = "lblHajimeKi301"
        Me.lblHajimeKi301.Size = New System.Drawing.Size(100, 20)
        Me.lblHajimeKi301.TabIndex = 51
        Me.lblHajimeKi301.Text = "　作業選択"
        Me.lblHajimeKi301.UseCompatibleTextRendering = true
        '
        'lblHajimeKi201
        '
        Me.lblHajimeKi201.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi201.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi201.Location = New System.Drawing.Point(18, 57)
        Me.lblHajimeKi201.Name = "lblHajimeKi201"
        Me.lblHajimeKi201.Size = New System.Drawing.Size(100, 20)
        Me.lblHajimeKi201.TabIndex = 50
        Me.lblHajimeKi201.Text = "　初期設定"
        Me.lblHajimeKi201.UseCompatibleTextRendering = true
        '
        'lblHajimeKi101
        '
        Me.lblHajimeKi101.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeKi101.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeKi101.Location = New System.Drawing.Point(18, 29)
        Me.lblHajimeKi101.Name = "lblHajimeKi101"
        Me.lblHajimeKi101.Size = New System.Drawing.Size(100, 20)
        Me.lblHajimeKi101.TabIndex = 49
        Me.lblHajimeKi101.Text = "　接続設定"
        Me.lblHajimeKi101.UseCompatibleTextRendering = true
        '
        'lblFirstDescription1
        '
        Me.lblFirstDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblFirstDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblFirstDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblFirstDescription1.Location = New System.Drawing.Point(35, 49)
        Me.lblFirstDescription1.Name = "lblFirstDescription1"
        Me.lblFirstDescription1.Size = New System.Drawing.Size(820, 35)
        Me.lblFirstDescription1.TabIndex = 4
        Me.lblFirstDescription1.Text = "「次へ」ボタンを押して下さい。"
        Me.lblFirstDescription1.UseCompatibleTextRendering = true
        '
        'lblKFirstLabel
        '
        Me.lblKFirstLabel.BackColor = System.Drawing.SystemColors.Menu
        Me.lblKFirstLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKFirstLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.lblKFirstLabel.Location = New System.Drawing.Point(30, 20)
        Me.lblKFirstLabel.Name = "lblKFirstLabel"
        Me.lblKFirstLabel.Size = New System.Drawing.Size(820, 80)
        Me.lblKFirstLabel.TabIndex = 3
        Me.lblKFirstLabel.Text = "「賃貸革命V7」から「賃貸革命10」のコンバート作業を行います。"
        Me.lblKFirstLabel.UseCompatibleTextRendering = true
        '
        'lblHFirstLabel
        '
        Me.lblHFirstLabel.BackColor = System.Drawing.SystemColors.Menu
        Me.lblHFirstLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHFirstLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.lblHFirstLabel.Location = New System.Drawing.Point(35, 20)
        Me.lblHFirstLabel.Name = "lblHFirstLabel"
        Me.lblHFirstLabel.Size = New System.Drawing.Size(820, 80)
        Me.lblHFirstLabel.TabIndex = 149
        Me.lblHFirstLabel.Text = "「中間ファイル」から「賃貸革命10」のコンバート作業を行います。"
        Me.lblHFirstLabel.UseCompatibleTextRendering = true
        '
        'grpHFirstNaiyo
        '
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH312)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH311)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH342)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH322)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH341)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH321)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH302)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH202)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH102)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH301)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH201)
        Me.grpHFirstNaiyo.Controls.Add(Me.lblHajimeH101)
        Me.grpHFirstNaiyo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpHFirstNaiyo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.grpHFirstNaiyo.Location = New System.Drawing.Point(65, 115)
        Me.grpHFirstNaiyo.Name = "grpHFirstNaiyo"
        Me.grpHFirstNaiyo.Size = New System.Drawing.Size(767, 259)
        Me.grpHFirstNaiyo.TabIndex = 149
        Me.grpHFirstNaiyo.TabStop = false
        Me.grpHFirstNaiyo.Text = "【コンバート作業内容】"
        '
        'lblHajimeH312
        '
        Me.lblHajimeH312.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH312.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH312.Location = New System.Drawing.Point(290, 114)
        Me.lblHajimeH312.Name = "lblHajimeH312"
        Me.lblHajimeH312.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeH312.TabIndex = 64
        Me.lblHajimeH312.Text = "… コンバートを実行する前の作業を行います。"
        Me.lblHajimeH312.UseCompatibleTextRendering = true
        '
        'lblHajimeH311
        '
        Me.lblHajimeH311.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH311.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH311.Location = New System.Drawing.Point(124, 114)
        Me.lblHajimeH311.Name = "lblHajimeH311"
        Me.lblHajimeH311.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeH311.TabIndex = 63
        Me.lblHajimeH311.Text = "　　・事前作業"
        Me.lblHajimeH311.UseCompatibleTextRendering = true
        '
        'lblHajimeH342
        '
        Me.lblHajimeH342.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH342.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH342.Location = New System.Drawing.Point(290, 168)
        Me.lblHajimeH342.Name = "lblHajimeH342"
        Me.lblHajimeH342.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeH342.TabIndex = 62
        Me.lblHajimeH342.Text = "… コンバートを実行した後の「賃貸革命10」データの調整を行います。"
        Me.lblHajimeH342.UseCompatibleTextRendering = true
        '
        'lblHajimeH322
        '
        Me.lblHajimeH322.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH322.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH322.Location = New System.Drawing.Point(290, 141)
        Me.lblHajimeH322.Name = "lblHajimeH322"
        Me.lblHajimeH322.Size = New System.Drawing.Size(454, 20)
        Me.lblHajimeH322.TabIndex = 60
        Me.lblHajimeH322.Text = "… 中間ファイルのチェックとデータコンバートを行います。"
        Me.lblHajimeH322.UseCompatibleTextRendering = true
        '
        'lblHajimeH341
        '
        Me.lblHajimeH341.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH341.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH341.Location = New System.Drawing.Point(124, 168)
        Me.lblHajimeH341.Name = "lblHajimeH341"
        Me.lblHajimeH341.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeH341.TabIndex = 58
        Me.lblHajimeH341.Text = "　　・事後作業"
        Me.lblHajimeH341.UseCompatibleTextRendering = true
        '
        'lblHajimeH321
        '
        Me.lblHajimeH321.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH321.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH321.Location = New System.Drawing.Point(124, 141)
        Me.lblHajimeH321.Name = "lblHajimeH321"
        Me.lblHajimeH321.Size = New System.Drawing.Size(160, 20)
        Me.lblHajimeH321.TabIndex = 56
        Me.lblHajimeH321.Text = "　　・データコンバート"
        Me.lblHajimeH321.UseCompatibleTextRendering = true
        '
        'lblHajimeH302
        '
        Me.lblHajimeH302.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH302.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH302.Location = New System.Drawing.Point(124, 85)
        Me.lblHajimeH302.Name = "lblHajimeH302"
        Me.lblHajimeH302.Size = New System.Drawing.Size(620, 20)
        Me.lblHajimeH302.TabIndex = 54
        Me.lblHajimeH302.Text = "… コンバート作業を行います。(以下の内容から選択)"
        Me.lblHajimeH302.UseCompatibleTextRendering = true
        '
        'lblHajimeH202
        '
        Me.lblHajimeH202.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH202.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH202.Location = New System.Drawing.Point(124, 57)
        Me.lblHajimeH202.Name = "lblHajimeH202"
        Me.lblHajimeH202.Size = New System.Drawing.Size(620, 20)
        Me.lblHajimeH202.TabIndex = 53
        Me.lblHajimeH202.Text = "… 「ログファイル格納パス」の設定など、簡単な初期設定を行います。"
        Me.lblHajimeH202.UseCompatibleTextRendering = true
        '
        'lblHajimeH102
        '
        Me.lblHajimeH102.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH102.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH102.Location = New System.Drawing.Point(124, 28)
        Me.lblHajimeH102.Name = "lblHajimeH102"
        Me.lblHajimeH102.Size = New System.Drawing.Size(620, 20)
        Me.lblHajimeH102.TabIndex = 52
        Me.lblHajimeH102.Text = "… 「賃貸革命10」のデータベースの接続設定を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)
        Me.lblHajimeH102.UseCompatibleTextRendering = true
        '
        'lblHajimeH301
        '
        Me.lblHajimeH301.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH301.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH301.Location = New System.Drawing.Point(18, 85)
        Me.lblHajimeH301.Name = "lblHajimeH301"
        Me.lblHajimeH301.Size = New System.Drawing.Size(100, 20)
        Me.lblHajimeH301.TabIndex = 51
        Me.lblHajimeH301.Text = "　作業選択"
        Me.lblHajimeH301.UseCompatibleTextRendering = true
        '
        'lblHajimeH201
        '
        Me.lblHajimeH201.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH201.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH201.Location = New System.Drawing.Point(18, 57)
        Me.lblHajimeH201.Name = "lblHajimeH201"
        Me.lblHajimeH201.Size = New System.Drawing.Size(100, 20)
        Me.lblHajimeH201.TabIndex = 50
        Me.lblHajimeH201.Text = "　初期設定"
        Me.lblHajimeH201.UseCompatibleTextRendering = true
        '
        'lblHajimeH101
        '
        Me.lblHajimeH101.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHajimeH101.ForeColor = System.Drawing.Color.Black
        Me.lblHajimeH101.Location = New System.Drawing.Point(18, 29)
        Me.lblHajimeH101.Name = "lblHajimeH101"
        Me.lblHajimeH101.Size = New System.Drawing.Size(100, 20)
        Me.lblHajimeH101.TabIndex = 49
        Me.lblHajimeH101.Text = "　接続設定"
        Me.lblHajimeH101.UseCompatibleTextRendering = true
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"),System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(698, 390)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(191, 138)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 147
        Me.PictureBox2.TabStop = false
        '
        'tabPageSession
        '
        Me.tabPageSession.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageSession.Controls.Add(Me.grpTimeOut)
        Me.tabPageSession.Controls.Add(Me.lblSessionCaution)
        Me.tabPageSession.Controls.Add(Me.lblKSessionDescription1)
        Me.tabPageSession.Controls.Add(Me.lblHSessionDescription1)
        Me.tabPageSession.Controls.Add(Me.Label36)
        Me.tabPageSession.Controls.Add(Me.Label35)
        Me.tabPageSession.Controls.Add(Me.lblNetworklib)
        Me.tabPageSession.Controls.Add(Me.Label30)
        Me.tabPageSession.Controls.Add(Me.Label13)
        Me.tabPageSession.Controls.Add(Me.grp10ConnectInfo)
        Me.tabPageSession.Controls.Add(Me.grpV7ConnectInfo)
        Me.tabPageSession.Controls.Add(Me.btnDefConInfoRead)
        Me.tabPageSession.Controls.Add(Me.btnConnectTest)
        Me.tabPageSession.Location = New System.Drawing.Point(4, 27)
        Me.tabPageSession.Name = "tabPageSession"
        Me.tabPageSession.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageSession.Size = New System.Drawing.Size(912, 549)
        Me.tabPageSession.TabIndex = 1
        Me.tabPageSession.Text = " 接続設定"
        '
        'grpTimeOut
        '
        Me.grpTimeOut.Controls.Add(Me.txtTimeOut)
        Me.grpTimeOut.Controls.Add(Me.lblTimeOutSec)
        Me.grpTimeOut.Location = New System.Drawing.Point(769, 20)
        Me.grpTimeOut.Name = "grpTimeOut"
        Me.grpTimeOut.Size = New System.Drawing.Size(118, 64)
        Me.grpTimeOut.TabIndex = 14
        Me.grpTimeOut.TabStop = false
        Me.grpTimeOut.Text = " タイムアウト値 "
        '
        'txtTimeOut
        '
        Me.txtTimeOut.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtTimeOut.Location = New System.Drawing.Point(33, 24)
        Me.txtTimeOut.Name = "txtTimeOut"
        Me.txtTimeOut.Size = New System.Drawing.Size(43, 27)
        Me.txtTimeOut.TabIndex = 10
        Me.txtTimeOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblTimeOutSec
        '
        Me.lblTimeOutSec.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblTimeOutSec.Location = New System.Drawing.Point(82, 28)
        Me.lblTimeOutSec.Name = "lblTimeOutSec"
        Me.lblTimeOutSec.Size = New System.Drawing.Size(29, 19)
        Me.lblTimeOutSec.TabIndex = 11
        Me.lblTimeOutSec.Text = "秒"
        Me.lblTimeOutSec.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblTimeOutSec.UseCompatibleTextRendering = true
        '
        'lblSessionCaution
        '
        Me.lblSessionCaution.BackColor = System.Drawing.SystemColors.Menu
        Me.lblSessionCaution.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblSessionCaution.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblSessionCaution.Location = New System.Drawing.Point(30, 90)
        Me.lblSessionCaution.Name = "lblSessionCaution"
        Me.lblSessionCaution.Size = New System.Drawing.Size(820, 24)
        Me.lblSessionCaution.TabIndex = 1
        Me.lblSessionCaution.Text = "※正常接続が確認できない場合、[次へ] に進むことはできません。(正常接続できない場合はサポートへお問い合わせ下さい)"
        Me.lblSessionCaution.UseCompatibleTextRendering = true
        '
        'lblKSessionDescription1
        '
        Me.lblKSessionDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblKSessionDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKSessionDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblKSessionDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblKSessionDescription1.Name = "lblKSessionDescription1"
        Me.lblKSessionDescription1.Size = New System.Drawing.Size(820, 80)
        Me.lblKSessionDescription1.TabIndex = 0
        Me.lblKSessionDescription1.Text = "データベースの接続設定を行います。([初期値読込] を押すと接続情報の初期設定値を読み込みます)"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"「賃貸革命V7」と「賃貸革命10」の接続情報を設定し、[接続確"& _ 
    "認] ボタンを押して下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"正常接続が確認できたら [次へ] ボタンを押して下さい。"
        Me.lblKSessionDescription1.UseCompatibleTextRendering = true
        '
        'lblHSessionDescription1
        '
        Me.lblHSessionDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblHSessionDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHSessionDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblHSessionDescription1.Location = New System.Drawing.Point(30, 21)
        Me.lblHSessionDescription1.Name = "lblHSessionDescription1"
        Me.lblHSessionDescription1.Size = New System.Drawing.Size(820, 80)
        Me.lblHSessionDescription1.TabIndex = 15
        Me.lblHSessionDescription1.Text = "データベースの接続設定を行います。([初期値読込] を押すと接続情報の初期設定値を読み込みます)"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"「賃貸革命10」の接続情報を設定し、[接続確認] ボタンを押し"& _ 
    "て下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"正常接続が確認できたら [次へ] ボタンを押して下さい。"
        Me.lblHSessionDescription1.UseCompatibleTextRendering = true
        '
        'Label36
        '
        Me.Label36.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label36.Location = New System.Drawing.Point(34, 401)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(140, 20)
        Me.Label36.TabIndex = 5
        Me.Label36.Text = "パスワード"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label36.UseCompatibleTextRendering = true
        '
        'Label35
        '
        Me.Label35.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label35.Location = New System.Drawing.Point(34, 361)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(140, 20)
        Me.Label35.TabIndex = 4
        Me.Label35.Text = "ログイン"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label35.UseCompatibleTextRendering = true
        '
        'lblNetworklib
        '
        Me.lblNetworklib.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblNetworklib.Location = New System.Drawing.Point(34, 444)
        Me.lblNetworklib.Name = "lblNetworklib"
        Me.lblNetworklib.Size = New System.Drawing.Size(140, 20)
        Me.lblNetworklib.TabIndex = 6
        Me.lblNetworklib.Text = "ネットワークライブラリ"
        Me.lblNetworklib.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblNetworklib.UseCompatibleTextRendering = true
        '
        'Label30
        '
        Me.Label30.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label30.Location = New System.Drawing.Point(34, 281)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(140, 20)
        Me.Label30.TabIndex = 2
        Me.Label30.Text = "サーバー名"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label30.UseCompatibleTextRendering = true
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label13.Location = New System.Drawing.Point(34, 321)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(140, 20)
        Me.Label13.TabIndex = 3
        Me.Label13.Text = "カタログ名"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label13.UseCompatibleTextRendering = true
        '
        'grp10ConnectInfo
        '
        Me.grp10ConnectInfo.Controls.Add(Me.cmbV10Networklib)
        Me.grp10ConnectInfo.Controls.Add(Me.grpV10Authent)
        Me.grp10ConnectInfo.Controls.Add(Me.txtV10Pass)
        Me.grp10ConnectInfo.Controls.Add(Me.txtV10User)
        Me.grp10ConnectInfo.Controls.Add(Me.txtV10Catalog)
        Me.grp10ConnectInfo.Controls.Add(Me.txtV10Server)
        Me.grp10ConnectInfo.Controls.Add(Me.lblV10)
        Me.grp10ConnectInfo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grp10ConnectInfo.Location = New System.Drawing.Point(529, 159)
        Me.grp10ConnectInfo.Name = "grp10ConnectInfo"
        Me.grp10ConnectInfo.Size = New System.Drawing.Size(300, 330)
        Me.grp10ConnectInfo.TabIndex = 8
        Me.grp10ConnectInfo.TabStop = false
        Me.grp10ConnectInfo.Text = "【移行先接続情報】"
        '
        'cmbV10Networklib
        '
        Me.cmbV10Networklib.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbV10Networklib.FormattingEnabled = true
        Me.cmbV10Networklib.Location = New System.Drawing.Point(10, 280)
        Me.cmbV10Networklib.Name = "cmbV10Networklib"
        Me.cmbV10Networklib.Size = New System.Drawing.Size(280, 28)
        Me.cmbV10Networklib.TabIndex = 7
        '
        'grpV10Authent
        '
        Me.grpV10Authent.Controls.Add(Me.optV10Authent2)
        Me.grpV10Authent.Controls.Add(Me.optV10Authent1)
        Me.grpV10Authent.Location = New System.Drawing.Point(6, 56)
        Me.grpV10Authent.Name = "grpV10Authent"
        Me.grpV10Authent.Size = New System.Drawing.Size(288, 50)
        Me.grpV10Authent.TabIndex = 1
        Me.grpV10Authent.TabStop = false
        '
        'optV10Authent2
        '
        Me.optV10Authent2.AutoSize = true
        Me.optV10Authent2.Location = New System.Drawing.Point(160, 18)
        Me.optV10Authent2.Name = "optV10Authent2"
        Me.optV10Authent2.Size = New System.Drawing.Size(114, 24)
        Me.optV10Authent2.TabIndex = 1
        Me.optV10Authent2.Text = "Windows 認証"
        Me.optV10Authent2.UseVisualStyleBackColor = true
        '
        'optV10Authent1
        '
        Me.optV10Authent1.AutoSize = true
        Me.optV10Authent1.Checked = true
        Me.optV10Authent1.Location = New System.Drawing.Point(12, 18)
        Me.optV10Authent1.Name = "optV10Authent1"
        Me.optV10Authent1.Size = New System.Drawing.Size(123, 24)
        Me.optV10Authent1.TabIndex = 0
        Me.optV10Authent1.TabStop = true
        Me.optV10Authent1.Text = "SQLServer 認証"
        Me.optV10Authent1.UseVisualStyleBackColor = true
        '
        'txtV10Pass
        '
        Me.txtV10Pass.Location = New System.Drawing.Point(10, 238)
        Me.txtV10Pass.Name = "txtV10Pass"
        Me.txtV10Pass.Size = New System.Drawing.Size(280, 27)
        Me.txtV10Pass.TabIndex = 5
        Me.txtV10Pass.Text = "***************"
        '
        'txtV10User
        '
        Me.txtV10User.Location = New System.Drawing.Point(10, 201)
        Me.txtV10User.Name = "txtV10User"
        Me.txtV10User.Size = New System.Drawing.Size(280, 27)
        Me.txtV10User.TabIndex = 4
        Me.txtV10User.Text = "sa"
        '
        'txtV10Catalog
        '
        Me.txtV10Catalog.Location = New System.Drawing.Point(10, 160)
        Me.txtV10Catalog.Name = "txtV10Catalog"
        Me.txtV10Catalog.Size = New System.Drawing.Size(280, 27)
        Me.txtV10Catalog.TabIndex = 3
        Me.txtV10Catalog.Text = "fk8db"
        '
        'txtV10Server
        '
        Me.txtV10Server.Location = New System.Drawing.Point(10, 120)
        Me.txtV10Server.Name = "txtV10Server"
        Me.txtV10Server.Size = New System.Drawing.Size(280, 27)
        Me.txtV10Server.TabIndex = 2
        Me.txtV10Server.Text = "PC-NJC\SQL2012"
        '
        'lblV10
        '
        Me.lblV10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblV10.Image = CType(resources.GetObject("lblV10.Image"),System.Drawing.Image)
        Me.lblV10.Location = New System.Drawing.Point(6, 30)
        Me.lblV10.Name = "lblV10"
        Me.lblV10.Size = New System.Drawing.Size(288, 24)
        Me.lblV10.TabIndex = 0
        Me.lblV10.Text = "賃貸革命10"
        Me.lblV10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpV7ConnectInfo
        '
        Me.grpV7ConnectInfo.Controls.Add(Me.cmbV7Networklib)
        Me.grpV7ConnectInfo.Controls.Add(Me.grpV7Authent)
        Me.grpV7ConnectInfo.Controls.Add(Me.txtV7Pass)
        Me.grpV7ConnectInfo.Controls.Add(Me.txtV7User)
        Me.grpV7ConnectInfo.Controls.Add(Me.txtV7Catalog)
        Me.grpV7ConnectInfo.Controls.Add(Me.txtV7Server)
        Me.grpV7ConnectInfo.Controls.Add(Me.lblV7)
        Me.grpV7ConnectInfo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpV7ConnectInfo.Location = New System.Drawing.Point(190, 158)
        Me.grpV7ConnectInfo.Name = "grpV7ConnectInfo"
        Me.grpV7ConnectInfo.Size = New System.Drawing.Size(300, 330)
        Me.grpV7ConnectInfo.TabIndex = 7
        Me.grpV7ConnectInfo.TabStop = false
        Me.grpV7ConnectInfo.Text = "【移行元接続情報】"
        '
        'cmbV7Networklib
        '
        Me.cmbV7Networklib.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbV7Networklib.FormattingEnabled = true
        Me.cmbV7Networklib.Location = New System.Drawing.Point(10, 280)
        Me.cmbV7Networklib.Name = "cmbV7Networklib"
        Me.cmbV7Networklib.Size = New System.Drawing.Size(280, 28)
        Me.cmbV7Networklib.TabIndex = 6
        '
        'grpV7Authent
        '
        Me.grpV7Authent.Controls.Add(Me.optV7Authent2)
        Me.grpV7Authent.Controls.Add(Me.optV7Authent1)
        Me.grpV7Authent.Location = New System.Drawing.Point(6, 56)
        Me.grpV7Authent.Name = "grpV7Authent"
        Me.grpV7Authent.Size = New System.Drawing.Size(288, 50)
        Me.grpV7Authent.TabIndex = 1
        Me.grpV7Authent.TabStop = false
        '
        'optV7Authent2
        '
        Me.optV7Authent2.AutoSize = true
        Me.optV7Authent2.Location = New System.Drawing.Point(160, 18)
        Me.optV7Authent2.Name = "optV7Authent2"
        Me.optV7Authent2.Size = New System.Drawing.Size(114, 24)
        Me.optV7Authent2.TabIndex = 1
        Me.optV7Authent2.Text = "Windows 認証"
        Me.optV7Authent2.UseVisualStyleBackColor = true
        '
        'optV7Authent1
        '
        Me.optV7Authent1.AutoSize = true
        Me.optV7Authent1.Checked = true
        Me.optV7Authent1.Location = New System.Drawing.Point(12, 18)
        Me.optV7Authent1.Name = "optV7Authent1"
        Me.optV7Authent1.Size = New System.Drawing.Size(123, 24)
        Me.optV7Authent1.TabIndex = 0
        Me.optV7Authent1.TabStop = true
        Me.optV7Authent1.Text = "SQLServer 認証"
        Me.optV7Authent1.UseVisualStyleBackColor = true
        '
        'txtV7Pass
        '
        Me.txtV7Pass.Location = New System.Drawing.Point(10, 240)
        Me.txtV7Pass.Name = "txtV7Pass"
        Me.txtV7Pass.Size = New System.Drawing.Size(280, 27)
        Me.txtV7Pass.TabIndex = 5
        Me.txtV7Pass.Text = "***************"
        '
        'txtV7User
        '
        Me.txtV7User.Location = New System.Drawing.Point(10, 200)
        Me.txtV7User.Name = "txtV7User"
        Me.txtV7User.Size = New System.Drawing.Size(280, 27)
        Me.txtV7User.TabIndex = 4
        Me.txtV7User.Text = "sa"
        '
        'txtV7Catalog
        '
        Me.txtV7Catalog.Location = New System.Drawing.Point(10, 160)
        Me.txtV7Catalog.Name = "txtV7Catalog"
        Me.txtV7Catalog.Size = New System.Drawing.Size(280, 27)
        Me.txtV7Catalog.TabIndex = 3
        Me.txtV7Catalog.Text = "fk5dtsql"
        '
        'txtV7Server
        '
        Me.txtV7Server.Location = New System.Drawing.Point(10, 120)
        Me.txtV7Server.Name = "txtV7Server"
        Me.txtV7Server.Size = New System.Drawing.Size(280, 27)
        Me.txtV7Server.TabIndex = 2
        Me.txtV7Server.Text = "PC-NJC\SQL2008"
        '
        'lblV7
        '
        Me.lblV7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblV7.Image = CType(resources.GetObject("lblV7.Image"),System.Drawing.Image)
        Me.lblV7.Location = New System.Drawing.Point(10, 30)
        Me.lblV7.Name = "lblV7"
        Me.lblV7.Size = New System.Drawing.Size(284, 24)
        Me.lblV7.TabIndex = 0
        Me.lblV7.Text = "賃貸革命V7"
        Me.lblV7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnDefConInfoRead
        '
        Me.btnDefConInfoRead.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnDefConInfoRead.Image = CType(resources.GetObject("btnDefConInfoRead.Image"),System.Drawing.Image)
        Me.btnDefConInfoRead.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDefConInfoRead.Location = New System.Drawing.Point(632, 506)
        Me.btnDefConInfoRead.Name = "btnDefConInfoRead"
        Me.btnDefConInfoRead.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnDefConInfoRead.Size = New System.Drawing.Size(120, 30)
        Me.btnDefConInfoRead.TabIndex = 12
        Me.btnDefConInfoRead.Text = "設定値再読込"
        Me.btnDefConInfoRead.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnDefConInfoRead.UseVisualStyleBackColor = true
        '
        'btnConnectTest
        '
        Me.btnConnectTest.Font = New System.Drawing.Font("メイリオ", 11.25!)
        Me.btnConnectTest.Image = CType(resources.GetObject("btnConnectTest.Image"),System.Drawing.Image)
        Me.btnConnectTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnConnectTest.Location = New System.Drawing.Point(769, 506)
        Me.btnConnectTest.Name = "btnConnectTest"
        Me.btnConnectTest.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnConnectTest.Size = New System.Drawing.Size(120, 30)
        Me.btnConnectTest.TabIndex = 13
        Me.btnConnectTest.Text = " 接続確認"
        Me.btnConnectTest.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnConnectTest.UseVisualStyleBackColor = true
        '
        'tabPageSyoki
        '
        Me.tabPageSyoki.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageSyoki.Controls.Add(Me.pnlConvertType)
        Me.tabPageSyoki.Controls.Add(Me.pnlUnyoKaisi)
        Me.tabPageSyoki.Controls.Add(Me.pnlOptSelect)
        Me.tabPageSyoki.Controls.Add(Me.pnlUserName)
        Me.tabPageSyoki.Controls.Add(Me.pnlLogPath)
        Me.tabPageSyoki.Controls.Add(Me.lblSyokiDescription1)
        Me.tabPageSyoki.Location = New System.Drawing.Point(4, 27)
        Me.tabPageSyoki.Name = "tabPageSyoki"
        Me.tabPageSyoki.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageSyoki.Size = New System.Drawing.Size(912, 549)
        Me.tabPageSyoki.TabIndex = 0
        Me.tabPageSyoki.Text = " 初期設定"
        '
        'pnlConvertType
        '
        Me.pnlConvertType.Controls.Add(Me.Label106)
        Me.pnlConvertType.Controls.Add(Me.pnlCVType)
        Me.pnlConvertType.Controls.Add(Me.Label107)
        Me.pnlConvertType.Location = New System.Drawing.Point(492, 167)
        Me.pnlConvertType.Name = "pnlConvertType"
        Me.pnlConvertType.Size = New System.Drawing.Size(414, 288)
        Me.pnlConvertType.TabIndex = 25
        '
        'Label106
        '
        Me.Label106.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label106.Location = New System.Drawing.Point(5, 4)
        Me.Label106.Name = "Label106"
        Me.Label106.Size = New System.Drawing.Size(128, 19)
        Me.Label106.TabIndex = 167
        Me.Label106.Text = "コンバートタイプ"
        Me.Label106.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label106.UseCompatibleTextRendering = true
        '
        'pnlCVType
        '
        Me.pnlCVType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCVType.Controls.Add(Me.Label3)
        Me.pnlCVType.Controls.Add(Me.optCVNew)
        Me.pnlCVType.Controls.Add(Me.Label4)
        Me.pnlCVType.Controls.Add(Me.optCVAdd)
        Me.pnlCVType.Location = New System.Drawing.Point(23, 44)
        Me.pnlCVType.Name = "pnlCVType"
        Me.pnlCVType.Size = New System.Drawing.Size(379, 235)
        Me.pnlCVType.TabIndex = 169
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label3.Location = New System.Drawing.Point(36, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(336, 108)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "賃貸革命10に登録されているデータを残したまま、データの追加登録(コンバート)を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※登録データの初期化・上書きは行いません。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(例：物件情報データの"& _ 
    "追加登録において、キーとなる""物件No""と同値のデータが賃貸革命10へ既に登録されている場合、キー以外の項目が異なっていても移行しません)"
        Me.Label3.UseCompatibleTextRendering = true
        '
        'optCVNew
        '
        Me.optCVNew.AutoSize = true
        Me.optCVNew.Checked = true
        Me.optCVNew.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.optCVNew.Location = New System.Drawing.Point(12, 3)
        Me.optCVNew.Name = "optCVNew"
        Me.optCVNew.Size = New System.Drawing.Size(110, 22)
        Me.optCVNew.TabIndex = 0
        Me.optCVNew.TabStop = true
        Me.optCVNew.Text = "新規コンバート"
        Me.optCVNew.UseVisualStyleBackColor = true
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label4.Location = New System.Drawing.Point(36, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(336, 77)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "賃貸革命10の登録データを初期化して、データの新規登録(コンバート)を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※初期化の対象は、""対象項目選択""画面で選択した項目のみとなります。"
        Me.Label4.UseCompatibleTextRendering = true
        '
        'optCVAdd
        '
        Me.optCVAdd.AutoSize = true
        Me.optCVAdd.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.optCVAdd.Location = New System.Drawing.Point(12, 98)
        Me.optCVAdd.Name = "optCVAdd"
        Me.optCVAdd.Size = New System.Drawing.Size(110, 22)
        Me.optCVAdd.TabIndex = 2
        Me.optCVAdd.Text = "追加コンバート"
        Me.optCVAdd.UseVisualStyleBackColor = true
        '
        'Label107
        '
        Me.Label107.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label107.Location = New System.Drawing.Point(25, 23)
        Me.Label107.Name = "Label107"
        Me.Label107.Size = New System.Drawing.Size(284, 20)
        Me.Label107.TabIndex = 168
        Me.Label107.Text = "コンバート方法を選択して下さい。"
        Me.Label107.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label107.UseCompatibleTextRendering = true
        '
        'pnlUnyoKaisi
        '
        Me.pnlUnyoKaisi.Controls.Add(Me.txtUnyoYYYYMM)
        Me.pnlUnyoKaisi.Controls.Add(Me.Label91)
        Me.pnlUnyoKaisi.Controls.Add(Me.Label98)
        Me.pnlUnyoKaisi.Controls.Add(Me.Label137)
        Me.pnlUnyoKaisi.Location = New System.Drawing.Point(46, 79)
        Me.pnlUnyoKaisi.Name = "pnlUnyoKaisi"
        Me.pnlUnyoKaisi.Size = New System.Drawing.Size(414, 82)
        Me.pnlUnyoKaisi.TabIndex = 24
        '
        'txtUnyoYYYYMM
        '
        Me.txtUnyoYYYYMM.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtUnyoYYYYMM.Location = New System.Drawing.Point(322, 24)
        Me.txtUnyoYYYYMM.Name = "txtUnyoYYYYMM"
        Me.txtUnyoYYYYMM.Size = New System.Drawing.Size(88, 25)
        Me.txtUnyoYYYYMM.TabIndex = 4
        Me.txtUnyoYYYYMM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label91
        '
        Me.Label91.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label91.Location = New System.Drawing.Point(3, 4)
        Me.Label91.Name = "Label91"
        Me.Label91.Size = New System.Drawing.Size(128, 19)
        Me.Label91.TabIndex = 1
        Me.Label91.Text = "運用開始年月"
        Me.Label91.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label91.UseCompatibleTextRendering = true
        '
        'Label98
        '
        Me.Label98.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label98.Location = New System.Drawing.Point(23, 23)
        Me.Label98.Name = "Label98"
        Me.Label98.Size = New System.Drawing.Size(299, 20)
        Me.Label98.TabIndex = 2
        Me.Label98.Text = "賃貸革命10の運用開始時の年月を設定して下さい。"
        Me.Label98.UseCompatibleTextRendering = true
        '
        'Label137
        '
        Me.Label137.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label137.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label137.Location = New System.Drawing.Point(23, 43)
        Me.Label137.Name = "Label137"
        Me.Label137.Size = New System.Drawing.Size(293, 35)
        Me.Label137.TabIndex = 3
        Me.Label137.Text = "※必ず設定する必要があります。未設定の場合は"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　本プログラム実行時の年月が設定されます。"
        Me.Label137.UseCompatibleTextRendering = true
        '
        'pnlOptSelect
        '
        Me.pnlOptSelect.Controls.Add(Me.lblKOpSeleTitle)
        Me.pnlOptSelect.Controls.Add(Me.pnlOptionSelect)
        Me.pnlOptSelect.Controls.Add(Me.lblHOpSeleTitle)
        Me.pnlOptSelect.Controls.Add(Me.lblHOpSeleNaiyo)
        Me.pnlOptSelect.Controls.Add(Me.lblSOpSeleNaiyo)
        Me.pnlOptSelect.Controls.Add(Me.lblKOpSeleNaiyo)
        Me.pnlOptSelect.Location = New System.Drawing.Point(46, 167)
        Me.pnlOptSelect.Name = "pnlOptSelect"
        Me.pnlOptSelect.Size = New System.Drawing.Size(414, 288)
        Me.pnlOptSelect.TabIndex = 23
        '
        'lblKOpSeleTitle
        '
        Me.lblKOpSeleTitle.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKOpSeleTitle.Location = New System.Drawing.Point(3, 6)
        Me.lblKOpSeleTitle.Name = "lblKOpSeleTitle"
        Me.lblKOpSeleTitle.Size = New System.Drawing.Size(183, 19)
        Me.lblKOpSeleTitle.TabIndex = 8
        Me.lblKOpSeleTitle.Text = "賃貸革命V7オプション選択"
        Me.lblKOpSeleTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblKOpSeleTitle.UseCompatibleTextRendering = true
        '
        'pnlOptionSelect
        '
        Me.pnlOptionSelect.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectKaikei)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectKy)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectReform)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectSq)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectClaim)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectNk)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectFB)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectSh)
        Me.pnlOptionSelect.Controls.Add(Me.chkOptionSelectSyoki)
        Me.pnlOptionSelect.Location = New System.Drawing.Point(23, 128)
        Me.pnlOptionSelect.Name = "pnlOptionSelect"
        Me.pnlOptionSelect.Size = New System.Drawing.Size(379, 144)
        Me.pnlOptionSelect.TabIndex = 11
        '
        'chkOptionSelectKaikei
        '
        Me.chkOptionSelectKaikei.Checked = true
        Me.chkOptionSelectKaikei.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectKaikei.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectKaikei.Location = New System.Drawing.Point(191, 86)
        Me.chkOptionSelectKaikei.Name = "chkOptionSelectKaikei"
        Me.chkOptionSelectKaikei.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectKaikei.TabIndex = 8
        Me.chkOptionSelectKaikei.Text = "会計連動"
        Me.chkOptionSelectKaikei.UseVisualStyleBackColor = true
        '
        'chkOptionSelectKy
        '
        Me.chkOptionSelectKy.Checked = true
        Me.chkOptionSelectKy.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectKy.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectKy.Location = New System.Drawing.Point(35, 8)
        Me.chkOptionSelectKy.Name = "chkOptionSelectKy"
        Me.chkOptionSelectKy.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectKy.TabIndex = 0
        Me.chkOptionSelectKy.Text = "契約管理"
        Me.chkOptionSelectKy.UseVisualStyleBackColor = true
        '
        'chkOptionSelectReform
        '
        Me.chkOptionSelectReform.Checked = true
        Me.chkOptionSelectReform.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectReform.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectReform.Location = New System.Drawing.Point(191, 60)
        Me.chkOptionSelectReform.Name = "chkOptionSelectReform"
        Me.chkOptionSelectReform.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectReform.TabIndex = 7
        Me.chkOptionSelectReform.Text = "リフォーム"
        Me.chkOptionSelectReform.UseVisualStyleBackColor = true
        '
        'chkOptionSelectSq
        '
        Me.chkOptionSelectSq.Checked = true
        Me.chkOptionSelectSq.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectSq.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectSq.Location = New System.Drawing.Point(35, 34)
        Me.chkOptionSelectSq.Name = "chkOptionSelectSq"
        Me.chkOptionSelectSq.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectSq.TabIndex = 1
        Me.chkOptionSelectSq.Text = "請求管理"
        Me.chkOptionSelectSq.UseVisualStyleBackColor = true
        '
        'chkOptionSelectClaim
        '
        Me.chkOptionSelectClaim.Checked = true
        Me.chkOptionSelectClaim.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectClaim.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectClaim.Location = New System.Drawing.Point(191, 34)
        Me.chkOptionSelectClaim.Name = "chkOptionSelectClaim"
        Me.chkOptionSelectClaim.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectClaim.TabIndex = 6
        Me.chkOptionSelectClaim.Text = "クレーム"
        Me.chkOptionSelectClaim.UseVisualStyleBackColor = true
        '
        'chkOptionSelectNk
        '
        Me.chkOptionSelectNk.Checked = true
        Me.chkOptionSelectNk.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectNk.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectNk.Location = New System.Drawing.Point(35, 60)
        Me.chkOptionSelectNk.Name = "chkOptionSelectNk"
        Me.chkOptionSelectNk.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectNk.TabIndex = 2
        Me.chkOptionSelectNk.Text = "入金管理"
        Me.chkOptionSelectNk.UseVisualStyleBackColor = true
        '
        'chkOptionSelectFB
        '
        Me.chkOptionSelectFB.Checked = true
        Me.chkOptionSelectFB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectFB.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectFB.Location = New System.Drawing.Point(191, 8)
        Me.chkOptionSelectFB.Name = "chkOptionSelectFB"
        Me.chkOptionSelectFB.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectFB.TabIndex = 5
        Me.chkOptionSelectFB.Text = "ファームバンキング"
        Me.chkOptionSelectFB.UseVisualStyleBackColor = true
        '
        'chkOptionSelectSh
        '
        Me.chkOptionSelectSh.Checked = true
        Me.chkOptionSelectSh.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectSh.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectSh.Location = New System.Drawing.Point(35, 86)
        Me.chkOptionSelectSh.Name = "chkOptionSelectSh"
        Me.chkOptionSelectSh.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectSh.TabIndex = 3
        Me.chkOptionSelectSh.Text = "支払管理"
        Me.chkOptionSelectSh.UseVisualStyleBackColor = true
        '
        'chkOptionSelectSyoki
        '
        Me.chkOptionSelectSyoki.Checked = true
        Me.chkOptionSelectSyoki.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOptionSelectSyoki.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOptionSelectSyoki.Location = New System.Drawing.Point(35, 112)
        Me.chkOptionSelectSyoki.Name = "chkOptionSelectSyoki"
        Me.chkOptionSelectSyoki.Size = New System.Drawing.Size(150, 20)
        Me.chkOptionSelectSyoki.TabIndex = 4
        Me.chkOptionSelectSyoki.Text = "初期設定"
        Me.chkOptionSelectSyoki.UseVisualStyleBackColor = true
        '
        'lblHOpSeleTitle
        '
        Me.lblHOpSeleTitle.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHOpSeleTitle.Location = New System.Drawing.Point(3, 6)
        Me.lblHOpSeleTitle.Name = "lblHOpSeleTitle"
        Me.lblHOpSeleTitle.Size = New System.Drawing.Size(183, 19)
        Me.lblHOpSeleTitle.TabIndex = 19
        Me.lblHOpSeleTitle.Text = "賃貸革命10オプション選択"
        Me.lblHOpSeleTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblHOpSeleTitle.UseCompatibleTextRendering = true
        '
        'lblHOpSeleNaiyo
        '
        Me.lblHOpSeleNaiyo.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHOpSeleNaiyo.Location = New System.Drawing.Point(23, 25)
        Me.lblHOpSeleNaiyo.Name = "lblHOpSeleNaiyo"
        Me.lblHOpSeleNaiyo.Size = New System.Drawing.Size(384, 80)
        Me.lblHOpSeleNaiyo.TabIndex = 20
        Me.lblHOpSeleNaiyo.Text = "賃貸革命10の購入オプションにチェックを入れて下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※未購入オプションにチェックを入れても、10でデータを利用することはできません。"
        Me.lblHOpSeleNaiyo.UseCompatibleTextRendering = true
        '
        'lblSOpSeleNaiyo
        '
        Me.lblSOpSeleNaiyo.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblSOpSeleNaiyo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblSOpSeleNaiyo.Location = New System.Drawing.Point(23, 105)
        Me.lblSOpSeleNaiyo.Name = "lblSOpSeleNaiyo"
        Me.lblSOpSeleNaiyo.Size = New System.Drawing.Size(384, 20)
        Me.lblSOpSeleNaiyo.TabIndex = 10
        Me.lblSOpSeleNaiyo.Text = "※不明な場合、サポートへご確認下さい。"
        Me.lblSOpSeleNaiyo.UseCompatibleTextRendering = true
        '
        'lblKOpSeleNaiyo
        '
        Me.lblKOpSeleNaiyo.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKOpSeleNaiyo.Location = New System.Drawing.Point(23, 25)
        Me.lblKOpSeleNaiyo.Name = "lblKOpSeleNaiyo"
        Me.lblKOpSeleNaiyo.Size = New System.Drawing.Size(384, 80)
        Me.lblKOpSeleNaiyo.TabIndex = 9
        Me.lblKOpSeleNaiyo.Text = "賃貸革命V7で使用していた項目(購入オプション)にチェックを入れて下さい。チェックがない場合はコンバート対象外となります。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※任意オプションについて、V7では使"& _ 
    "用、10では未使用の場合、コンバートを行っても、10でデータを利用することはできません。"
        Me.lblKOpSeleNaiyo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblKOpSeleNaiyo.UseCompatibleTextRendering = true
        '
        'pnlUserName
        '
        Me.pnlUserName.Controls.Add(Me.lblExecutor)
        Me.pnlUserName.Controls.Add(Me.txtRecUser)
        Me.pnlUserName.Controls.Add(Me.Label58)
        Me.pnlUserName.Location = New System.Drawing.Point(492, 461)
        Me.pnlUserName.Name = "pnlUserName"
        Me.pnlUserName.Size = New System.Drawing.Size(414, 82)
        Me.pnlUserName.TabIndex = 22
        '
        'lblExecutor
        '
        Me.lblExecutor.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblExecutor.Location = New System.Drawing.Point(3, 6)
        Me.lblExecutor.Name = "lblExecutor"
        Me.lblExecutor.Size = New System.Drawing.Size(128, 19)
        Me.lblExecutor.TabIndex = 16
        Me.lblExecutor.Text = "作業者名"
        Me.lblExecutor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblExecutor.UseCompatibleTextRendering = true
        '
        'txtRecUser
        '
        Me.txtRecUser.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtRecUser.Location = New System.Drawing.Point(23, 46)
        Me.txtRecUser.Name = "txtRecUser"
        Me.txtRecUser.Size = New System.Drawing.Size(384, 25)
        Me.txtRecUser.TabIndex = 18
        '
        'Label58
        '
        Me.Label58.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label58.Location = New System.Drawing.Point(23, 25)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(384, 20)
        Me.Label58.TabIndex = 17
        Me.Label58.Text = "設定値は、コンバートしたデータベースの履歴に登録されます。"
        Me.Label58.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label58.UseCompatibleTextRendering = true
        '
        'pnlLogPath
        '
        Me.pnlLogPath.Controls.Add(Me.lblLogDir)
        Me.pnlLogPath.Controls.Add(Me.Label57)
        Me.pnlLogPath.Controls.Add(Me.btnLogDirSeach)
        Me.pnlLogPath.Controls.Add(Me.txtLogDirPath)
        Me.pnlLogPath.Location = New System.Drawing.Point(46, 461)
        Me.pnlLogPath.Name = "pnlLogPath"
        Me.pnlLogPath.Size = New System.Drawing.Size(414, 82)
        Me.pnlLogPath.TabIndex = 21
        '
        'lblLogDir
        '
        Me.lblLogDir.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblLogDir.Location = New System.Drawing.Point(3, 6)
        Me.lblLogDir.Name = "lblLogDir"
        Me.lblLogDir.Size = New System.Drawing.Size(128, 19)
        Me.lblLogDir.TabIndex = 12
        Me.lblLogDir.Text = "ログファイル格納先"
        Me.lblLogDir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblLogDir.UseCompatibleTextRendering = true
        '
        'Label57
        '
        Me.Label57.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label57.Location = New System.Drawing.Point(23, 25)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(384, 20)
        Me.Label57.TabIndex = 13
        Me.Label57.Text = "コンバート結果を出力するファイルの格納先を設定して下さい。"
        Me.Label57.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label57.UseCompatibleTextRendering = true
        '
        'btnLogDirSeach
        '
        Me.btnLogDirSeach.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnLogDirSeach.Location = New System.Drawing.Point(377, 46)
        Me.btnLogDirSeach.Name = "btnLogDirSeach"
        Me.btnLogDirSeach.Size = New System.Drawing.Size(30, 25)
        Me.btnLogDirSeach.TabIndex = 15
        Me.btnLogDirSeach.Text = "..."
        Me.btnLogDirSeach.UseVisualStyleBackColor = true
        '
        'txtLogDirPath
        '
        Me.txtLogDirPath.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtLogDirPath.Location = New System.Drawing.Point(23, 46)
        Me.txtLogDirPath.Name = "txtLogDirPath"
        Me.txtLogDirPath.Size = New System.Drawing.Size(345, 25)
        Me.txtLogDirPath.TabIndex = 14
        '
        'lblSyokiDescription1
        '
        Me.lblSyokiDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblSyokiDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblSyokiDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblSyokiDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblSyokiDescription1.Name = "lblSyokiDescription1"
        Me.lblSyokiDescription1.Size = New System.Drawing.Size(820, 80)
        Me.lblSyokiDescription1.TabIndex = 0
        Me.lblSyokiDescription1.Text = "コンバーターの初期設定を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"設定が完了したら「次へ」ボタンを押してください。"
        Me.lblSyokiDescription1.UseCompatibleTextRendering = true
        '
        'tabPageMenu
        '
        Me.tabPageMenu.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageMenu.Controls.Add(Me.grpMenuHojyo)
        Me.tabPageMenu.Controls.Add(Me.lblMenuCaution)
        Me.tabPageMenu.Controls.Add(Me.grpMenuJizen)
        Me.tabPageMenu.Controls.Add(Me.grpMenuDatacv)
        Me.tabPageMenu.Controls.Add(Me.grpMenuGazocv)
        Me.tabPageMenu.Controls.Add(Me.grpMenuJigo)
        Me.tabPageMenu.Controls.Add(Me.lblMenuDescription1)
        Me.tabPageMenu.Location = New System.Drawing.Point(4, 27)
        Me.tabPageMenu.Name = "tabPageMenu"
        Me.tabPageMenu.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageMenu.Size = New System.Drawing.Size(912, 549)
        Me.tabPageMenu.TabIndex = 12
        Me.tabPageMenu.Text = "tabPageSelect"
        '
        'grpMenuHojyo
        '
        Me.grpMenuHojyo.Controls.Add(Me.btnMenuHojyo)
        Me.grpMenuHojyo.Controls.Add(Me.lblMenuHojyo)
        Me.grpMenuHojyo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpMenuHojyo.ForeColor = System.Drawing.Color.Navy
        Me.grpMenuHojyo.Location = New System.Drawing.Point(54, 452)
        Me.grpMenuHojyo.Name = "grpMenuHojyo"
        Me.grpMenuHojyo.Size = New System.Drawing.Size(686, 80)
        Me.grpMenuHojyo.TabIndex = 142
        Me.grpMenuHojyo.TabStop = false
        Me.grpMenuHojyo.Text = " ※補助機能 "
        '
        'btnMenuHojyo
        '
        Me.btnMenuHojyo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMenuHojyo.ForeColor = System.Drawing.Color.Black
        Me.btnMenuHojyo.Location = New System.Drawing.Point(31, 30)
        Me.btnMenuHojyo.Name = "btnMenuHojyo"
        Me.btnMenuHojyo.Size = New System.Drawing.Size(101, 33)
        Me.btnMenuHojyo.TabIndex = 19
        Me.btnMenuHojyo.Text = "補助機能"
        Me.btnMenuHojyo.UseVisualStyleBackColor = true
        '
        'lblMenuHojyo
        '
        Me.lblMenuHojyo.BackColor = System.Drawing.SystemColors.Menu
        Me.lblMenuHojyo.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblMenuHojyo.ForeColor = System.Drawing.Color.Black
        Me.lblMenuHojyo.Location = New System.Drawing.Point(148, 19)
        Me.lblMenuHojyo.Name = "lblMenuHojyo"
        Me.lblMenuHojyo.Size = New System.Drawing.Size(513, 55)
        Me.lblMenuHojyo.TabIndex = 18
        Me.lblMenuHojyo.Text = "一部の賃貸革命V7データの確認を行います。"
        Me.lblMenuHojyo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblMenuHojyo.UseCompatibleTextRendering = true
        '
        'lblMenuCaution
        '
        Me.lblMenuCaution.BackColor = System.Drawing.SystemColors.Menu
        Me.lblMenuCaution.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblMenuCaution.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblMenuCaution.Location = New System.Drawing.Point(30, 63)
        Me.lblMenuCaution.Name = "lblMenuCaution"
        Me.lblMenuCaution.Size = New System.Drawing.Size(701, 24)
        Me.lblMenuCaution.TabIndex = 147
        Me.lblMenuCaution.Text = "※先に作業を完了させないと選択できない項目があります。"
        Me.lblMenuCaution.UseCompatibleTextRendering = true
        '
        'grpMenuJizen
        '
        Me.grpMenuJizen.Controls.Add(Me.lblKMenuJizen)
        Me.grpMenuJizen.Controls.Add(Me.lblHMenuJizen)
        Me.grpMenuJizen.Controls.Add(Me.btnMenuJizen)
        Me.grpMenuJizen.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpMenuJizen.ForeColor = System.Drawing.Color.Navy
        Me.grpMenuJizen.Location = New System.Drawing.Point(53, 108)
        Me.grpMenuJizen.Name = "grpMenuJizen"
        Me.grpMenuJizen.Size = New System.Drawing.Size(691, 80)
        Me.grpMenuJizen.TabIndex = 144
        Me.grpMenuJizen.TabStop = false
        Me.grpMenuJizen.Text = " 1. 事前作業 "
        '
        'lblKMenuJizen
        '
        Me.lblKMenuJizen.BackColor = System.Drawing.SystemColors.Menu
        Me.lblKMenuJizen.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblKMenuJizen.ForeColor = System.Drawing.Color.Black
        Me.lblKMenuJizen.Location = New System.Drawing.Point(148, 19)
        Me.lblKMenuJizen.Name = "lblKMenuJizen"
        Me.lblKMenuJizen.Size = New System.Drawing.Size(518, 55)
        Me.lblKMenuJizen.TabIndex = 9
        Me.lblKMenuJizen.Text = "コンバートを実行する前の賃貸革命V7データの調整を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"調整が必要と思われる内容のリストを出力し、それを元に調整を行って下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※コンバートを行う前に"& _ 
    "必ず行って下さい。"
        Me.lblKMenuJizen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblKMenuJizen.UseCompatibleTextRendering = true
        '
        'lblHMenuJizen
        '
        Me.lblHMenuJizen.BackColor = System.Drawing.SystemColors.Menu
        Me.lblHMenuJizen.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblHMenuJizen.ForeColor = System.Drawing.Color.Black
        Me.lblHMenuJizen.Location = New System.Drawing.Point(148, 19)
        Me.lblHMenuJizen.Name = "lblHMenuJizen"
        Me.lblHMenuJizen.Size = New System.Drawing.Size(518, 55)
        Me.lblHMenuJizen.TabIndex = 11
        Me.lblHMenuJizen.Text = "中間ファイルの作成を行います。"
        Me.lblHMenuJizen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblHMenuJizen.UseCompatibleTextRendering = true
        '
        'btnMenuJizen
        '
        Me.btnMenuJizen.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMenuJizen.ForeColor = System.Drawing.Color.Black
        Me.btnMenuJizen.Location = New System.Drawing.Point(32, 30)
        Me.btnMenuJizen.Name = "btnMenuJizen"
        Me.btnMenuJizen.Size = New System.Drawing.Size(101, 33)
        Me.btnMenuJizen.TabIndex = 10
        Me.btnMenuJizen.Text = "事前調整"
        Me.btnMenuJizen.UseVisualStyleBackColor = true
        '
        'grpMenuDatacv
        '
        Me.grpMenuDatacv.Controls.Add(Me.lblKMenuDatacv)
        Me.grpMenuDatacv.Controls.Add(Me.lblHMenuDatacv)
        Me.grpMenuDatacv.Controls.Add(Me.btnMenuDatacv)
        Me.grpMenuDatacv.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpMenuDatacv.ForeColor = System.Drawing.Color.Navy
        Me.grpMenuDatacv.Location = New System.Drawing.Point(53, 194)
        Me.grpMenuDatacv.Name = "grpMenuDatacv"
        Me.grpMenuDatacv.Size = New System.Drawing.Size(686, 80)
        Me.grpMenuDatacv.TabIndex = 143
        Me.grpMenuDatacv.TabStop = false
        Me.grpMenuDatacv.Text = " 2. データコンバート "
        '
        'lblKMenuDatacv
        '
        Me.lblKMenuDatacv.BackColor = System.Drawing.SystemColors.Menu
        Me.lblKMenuDatacv.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblKMenuDatacv.ForeColor = System.Drawing.Color.Black
        Me.lblKMenuDatacv.Location = New System.Drawing.Point(148, 19)
        Me.lblKMenuDatacv.Name = "lblKMenuDatacv"
        Me.lblKMenuDatacv.Size = New System.Drawing.Size(513, 55)
        Me.lblKMenuDatacv.TabIndex = 12
        Me.lblKMenuDatacv.Text = "データコンバートを行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※""事前調整""を先に完了させておく必要があります。"
        Me.lblKMenuDatacv.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblKMenuDatacv.UseCompatibleTextRendering = true
        '
        'lblHMenuDatacv
        '
        Me.lblHMenuDatacv.BackColor = System.Drawing.SystemColors.Menu
        Me.lblHMenuDatacv.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblHMenuDatacv.ForeColor = System.Drawing.Color.Black
        Me.lblHMenuDatacv.Location = New System.Drawing.Point(148, 19)
        Me.lblHMenuDatacv.Name = "lblHMenuDatacv"
        Me.lblHMenuDatacv.Size = New System.Drawing.Size(513, 55)
        Me.lblHMenuDatacv.TabIndex = 14
        Me.lblHMenuDatacv.Text = "中間ファイルのチェックとデータコンバートを行います。"
        Me.lblHMenuDatacv.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblHMenuDatacv.UseCompatibleTextRendering = true
        '
        'btnMenuDatacv
        '
        Me.btnMenuDatacv.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMenuDatacv.ForeColor = System.Drawing.Color.Black
        Me.btnMenuDatacv.Location = New System.Drawing.Point(32, 30)
        Me.btnMenuDatacv.Name = "btnMenuDatacv"
        Me.btnMenuDatacv.Size = New System.Drawing.Size(101, 33)
        Me.btnMenuDatacv.TabIndex = 13
        Me.btnMenuDatacv.Text = "移行実行"
        Me.btnMenuDatacv.UseVisualStyleBackColor = true
        '
        'grpMenuGazocv
        '
        Me.grpMenuGazocv.Controls.Add(Me.btnMenuGazocv)
        Me.grpMenuGazocv.Controls.Add(Me.lblMenuGazocv)
        Me.grpMenuGazocv.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpMenuGazocv.ForeColor = System.Drawing.Color.Navy
        Me.grpMenuGazocv.Location = New System.Drawing.Point(53, 280)
        Me.grpMenuGazocv.Name = "grpMenuGazocv"
        Me.grpMenuGazocv.Size = New System.Drawing.Size(686, 80)
        Me.grpMenuGazocv.TabIndex = 142
        Me.grpMenuGazocv.TabStop = false
        Me.grpMenuGazocv.Text = " 3.画像コンバート "
        '
        'btnMenuGazocv
        '
        Me.btnMenuGazocv.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMenuGazocv.ForeColor = System.Drawing.Color.Black
        Me.btnMenuGazocv.Location = New System.Drawing.Point(32, 30)
        Me.btnMenuGazocv.Name = "btnMenuGazocv"
        Me.btnMenuGazocv.Size = New System.Drawing.Size(101, 33)
        Me.btnMenuGazocv.TabIndex = 16
        Me.btnMenuGazocv.Text = "移行実行"
        Me.btnMenuGazocv.UseVisualStyleBackColor = true
        '
        'lblMenuGazocv
        '
        Me.lblMenuGazocv.BackColor = System.Drawing.SystemColors.Menu
        Me.lblMenuGazocv.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblMenuGazocv.ForeColor = System.Drawing.Color.Black
        Me.lblMenuGazocv.Location = New System.Drawing.Point(148, 19)
        Me.lblMenuGazocv.Name = "lblMenuGazocv"
        Me.lblMenuGazocv.Size = New System.Drawing.Size(513, 55)
        Me.lblMenuGazocv.TabIndex = 15
        Me.lblMenuGazocv.Text = "画像コンバートを行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※""データコンバート""を先に完了させておく必要があります。"
        Me.lblMenuGazocv.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblMenuGazocv.UseCompatibleTextRendering = true
        '
        'grpMenuJigo
        '
        Me.grpMenuJigo.Controls.Add(Me.btnMenuJigo)
        Me.grpMenuJigo.Controls.Add(Me.lblMenuJigo)
        Me.grpMenuJigo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpMenuJigo.ForeColor = System.Drawing.Color.Navy
        Me.grpMenuJigo.Location = New System.Drawing.Point(54, 366)
        Me.grpMenuJigo.Name = "grpMenuJigo"
        Me.grpMenuJigo.Size = New System.Drawing.Size(686, 80)
        Me.grpMenuJigo.TabIndex = 141
        Me.grpMenuJigo.TabStop = false
        Me.grpMenuJigo.Text = " 4. 事後作業 "
        '
        'btnMenuJigo
        '
        Me.btnMenuJigo.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMenuJigo.ForeColor = System.Drawing.Color.Black
        Me.btnMenuJigo.Location = New System.Drawing.Point(31, 30)
        Me.btnMenuJigo.Name = "btnMenuJigo"
        Me.btnMenuJigo.Size = New System.Drawing.Size(101, 33)
        Me.btnMenuJigo.TabIndex = 19
        Me.btnMenuJigo.Text = "事後調整"
        Me.btnMenuJigo.UseVisualStyleBackColor = true
        '
        'lblMenuJigo
        '
        Me.lblMenuJigo.BackColor = System.Drawing.SystemColors.Menu
        Me.lblMenuJigo.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblMenuJigo.ForeColor = System.Drawing.Color.Black
        Me.lblMenuJigo.Location = New System.Drawing.Point(148, 19)
        Me.lblMenuJigo.Name = "lblMenuJigo"
        Me.lblMenuJigo.Size = New System.Drawing.Size(513, 55)
        Me.lblMenuJigo.TabIndex = 18
        Me.lblMenuJigo.Text = "コンバートを実行した後の賃貸革命10データの調整を行います。"
        Me.lblMenuJigo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblMenuJigo.UseCompatibleTextRendering = true
        '
        'lblMenuDescription1
        '
        Me.lblMenuDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblMenuDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblMenuDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblMenuDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblMenuDescription1.Name = "lblMenuDescription1"
        Me.lblMenuDescription1.Size = New System.Drawing.Size(820, 35)
        Me.lblMenuDescription1.TabIndex = 140
        Me.lblMenuDescription1.Text = "以下の内容を確認し、必要な作業を行って下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(ボタンを押して下さい)"
        Me.lblMenuDescription1.UseCompatibleTextRendering = true
        '
        'tabPageJizen
        '
        Me.tabPageJizen.BackColor = System.Drawing.SystemColors.MenuBar
        Me.tabPageJizen.Controls.Add(Me.pnlJizenListPath)
        Me.tabPageJizen.Controls.Add(Me.lblLine1)
        Me.tabPageJizen.Controls.Add(Me.lblHidden3)
        Me.tabPageJizen.Controls.Add(Me.tabCtrlJizen)
        Me.tabPageJizen.Controls.Add(Me.lblCautionDescription)
        Me.tabPageJizen.Controls.Add(Me.lblJizenDescription1)
        Me.tabPageJizen.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJizen.Name = "tabPageJizen"
        Me.tabPageJizen.Size = New System.Drawing.Size(912, 549)
        Me.tabPageJizen.TabIndex = 7
        Me.tabPageJizen.Text = " 事前調整作業"
        '
        'pnlJizenListPath
        '
        Me.pnlJizenListPath.Controls.Add(Me.Label254)
        Me.pnlJizenListPath.Controls.Add(Me.btnJizenListDirSeach)
        Me.pnlJizenListPath.Controls.Add(Me.txtJizenListPath)
        Me.pnlJizenListPath.Location = New System.Drawing.Point(31, 86)
        Me.pnlJizenListPath.Name = "pnlJizenListPath"
        Me.pnlJizenListPath.Size = New System.Drawing.Size(420, 60)
        Me.pnlJizenListPath.TabIndex = 47
        '
        'Label254
        '
        Me.Label254.AutoSize = true
        Me.Label254.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label254.Location = New System.Drawing.Point(5, 5)
        Me.Label254.Name = "Label254"
        Me.Label254.Size = New System.Drawing.Size(116, 18)
        Me.Label254.TabIndex = 3
        Me.Label254.Text = "出力ファイル格納先"
        '
        'btnJizenListDirSeach
        '
        Me.btnJizenListDirSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnJizenListDirSeach.Location = New System.Drawing.Point(383, 27)
        Me.btnJizenListDirSeach.Name = "btnJizenListDirSeach"
        Me.btnJizenListDirSeach.Size = New System.Drawing.Size(30, 25)
        Me.btnJizenListDirSeach.TabIndex = 2
        Me.btnJizenListDirSeach.Text = "..."
        Me.btnJizenListDirSeach.UseVisualStyleBackColor = true
        '
        'txtJizenListPath
        '
        Me.txtJizenListPath.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtJizenListPath.Location = New System.Drawing.Point(23, 28)
        Me.txtJizenListPath.Name = "txtJizenListPath"
        Me.txtJizenListPath.Size = New System.Drawing.Size(354, 24)
        Me.txtJizenListPath.TabIndex = 1
        '
        'lblLine1
        '
        Me.lblLine1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLine1.Location = New System.Drawing.Point(31, 175)
        Me.lblLine1.Name = "lblLine1"
        Me.lblLine1.Size = New System.Drawing.Size(853, 2)
        Me.lblLine1.TabIndex = 3
        '
        'lblHidden3
        '
        Me.lblHidden3.Location = New System.Drawing.Point(6, 120)
        Me.lblHidden3.Name = "lblHidden3"
        Me.lblHidden3.Size = New System.Drawing.Size(18, 55)
        Me.lblHidden3.TabIndex = 2
        Me.lblHidden3.Text = "　"
        '
        'tabCtrlJizen
        '
        Me.tabCtrlJizen.Controls.Add(Me.tabPageJizen1)
        Me.tabCtrlJizen.Controls.Add(Me.tabPageJizen2)
        Me.tabCtrlJizen.Controls.Add(Me.tabPageJizen3)
        Me.tabCtrlJizen.Controls.Add(Me.tabPageJizen4)
        Me.tabCtrlJizen.Controls.Add(Me.tabPageJizen5)
        Me.tabCtrlJizen.Controls.Add(Me.tabPageHJizen1)
        Me.tabCtrlJizen.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.tabCtrlJizen.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.tabCtrlJizen.Location = New System.Drawing.Point(30, 150)
        Me.tabCtrlJizen.Name = "tabCtrlJizen"
        Me.tabCtrlJizen.SelectedIndex = 0
        Me.tabCtrlJizen.Size = New System.Drawing.Size(855, 394)
        Me.tabCtrlJizen.TabIndex = 4
        Me.tabCtrlJizen.TabStop = false
        '
        'tabPageJizen1
        '
        Me.tabPageJizen1.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJizen1.Controls.Add(Me.grpJizen1)
        Me.tabPageJizen1.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJizen1.Name = "tabPageJizen1"
        Me.tabPageJizen1.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageJizen1.Size = New System.Drawing.Size(847, 363)
        Me.tabPageJizen1.TabIndex = 1
        Me.tabPageJizen1.Text = " 事前作業1"
        '
        'grpJizen1
        '
        Me.grpJizen1.BackColor = System.Drawing.SystemColors.Menu
        Me.grpJizen1.Controls.Add(Me.pnlJizenHurikae)
        Me.grpJizen1.Controls.Add(Me.pnlJizenAzu)
        Me.grpJizen1.Controls.Add(Me.pnlJizenKai)
        Me.grpJizen1.Controls.Add(Me.lblJizenPageCnt1)
        Me.grpJizen1.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpJizen1.Location = New System.Drawing.Point(6, 15)
        Me.grpJizen1.Name = "grpJizen1"
        Me.grpJizen1.Size = New System.Drawing.Size(835, 342)
        Me.grpJizen1.TabIndex = 0
        Me.grpJizen1.TabStop = false
        Me.grpJizen1.Text = "【ユーザによる事前確認・調整】"
        '
        'pnlJizenHurikae
        '
        Me.pnlJizenHurikae.Controls.Add(Me.chkJizenHurikae)
        Me.pnlJizenHurikae.Controls.Add(Me.Label88)
        Me.pnlJizenHurikae.Location = New System.Drawing.Point(424, 18)
        Me.pnlJizenHurikae.Name = "pnlJizenHurikae"
        Me.pnlJizenHurikae.Size = New System.Drawing.Size(400, 80)
        Me.pnlJizenHurikae.TabIndex = 6
        '
        'chkJizenHurikae
        '
        Me.chkJizenHurikae.AutoSize = true
        Me.chkJizenHurikae.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenHurikae.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenHurikae.Name = "chkJizenHurikae"
        Me.chkJizenHurikae.Size = New System.Drawing.Size(123, 22)
        Me.chkJizenHurikae.TabIndex = 6
        Me.chkJizenHurikae.Text = "口座振替入金処理"
        Me.chkJizenHurikae.UseVisualStyleBackColor = true
        '
        'Label88
        '
        Me.Label88.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label88.Location = New System.Drawing.Point(36, 29)
        Me.Label88.Name = "Label88"
        Me.Label88.Size = New System.Drawing.Size(340, 50)
        Me.Label88.TabIndex = 7
        Me.Label88.Text = "賃貸革命V7にて行った口座振替請求に対する振替入金について、事前に処理しておく必要があります。賃貸革命V7で必要な口座振替入金処理を行って下さい。"
        Me.Label88.UseCompatibleTextRendering = true
        '
        'pnlJizenAzu
        '
        Me.pnlJizenAzu.Controls.Add(Me.chkJizenAzu)
        Me.pnlJizenAzu.Controls.Add(Me.Label85)
        Me.pnlJizenAzu.Controls.Add(Me.Panel7)
        Me.pnlJizenAzu.Location = New System.Drawing.Point(6, 184)
        Me.pnlJizenAzu.Name = "pnlJizenAzu"
        Me.pnlJizenAzu.Size = New System.Drawing.Size(400, 150)
        Me.pnlJizenAzu.TabIndex = 11
        '
        'chkJizenAzu
        '
        Me.chkJizenAzu.AutoSize = true
        Me.chkJizenAzu.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenAzu.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenAzu.Name = "chkJizenAzu"
        Me.chkJizenAzu.Size = New System.Drawing.Size(123, 22)
        Me.chkJizenAzu.TabIndex = 3
        Me.chkJizenAzu.Text = "預り金データ処理"
        Me.chkJizenAzu.UseVisualStyleBackColor = true
        '
        'Label85
        '
        Me.Label85.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label85.Location = New System.Drawing.Point(36, 29)
        Me.Label85.Name = "Label85"
        Me.Label85.Size = New System.Drawing.Size(340, 50)
        Me.Label85.TabIndex = 4
        Me.Label85.Text = "「削除漏れや放置」などの預り金が存在する場合、削除などの処理を行っておく必要があります。賃貸革命V7または10で必要な預り金の処理を行って下さい。"
        Me.Label85.UseCompatibleTextRendering = true
        '
        'Panel7
        '
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel7.Controls.Add(Me.btnListJizenAzu)
        Me.Panel7.Controls.Add(Me.Label166)
        Me.Panel7.Controls.Add(Me.Label167)
        Me.Panel7.Controls.Add(Me.Label168)
        Me.Panel7.Controls.Add(Me.lblCntJizenAzu)
        Me.Panel7.Location = New System.Drawing.Point(36, 82)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(340, 60)
        Me.Panel7.TabIndex = 5
        '
        'btnListJizenAzu
        '
        Me.btnListJizenAzu.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenAzu.Image = CType(resources.GetObject("btnListJizenAzu.Image"),System.Drawing.Image)
        Me.btnListJizenAzu.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenAzu.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenAzu.Name = "btnListJizenAzu"
        Me.btnListJizenAzu.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenAzu.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenAzu.TabIndex = 1
        Me.btnListJizenAzu.Text = "リスト出力"
        Me.btnListJizenAzu.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenAzu.UseVisualStyleBackColor = true
        '
        'Label166
        '
        Me.Label166.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label166.Location = New System.Drawing.Point(8, 8)
        Me.Label166.Name = "Label166"
        Me.Label166.Size = New System.Drawing.Size(321, 16)
        Me.Label166.TabIndex = 0
        Me.Label166.Text = "V7の契約者預り金のリスト出力"
        Me.Label166.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label166.UseCompatibleTextRendering = true
        '
        'Label167
        '
        Me.Label167.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label167.Location = New System.Drawing.Point(307, 31)
        Me.Label167.Name = "Label167"
        Me.Label167.Size = New System.Drawing.Size(22, 16)
        Me.Label167.TabIndex = 4
        Me.Label167.Text = " )"
        Me.Label167.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label168
        '
        Me.Label168.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label168.Location = New System.Drawing.Point(131, 31)
        Me.Label168.Name = "Label168"
        Me.Label168.Size = New System.Drawing.Size(92, 16)
        Me.Label168.TabIndex = 2
        Me.Label168.Text = "( 抽出件数 = "
        Me.Label168.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenAzu
        '
        Me.lblCntJizenAzu.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenAzu.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenAzu.Name = "lblCntJizenAzu"
        Me.lblCntJizenAzu.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenAzu.TabIndex = 3
        Me.lblCntJizenAzu.Text = "9,999,999"
        Me.lblCntJizenAzu.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlJizenKai
        '
        Me.pnlJizenKai.Controls.Add(Me.chkJizenKai)
        Me.pnlJizenKai.Controls.Add(Me.Label84)
        Me.pnlJizenKai.Controls.Add(Me.Panel6)
        Me.pnlJizenKai.Location = New System.Drawing.Point(6, 18)
        Me.pnlJizenKai.Name = "pnlJizenKai"
        Me.pnlJizenKai.Size = New System.Drawing.Size(400, 166)
        Me.pnlJizenKai.TabIndex = 1
        '
        'chkJizenKai
        '
        Me.chkJizenKai.AutoSize = true
        Me.chkJizenKai.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenKai.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenKai.Name = "chkJizenKai"
        Me.chkJizenKai.Size = New System.Drawing.Size(147, 22)
        Me.chkJizenKai.TabIndex = 0
        Me.chkJizenKai.Text = "解約未処理データ処理"
        Me.chkJizenKai.UseVisualStyleBackColor = true
        '
        'Label84
        '
        Me.Label84.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label84.Location = New System.Drawing.Point(36, 29)
        Me.Label84.Name = "Label84"
        Me.Label84.Size = New System.Drawing.Size(340, 51)
        Me.Label84.TabIndex = 1
        Me.Label84.Text = "「単純な消し忘れ」などの解約未処理データについて、事前に"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"処理しておく必要があります。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"賃貸革命V7で必要な解約処理を完了させて下さい。"
        Me.Label84.UseCompatibleTextRendering = true
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Controls.Add(Me.btnListJizenKai)
        Me.Panel6.Controls.Add(Me.Label161)
        Me.Panel6.Controls.Add(Me.Label164)
        Me.Panel6.Controls.Add(Me.Label185)
        Me.Panel6.Controls.Add(Me.lblCntJizenKai)
        Me.Panel6.Location = New System.Drawing.Point(36, 82)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(340, 75)
        Me.Panel6.TabIndex = 2
        '
        'btnListJizenKai
        '
        Me.btnListJizenKai.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenKai.Image = CType(resources.GetObject("btnListJizenKai.Image"),System.Drawing.Image)
        Me.btnListJizenKai.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenKai.Location = New System.Drawing.Point(8, 44)
        Me.btnListJizenKai.Name = "btnListJizenKai"
        Me.btnListJizenKai.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenKai.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenKai.TabIndex = 1
        Me.btnListJizenKai.Text = "リスト出力"
        Me.btnListJizenKai.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenKai.UseVisualStyleBackColor = true
        '
        'Label161
        '
        Me.Label161.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label161.Location = New System.Drawing.Point(8, 8)
        Me.Label161.Name = "Label161"
        Me.Label161.Size = New System.Drawing.Size(321, 33)
        Me.Label161.TabIndex = 0
        Me.Label161.Text = "解約未処理データ(契約期間が終了かつ請求データが未作成の解約データ)リスト出力"
        Me.Label161.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label161.UseCompatibleTextRendering = true
        '
        'Label164
        '
        Me.Label164.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label164.Location = New System.Drawing.Point(307, 48)
        Me.Label164.Name = "Label164"
        Me.Label164.Size = New System.Drawing.Size(22, 16)
        Me.Label164.TabIndex = 4
        Me.Label164.Text = " )"
        Me.Label164.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label185
        '
        Me.Label185.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label185.Location = New System.Drawing.Point(131, 48)
        Me.Label185.Name = "Label185"
        Me.Label185.Size = New System.Drawing.Size(92, 16)
        Me.Label185.TabIndex = 2
        Me.Label185.Text = "( 抽出件数 = "
        Me.Label185.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenKai
        '
        Me.lblCntJizenKai.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenKai.Location = New System.Drawing.Point(218, 48)
        Me.lblCntJizenKai.Name = "lblCntJizenKai"
        Me.lblCntJizenKai.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenKai.TabIndex = 3
        Me.lblCntJizenKai.Text = "9,999,999"
        Me.lblCntJizenKai.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblJizenPageCnt1
        '
        Me.lblJizenPageCnt1.AutoSize = true
        Me.lblJizenPageCnt1.Location = New System.Drawing.Point(794, 321)
        Me.lblJizenPageCnt1.Name = "lblJizenPageCnt1"
        Me.lblJizenPageCnt1.Size = New System.Drawing.Size(35, 18)
        Me.lblJizenPageCnt1.TabIndex = 10
        Me.lblJizenPageCnt1.Text = "1 / 5"
        '
        'tabPageJizen2
        '
        Me.tabPageJizen2.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJizen2.Controls.Add(Me.grpJizen2)
        Me.tabPageJizen2.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJizen2.Name = "tabPageJizen2"
        Me.tabPageJizen2.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageJizen2.Size = New System.Drawing.Size(847, 363)
        Me.tabPageJizen2.TabIndex = 0
        Me.tabPageJizen2.Text = " 事前作業2"
        '
        'grpJizen2
        '
        Me.grpJizen2.BackColor = System.Drawing.SystemColors.Menu
        Me.grpJizen2.Controls.Add(Me.Panel12)
        Me.grpJizen2.Controls.Add(Me.pnlJizenHasseiOw)
        Me.grpJizen2.Controls.Add(Me.pnlJizenYanuso)
        Me.grpJizen2.Controls.Add(Me.pnlJizenBkhourei)
        Me.grpJizen2.Controls.Add(Me.lblJizenPageCnt2)
        Me.grpJizen2.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpJizen2.Location = New System.Drawing.Point(6, 15)
        Me.grpJizen2.Name = "grpJizen2"
        Me.grpJizen2.Size = New System.Drawing.Size(835, 342)
        Me.grpJizen2.TabIndex = 0
        Me.grpJizen2.TabStop = false
        Me.grpJizen2.Text = "【賃貸革命10仕様によるユーザ確認・調整】"
        '
        'Panel12
        '
        Me.Panel12.Controls.Add(Me.Panel5)
        Me.Panel12.Controls.Add(Me.chkJizenHasseiHen)
        Me.Panel12.Controls.Add(Me.Label165)
        Me.Panel12.Location = New System.Drawing.Point(429, 174)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(400, 120)
        Me.Panel12.TabIndex = 9
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Controls.Add(Me.btnListJizenHasseiHen)
        Me.Panel5.Controls.Add(Me.Label125)
        Me.Panel5.Controls.Add(Me.Label130)
        Me.Panel5.Controls.Add(Me.Label131)
        Me.Panel5.Controls.Add(Me.lblCntJizenHasseiHen)
        Me.Panel5.Location = New System.Drawing.Point(36, 52)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(340, 60)
        Me.Panel5.TabIndex = 11
        '
        'btnListJizenHasseiHen
        '
        Me.btnListJizenHasseiHen.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenHasseiHen.Image = CType(resources.GetObject("btnListJizenHasseiHen.Image"),System.Drawing.Image)
        Me.btnListJizenHasseiHen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenHasseiHen.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenHasseiHen.Name = "btnListJizenHasseiHen"
        Me.btnListJizenHasseiHen.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenHasseiHen.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenHasseiHen.TabIndex = 1
        Me.btnListJizenHasseiHen.Text = "リスト出力"
        Me.btnListJizenHasseiHen.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenHasseiHen.UseVisualStyleBackColor = true
        '
        'Label125
        '
        Me.Label125.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label125.Location = New System.Drawing.Point(8, 8)
        Me.Label125.Name = "Label125"
        Me.Label125.Size = New System.Drawing.Size(321, 16)
        Me.Label125.TabIndex = 0
        Me.Label125.Text = "「1,2,3,4,6,12ヶ月以外」の発生月設定データのリスト出力"
        Me.Label125.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label125.UseCompatibleTextRendering = true
        '
        'Label130
        '
        Me.Label130.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label130.Location = New System.Drawing.Point(307, 31)
        Me.Label130.Name = "Label130"
        Me.Label130.Size = New System.Drawing.Size(22, 16)
        Me.Label130.TabIndex = 4
        Me.Label130.Text = " )"
        Me.Label130.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label131
        '
        Me.Label131.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label131.Location = New System.Drawing.Point(131, 31)
        Me.Label131.Name = "Label131"
        Me.Label131.Size = New System.Drawing.Size(92, 16)
        Me.Label131.TabIndex = 2
        Me.Label131.Text = "( 抽出件数 = "
        Me.Label131.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenHasseiHen
        '
        Me.lblCntJizenHasseiHen.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenHasseiHen.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenHasseiHen.Name = "lblCntJizenHasseiHen"
        Me.lblCntJizenHasseiHen.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenHasseiHen.TabIndex = 3
        Me.lblCntJizenHasseiHen.Text = "9,999,999"
        Me.lblCntJizenHasseiHen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkJizenHasseiHen
        '
        Me.chkJizenHasseiHen.AutoSize = true
        Me.chkJizenHasseiHen.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenHasseiHen.Location = New System.Drawing.Point(7, 6)
        Me.chkJizenHasseiHen.Name = "chkJizenHasseiHen"
        Me.chkJizenHasseiHen.Size = New System.Drawing.Size(159, 22)
        Me.chkJizenHasseiHen.TabIndex = 9
        Me.chkJizenHasseiHen.Text = "随時変動費の発生月調整"
        Me.chkJizenHasseiHen.UseVisualStyleBackColor = true
        '
        'Label165
        '
        Me.Label165.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label165.Location = New System.Drawing.Point(36, 31)
        Me.Label165.Name = "Label165"
        Me.Label165.Size = New System.Drawing.Size(340, 18)
        Me.Label165.TabIndex = 10
        Me.Label165.Text = "※""家主固定控除の発生月調整""と同様"
        Me.Label165.UseCompatibleTextRendering = true
        '
        'pnlJizenHasseiOw
        '
        Me.pnlJizenHasseiOw.Controls.Add(Me.chkJizenHasseiOw)
        Me.pnlJizenHasseiOw.Controls.Add(Me.Label100)
        Me.pnlJizenHasseiOw.Controls.Add(Me.Panel3)
        Me.pnlJizenHasseiOw.Location = New System.Drawing.Point(429, 18)
        Me.pnlJizenHasseiOw.Name = "pnlJizenHasseiOw"
        Me.pnlJizenHasseiOw.Size = New System.Drawing.Size(400, 150)
        Me.pnlJizenHasseiOw.TabIndex = 13
        '
        'chkJizenHasseiOw
        '
        Me.chkJizenHasseiOw.AutoSize = true
        Me.chkJizenHasseiOw.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenHasseiOw.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenHasseiOw.Name = "chkJizenHasseiOw"
        Me.chkJizenHasseiOw.Size = New System.Drawing.Size(171, 22)
        Me.chkJizenHasseiOw.TabIndex = 6
        Me.chkJizenHasseiOw.Text = "家主固定控除の発生月調整"
        Me.chkJizenHasseiOw.UseVisualStyleBackColor = true
        '
        'Label100
        '
        Me.Label100.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label100.Location = New System.Drawing.Point(36, 29)
        Me.Label100.Name = "Label100"
        Me.Label100.Size = New System.Drawing.Size(340, 50)
        Me.Label100.TabIndex = 7
        Me.Label100.Text = "発生月が「1,2,3,4,6,12ヶ月以外」の場合は移行することができません。賃貸革命V7または10で移行されなかったデータの調整(その他請求・家主控除請求へ登録"& _ 
    "するなど)を行って下さい。"
        Me.Label100.UseCompatibleTextRendering = true
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.btnListJizenHasseiOw)
        Me.Panel3.Controls.Add(Me.Label89)
        Me.Panel3.Controls.Add(Me.Label118)
        Me.Panel3.Controls.Add(Me.Label120)
        Me.Panel3.Controls.Add(Me.lblCntJizenHasseiOw)
        Me.Panel3.Location = New System.Drawing.Point(36, 82)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(340, 60)
        Me.Panel3.TabIndex = 8
        '
        'btnListJizenHasseiOw
        '
        Me.btnListJizenHasseiOw.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenHasseiOw.Image = CType(resources.GetObject("btnListJizenHasseiOw.Image"),System.Drawing.Image)
        Me.btnListJizenHasseiOw.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenHasseiOw.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenHasseiOw.Name = "btnListJizenHasseiOw"
        Me.btnListJizenHasseiOw.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenHasseiOw.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenHasseiOw.TabIndex = 1
        Me.btnListJizenHasseiOw.Text = "リスト出力"
        Me.btnListJizenHasseiOw.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenHasseiOw.UseVisualStyleBackColor = true
        '
        'Label89
        '
        Me.Label89.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label89.Location = New System.Drawing.Point(8, 8)
        Me.Label89.Name = "Label89"
        Me.Label89.Size = New System.Drawing.Size(321, 16)
        Me.Label89.TabIndex = 0
        Me.Label89.Text = "「1,2,3,4,6,12ヶ月以外」の発生月設定データのリスト出力"
        Me.Label89.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label89.UseCompatibleTextRendering = true
        '
        'Label118
        '
        Me.Label118.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label118.Location = New System.Drawing.Point(307, 31)
        Me.Label118.Name = "Label118"
        Me.Label118.Size = New System.Drawing.Size(22, 16)
        Me.Label118.TabIndex = 4
        Me.Label118.Text = " )"
        Me.Label118.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label120
        '
        Me.Label120.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label120.Location = New System.Drawing.Point(131, 31)
        Me.Label120.Name = "Label120"
        Me.Label120.Size = New System.Drawing.Size(92, 16)
        Me.Label120.TabIndex = 2
        Me.Label120.Text = "( 抽出件数 = "
        Me.Label120.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenHasseiOw
        '
        Me.lblCntJizenHasseiOw.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenHasseiOw.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenHasseiOw.Name = "lblCntJizenHasseiOw"
        Me.lblCntJizenHasseiOw.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenHasseiOw.TabIndex = 3
        Me.lblCntJizenHasseiOw.Text = "9,999,999"
        Me.lblCntJizenHasseiOw.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlJizenYanuso
        '
        Me.pnlJizenYanuso.Controls.Add(Me.chkJizenYanuso)
        Me.pnlJizenYanuso.Controls.Add(Me.Label96)
        Me.pnlJizenYanuso.Controls.Add(Me.Panel8)
        Me.pnlJizenYanuso.Location = New System.Drawing.Point(6, 18)
        Me.pnlJizenYanuso.Name = "pnlJizenYanuso"
        Me.pnlJizenYanuso.Size = New System.Drawing.Size(400, 150)
        Me.pnlJizenYanuso.TabIndex = 0
        '
        'chkJizenYanuso
        '
        Me.chkJizenYanuso.AutoSize = true
        Me.chkJizenYanuso.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenYanuso.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenYanuso.Name = "chkJizenYanuso"
        Me.chkJizenYanuso.Size = New System.Drawing.Size(231, 22)
        Me.chkJizenYanuso.TabIndex = 0
        Me.chkJizenYanuso.Text = "家主送金先未設定データの確認・調整"
        Me.chkJizenYanuso.UseVisualStyleBackColor = true
        '
        'Label96
        '
        Me.Label96.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label96.Location = New System.Drawing.Point(36, 29)
        Me.Label96.Name = "Label96"
        Me.Label96.Size = New System.Drawing.Size(340, 50)
        Me.Label96.TabIndex = 1
        Me.Label96.Text = "家主送金先は必須項目の為、必ず設定する必要があります(V7の[物件管理]-[送金先]タブ、10の送金ルール)。未登録データが存在する場合、V7で未登録データの調整"& _ 
    "を行って下さい。"
        Me.Label96.UseCompatibleTextRendering = true
        '
        'Panel8
        '
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel8.Controls.Add(Me.btnListJizenYanuso)
        Me.Panel8.Controls.Add(Me.Label188)
        Me.Panel8.Controls.Add(Me.Label189)
        Me.Panel8.Controls.Add(Me.Label190)
        Me.Panel8.Controls.Add(Me.lblCntJizenYanuso)
        Me.Panel8.Location = New System.Drawing.Point(36, 82)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(340, 60)
        Me.Panel8.TabIndex = 2
        '
        'btnListJizenYanuso
        '
        Me.btnListJizenYanuso.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenYanuso.Image = CType(resources.GetObject("btnListJizenYanuso.Image"),System.Drawing.Image)
        Me.btnListJizenYanuso.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenYanuso.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenYanuso.Name = "btnListJizenYanuso"
        Me.btnListJizenYanuso.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenYanuso.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenYanuso.TabIndex = 1
        Me.btnListJizenYanuso.Text = "リスト出力"
        Me.btnListJizenYanuso.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenYanuso.UseVisualStyleBackColor = true
        '
        'Label188
        '
        Me.Label188.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label188.Location = New System.Drawing.Point(8, 8)
        Me.Label188.Name = "Label188"
        Me.Label188.Size = New System.Drawing.Size(321, 16)
        Me.Label188.TabIndex = 0
        Me.Label188.Text = "家主送金先未設定データのリスト出力"
        Me.Label188.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label188.UseCompatibleTextRendering = true
        '
        'Label189
        '
        Me.Label189.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label189.Location = New System.Drawing.Point(307, 31)
        Me.Label189.Name = "Label189"
        Me.Label189.Size = New System.Drawing.Size(22, 16)
        Me.Label189.TabIndex = 4
        Me.Label189.Text = " )"
        Me.Label189.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label190
        '
        Me.Label190.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label190.Location = New System.Drawing.Point(131, 31)
        Me.Label190.Name = "Label190"
        Me.Label190.Size = New System.Drawing.Size(92, 16)
        Me.Label190.TabIndex = 2
        Me.Label190.Text = "( 抽出件数 = "
        Me.Label190.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenYanuso
        '
        Me.lblCntJizenYanuso.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenYanuso.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenYanuso.Name = "lblCntJizenYanuso"
        Me.lblCntJizenYanuso.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenYanuso.TabIndex = 3
        Me.lblCntJizenYanuso.Text = "9,999,999"
        Me.lblCntJizenYanuso.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlJizenBkhourei
        '
        Me.pnlJizenBkhourei.Controls.Add(Me.chkJizenBkhourei)
        Me.pnlJizenBkhourei.Controls.Add(Me.Label94)
        Me.pnlJizenBkhourei.Controls.Add(Me.Panel4)
        Me.pnlJizenBkhourei.Location = New System.Drawing.Point(6, 174)
        Me.pnlJizenBkhourei.Name = "pnlJizenBkhourei"
        Me.pnlJizenBkhourei.Size = New System.Drawing.Size(400, 160)
        Me.pnlJizenBkhourei.TabIndex = 1
        '
        'chkJizenBkhourei
        '
        Me.chkJizenBkhourei.AutoSize = true
        Me.chkJizenBkhourei.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenBkhourei.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenBkhourei.Name = "chkJizenBkhourei"
        Me.chkJizenBkhourei.Size = New System.Drawing.Size(183, 22)
        Me.chkJizenBkhourei.TabIndex = 3
        Me.chkJizenBkhourei.Text = "物件法令・権利の設定値調整"
        Me.chkJizenBkhourei.UseVisualStyleBackColor = true
        '
        'Label94
        '
        Me.Label94.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label94.Location = New System.Drawing.Point(36, 29)
        Me.Label94.Name = "Label94"
        Me.Label94.Size = New System.Drawing.Size(340, 65)
        Me.Label94.TabIndex = 4
        Me.Label94.Text = "初期設定値(V7の初期導入時の登録値)で登録されたデータを対象として移行します(それ以外は移行できません)。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　①初期設定値以外の設定値で移行できないデータ"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)& _ 
    "　②初期設定値だが10に移行先がない為に移行できないデータ"
        Me.Label94.UseCompatibleTextRendering = true
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Label179)
        Me.Panel4.Controls.Add(Me.Label198)
        Me.Panel4.Controls.Add(Me.lblCntJizenBkhoureiNotCV)
        Me.Panel4.Controls.Add(Me.btnListJizenBkhourei)
        Me.Panel4.Controls.Add(Me.Label103)
        Me.Panel4.Controls.Add(Me.Label117)
        Me.Panel4.Controls.Add(Me.Label171)
        Me.Panel4.Controls.Add(Me.lblCntJizenBkhoureiNotDef)
        Me.Panel4.Location = New System.Drawing.Point(36, 97)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(340, 60)
        Me.Panel4.TabIndex = 5
        '
        'Label179
        '
        Me.Label179.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label179.Location = New System.Drawing.Point(307, 39)
        Me.Label179.Name = "Label179"
        Me.Label179.Size = New System.Drawing.Size(22, 16)
        Me.Label179.TabIndex = 7
        Me.Label179.Text = " )"
        Me.Label179.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label198
        '
        Me.Label198.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label198.Location = New System.Drawing.Point(131, 39)
        Me.Label198.Name = "Label198"
        Me.Label198.Size = New System.Drawing.Size(92, 16)
        Me.Label198.TabIndex = 5
        Me.Label198.Text = "( ②抽出件数 = "
        Me.Label198.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenBkhoureiNotCV
        '
        Me.lblCntJizenBkhoureiNotCV.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenBkhoureiNotCV.Location = New System.Drawing.Point(218, 39)
        Me.lblCntJizenBkhoureiNotCV.Name = "lblCntJizenBkhoureiNotCV"
        Me.lblCntJizenBkhoureiNotCV.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenBkhoureiNotCV.TabIndex = 6
        Me.lblCntJizenBkhoureiNotCV.Text = "9,999,999"
        Me.lblCntJizenBkhoureiNotCV.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnListJizenBkhourei
        '
        Me.btnListJizenBkhourei.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenBkhourei.Image = CType(resources.GetObject("btnListJizenBkhourei.Image"),System.Drawing.Image)
        Me.btnListJizenBkhourei.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenBkhourei.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenBkhourei.Name = "btnListJizenBkhourei"
        Me.btnListJizenBkhourei.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenBkhourei.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenBkhourei.TabIndex = 1
        Me.btnListJizenBkhourei.Text = "リスト出力"
        Me.btnListJizenBkhourei.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenBkhourei.UseVisualStyleBackColor = true
        '
        'Label103
        '
        Me.Label103.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label103.Location = New System.Drawing.Point(8, 8)
        Me.Label103.Name = "Label103"
        Me.Label103.Size = New System.Drawing.Size(321, 16)
        Me.Label103.TabIndex = 0
        Me.Label103.Text = "物件法令・権利データのリスト出力"
        Me.Label103.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label103.UseCompatibleTextRendering = true
        '
        'Label117
        '
        Me.Label117.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label117.Location = New System.Drawing.Point(307, 23)
        Me.Label117.Name = "Label117"
        Me.Label117.Size = New System.Drawing.Size(22, 16)
        Me.Label117.TabIndex = 4
        Me.Label117.Text = " )"
        Me.Label117.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label171
        '
        Me.Label171.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label171.Location = New System.Drawing.Point(131, 23)
        Me.Label171.Name = "Label171"
        Me.Label171.Size = New System.Drawing.Size(92, 16)
        Me.Label171.TabIndex = 2
        Me.Label171.Text = "( ①抽出件数 = "
        Me.Label171.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenBkhoureiNotDef
        '
        Me.lblCntJizenBkhoureiNotDef.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenBkhoureiNotDef.Location = New System.Drawing.Point(218, 23)
        Me.lblCntJizenBkhoureiNotDef.Name = "lblCntJizenBkhoureiNotDef"
        Me.lblCntJizenBkhoureiNotDef.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenBkhoureiNotDef.TabIndex = 3
        Me.lblCntJizenBkhoureiNotDef.Text = "9,999,999"
        Me.lblCntJizenBkhoureiNotDef.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblJizenPageCnt2
        '
        Me.lblJizenPageCnt2.AutoSize = true
        Me.lblJizenPageCnt2.Location = New System.Drawing.Point(794, 321)
        Me.lblJizenPageCnt2.Name = "lblJizenPageCnt2"
        Me.lblJizenPageCnt2.Size = New System.Drawing.Size(35, 18)
        Me.lblJizenPageCnt2.TabIndex = 12
        Me.lblJizenPageCnt2.Text = "2 / 5"
        '
        'tabPageJizen3
        '
        Me.tabPageJizen3.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJizen3.Controls.Add(Me.grpJizen3)
        Me.tabPageJizen3.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJizen3.Name = "tabPageJizen3"
        Me.tabPageJizen3.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageJizen3.Size = New System.Drawing.Size(847, 363)
        Me.tabPageJizen3.TabIndex = 2
        Me.tabPageJizen3.Text = " 事前作業3"
        '
        'grpJizen3
        '
        Me.grpJizen3.BackColor = System.Drawing.SystemColors.Menu
        Me.grpJizen3.Controls.Add(Me.pnlJizenKagi)
        Me.grpJizen3.Controls.Add(Me.pnlJizenSzenKyshutan)
        Me.grpJizen3.Controls.Add(Me.Label224)
        Me.grpJizen3.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpJizen3.Location = New System.Drawing.Point(6, 15)
        Me.grpJizen3.Name = "grpJizen3"
        Me.grpJizen3.Size = New System.Drawing.Size(835, 342)
        Me.grpJizen3.TabIndex = 0
        Me.grpJizen3.TabStop = false
        Me.grpJizen3.Text = "【賃貸革命10仕様によるユーザ確認・調整】"
        '
        'pnlJizenKagi
        '
        Me.pnlJizenKagi.Controls.Add(Me.chkJizenKagi)
        Me.pnlJizenKagi.Controls.Add(Me.Label262)
        Me.pnlJizenKagi.Controls.Add(Me.Panel19)
        Me.pnlJizenKagi.Location = New System.Drawing.Point(429, 18)
        Me.pnlJizenKagi.Name = "pnlJizenKagi"
        Me.pnlJizenKagi.Size = New System.Drawing.Size(400, 200)
        Me.pnlJizenKagi.TabIndex = 2
        '
        'chkJizenKagi
        '
        Me.chkJizenKagi.AutoSize = true
        Me.chkJizenKagi.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenKagi.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenKagi.Name = "chkJizenKagi"
        Me.chkJizenKagi.Size = New System.Drawing.Size(183, 22)
        Me.chkJizenKagi.TabIndex = 0
        Me.chkJizenKagi.Text = "部屋鍵・契約鍵データの確認"
        Me.chkJizenKagi.UseVisualStyleBackColor = true
        '
        'Label262
        '
        Me.Label262.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label262.Location = New System.Drawing.Point(36, 29)
        Me.Label262.Name = "Label262"
        Me.Label262.Size = New System.Drawing.Size(340, 65)
        Me.Label262.TabIndex = 1
        Me.Label262.Text = "10と仕様が異なる為、V7の部屋鍵と契約鍵のどちらかの情報を選択する必要があります(10では共通鍵と個別鍵(部屋鍵)の設定)。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"事前に鍵内容の確認と調整を行って"& _ 
    "下さい(必要な場合)。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"    (①部屋鍵データ, ②契約鍵データ)"
        Me.Label262.UseCompatibleTextRendering = true
        '
        'Panel19
        '
        Me.Panel19.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel19.Controls.Add(Me.btnListJizenKyKagi)
        Me.Panel19.Controls.Add(Me.Label266)
        Me.Panel19.Controls.Add(Me.Label269)
        Me.Panel19.Controls.Add(Me.lblCntJizenKyKagi)
        Me.Panel19.Controls.Add(Me.btnListJizenHyKagi)
        Me.Panel19.Controls.Add(Me.Label276)
        Me.Panel19.Controls.Add(Me.Label279)
        Me.Panel19.Controls.Add(Me.Label283)
        Me.Panel19.Controls.Add(Me.lblCntJizenHyKagi)
        Me.Panel19.Location = New System.Drawing.Point(36, 97)
        Me.Panel19.Name = "Panel19"
        Me.Panel19.Size = New System.Drawing.Size(340, 95)
        Me.Panel19.TabIndex = 2
        '
        'btnListJizenKyKagi
        '
        Me.btnListJizenKyKagi.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenKyKagi.Image = CType(resources.GetObject("btnListJizenKyKagi.Image"),System.Drawing.Image)
        Me.btnListJizenKyKagi.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenKyKagi.Location = New System.Drawing.Point(8, 56)
        Me.btnListJizenKyKagi.Name = "btnListJizenKyKagi"
        Me.btnListJizenKyKagi.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenKyKagi.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenKyKagi.TabIndex = 8
        Me.btnListJizenKyKagi.Text = "リスト出力"
        Me.btnListJizenKyKagi.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenKyKagi.UseVisualStyleBackColor = true
        '
        'Label266
        '
        Me.Label266.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label266.Location = New System.Drawing.Point(307, 59)
        Me.Label266.Name = "Label266"
        Me.Label266.Size = New System.Drawing.Size(22, 16)
        Me.Label266.TabIndex = 7
        Me.Label266.Text = " )"
        Me.Label266.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label269
        '
        Me.Label269.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label269.Location = New System.Drawing.Point(131, 59)
        Me.Label269.Name = "Label269"
        Me.Label269.Size = New System.Drawing.Size(92, 16)
        Me.Label269.TabIndex = 5
        Me.Label269.Text = "( ②抽出件数 = "
        Me.Label269.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenKyKagi
        '
        Me.lblCntJizenKyKagi.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenKyKagi.Location = New System.Drawing.Point(218, 59)
        Me.lblCntJizenKyKagi.Name = "lblCntJizenKyKagi"
        Me.lblCntJizenKyKagi.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenKyKagi.TabIndex = 6
        Me.lblCntJizenKyKagi.Text = "9,999,999"
        Me.lblCntJizenKyKagi.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnListJizenHyKagi
        '
        Me.btnListJizenHyKagi.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenHyKagi.Image = CType(resources.GetObject("btnListJizenHyKagi.Image"),System.Drawing.Image)
        Me.btnListJizenHyKagi.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenHyKagi.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenHyKagi.Name = "btnListJizenHyKagi"
        Me.btnListJizenHyKagi.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenHyKagi.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenHyKagi.TabIndex = 1
        Me.btnListJizenHyKagi.Text = "リスト出力"
        Me.btnListJizenHyKagi.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenHyKagi.UseVisualStyleBackColor = true
        '
        'Label276
        '
        Me.Label276.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label276.Location = New System.Drawing.Point(8, 8)
        Me.Label276.Name = "Label276"
        Me.Label276.Size = New System.Drawing.Size(321, 16)
        Me.Label276.TabIndex = 0
        Me.Label276.Text = "鍵データのリスト出力"
        Me.Label276.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label276.UseCompatibleTextRendering = true
        '
        'Label279
        '
        Me.Label279.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label279.Location = New System.Drawing.Point(307, 30)
        Me.Label279.Name = "Label279"
        Me.Label279.Size = New System.Drawing.Size(22, 16)
        Me.Label279.TabIndex = 4
        Me.Label279.Text = " )"
        Me.Label279.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label283
        '
        Me.Label283.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label283.Location = New System.Drawing.Point(131, 30)
        Me.Label283.Name = "Label283"
        Me.Label283.Size = New System.Drawing.Size(92, 16)
        Me.Label283.TabIndex = 2
        Me.Label283.Text = "( ①抽出件数 = "
        Me.Label283.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenHyKagi
        '
        Me.lblCntJizenHyKagi.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenHyKagi.Location = New System.Drawing.Point(218, 30)
        Me.lblCntJizenHyKagi.Name = "lblCntJizenHyKagi"
        Me.lblCntJizenHyKagi.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenHyKagi.TabIndex = 3
        Me.lblCntJizenHyKagi.Text = "9,999,999"
        Me.lblCntJizenHyKagi.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlJizenSzenKyshutan
        '
        Me.pnlJizenSzenKyshutan.Controls.Add(Me.chkJizenSzenKyshutan)
        Me.pnlJizenSzenKyshutan.Controls.Add(Me.Label173)
        Me.pnlJizenSzenKyshutan.Controls.Add(Me.Panel11)
        Me.pnlJizenSzenKyshutan.Location = New System.Drawing.Point(6, 18)
        Me.pnlJizenSzenKyshutan.Name = "pnlJizenSzenKyshutan"
        Me.pnlJizenSzenKyshutan.Size = New System.Drawing.Size(400, 180)
        Me.pnlJizenSzenKyshutan.TabIndex = 1
        '
        'chkJizenSzenKyshutan
        '
        Me.chkJizenSzenKyshutan.AutoSize = true
        Me.chkJizenSzenKyshutan.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenSzenKyshutan.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenSzenKyshutan.Name = "chkJizenSzenKyshutan"
        Me.chkJizenSzenKyshutan.Size = New System.Drawing.Size(243, 22)
        Me.chkJizenSzenKyshutan.TabIndex = 0
        Me.chkJizenSzenKyshutan.Text = "修繕対象負担者設定契約者データの確認"
        Me.chkJizenSzenKyshutan.UseVisualStyleBackColor = true
        '
        'Label173
        '
        Me.Label173.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label173.Location = New System.Drawing.Point(36, 29)
        Me.Label173.Name = "Label173"
        Me.Label173.Size = New System.Drawing.Size(340, 86)
        Me.Label173.TabIndex = 1
        Me.Label173.Text = "物件共有部の修繕の場合、V7では契約者も負担者に設定できますが、10ではできません。負担者に契約者が設定されているデータが存在する場合、V7のデータの調整を行って"& _ 
    "下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"例) V7の物件共有部のリフォームにおける契約者負担データを削除し、その他請求へ登録する。"
        Me.Label173.UseCompatibleTextRendering = true
        '
        'Panel11
        '
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel11.Controls.Add(Me.btnListJizenSzenKyshutan)
        Me.Panel11.Controls.Add(Me.Label202)
        Me.Panel11.Controls.Add(Me.Label203)
        Me.Panel11.Controls.Add(Me.Label204)
        Me.Panel11.Controls.Add(Me.lblCntJizenSzenKyshutan)
        Me.Panel11.Location = New System.Drawing.Point(36, 114)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(340, 60)
        Me.Panel11.TabIndex = 2
        '
        'btnListJizenSzenKyshutan
        '
        Me.btnListJizenSzenKyshutan.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenSzenKyshutan.Image = CType(resources.GetObject("btnListJizenSzenKyshutan.Image"),System.Drawing.Image)
        Me.btnListJizenSzenKyshutan.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenSzenKyshutan.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenSzenKyshutan.Name = "btnListJizenSzenKyshutan"
        Me.btnListJizenSzenKyshutan.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenSzenKyshutan.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenSzenKyshutan.TabIndex = 1
        Me.btnListJizenSzenKyshutan.Text = "リスト出力"
        Me.btnListJizenSzenKyshutan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenSzenKyshutan.UseVisualStyleBackColor = true
        '
        'Label202
        '
        Me.Label202.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label202.Location = New System.Drawing.Point(8, 8)
        Me.Label202.Name = "Label202"
        Me.Label202.Size = New System.Drawing.Size(321, 16)
        Me.Label202.TabIndex = 0
        Me.Label202.Text = "物件共有部修繕の負担者が契約者のデータのリスト出力"
        Me.Label202.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label202.UseCompatibleTextRendering = true
        '
        'Label203
        '
        Me.Label203.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label203.Location = New System.Drawing.Point(307, 31)
        Me.Label203.Name = "Label203"
        Me.Label203.Size = New System.Drawing.Size(22, 16)
        Me.Label203.TabIndex = 4
        Me.Label203.Text = " )"
        Me.Label203.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label204
        '
        Me.Label204.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label204.Location = New System.Drawing.Point(131, 31)
        Me.Label204.Name = "Label204"
        Me.Label204.Size = New System.Drawing.Size(92, 16)
        Me.Label204.TabIndex = 2
        Me.Label204.Text = "( 抽出件数 = "
        Me.Label204.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenSzenKyshutan
        '
        Me.lblCntJizenSzenKyshutan.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenSzenKyshutan.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenSzenKyshutan.Name = "lblCntJizenSzenKyshutan"
        Me.lblCntJizenSzenKyshutan.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenSzenKyshutan.TabIndex = 3
        Me.lblCntJizenSzenKyshutan.Text = "9,999,999"
        Me.lblCntJizenSzenKyshutan.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label224
        '
        Me.Label224.AutoSize = true
        Me.Label224.Location = New System.Drawing.Point(794, 321)
        Me.Label224.Name = "Label224"
        Me.Label224.Size = New System.Drawing.Size(35, 18)
        Me.Label224.TabIndex = 6
        Me.Label224.Text = "3 / 5"
        '
        'tabPageJizen4
        '
        Me.tabPageJizen4.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJizen4.Controls.Add(Me.grpJizen4)
        Me.tabPageJizen4.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJizen4.Name = "tabPageJizen4"
        Me.tabPageJizen4.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageJizen4.Size = New System.Drawing.Size(847, 363)
        Me.tabPageJizen4.TabIndex = 4
        Me.tabPageJizen4.Text = " 事前作業4"
        '
        'grpJizen4
        '
        Me.grpJizen4.BackColor = System.Drawing.SystemColors.Menu
        Me.grpJizen4.Controls.Add(Me.Panel15)
        Me.grpJizen4.Controls.Add(Me.Panel21)
        Me.grpJizen4.Controls.Add(Me.Label370)
        Me.grpJizen4.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpJizen4.Location = New System.Drawing.Point(6, 15)
        Me.grpJizen4.Name = "grpJizen4"
        Me.grpJizen4.Size = New System.Drawing.Size(835, 342)
        Me.grpJizen4.TabIndex = 0
        Me.grpJizen4.TabStop = false
        Me.grpJizen4.Text = "【賃貸革命10仕様によるユーザ確認・調整】"
        '
        'Panel15
        '
        Me.Panel15.Controls.Add(Me.chkJizenKozameigikana)
        Me.Panel15.Controls.Add(Me.Label353)
        Me.Panel15.Controls.Add(Me.Panel20)
        Me.Panel15.Location = New System.Drawing.Point(429, 18)
        Me.Panel15.Name = "Panel15"
        Me.Panel15.Size = New System.Drawing.Size(400, 235)
        Me.Panel15.TabIndex = 8
        '
        'chkJizenKozameigikana
        '
        Me.chkJizenKozameigikana.AutoSize = true
        Me.chkJizenKozameigikana.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenKozameigikana.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenKozameigikana.Name = "chkJizenKozameigikana"
        Me.chkJizenKozameigikana.Size = New System.Drawing.Size(363, 22)
        Me.chkJizenKozameigikana.TabIndex = 0
        Me.chkJizenKozameigikana.Text = "使用不可文字が含まれている口座名義カナデータの確認・調整"
        Me.chkJizenKozameigikana.UseVisualStyleBackColor = true
        '
        'Label353
        '
        Me.Label353.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label353.Location = New System.Drawing.Point(36, 29)
        Me.Label353.Name = "Label353"
        Me.Label353.Size = New System.Drawing.Size(345, 133)
        Me.Label353.TabIndex = 1
        Me.Label353.Text = resources.GetString("Label353.Text")
        Me.Label353.UseCompatibleTextRendering = true
        '
        'Panel20
        '
        Me.Panel20.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel20.Controls.Add(Me.btnListJizenKozameigikana)
        Me.Panel20.Controls.Add(Me.Label354)
        Me.Panel20.Controls.Add(Me.Label356)
        Me.Panel20.Controls.Add(Me.Label357)
        Me.Panel20.Controls.Add(Me.lblCntJizenKozameigikana)
        Me.Panel20.Location = New System.Drawing.Point(36, 165)
        Me.Panel20.Name = "Panel20"
        Me.Panel20.Size = New System.Drawing.Size(340, 60)
        Me.Panel20.TabIndex = 2
        '
        'btnListJizenKozameigikana
        '
        Me.btnListJizenKozameigikana.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenKozameigikana.Image = CType(resources.GetObject("btnListJizenKozameigikana.Image"),System.Drawing.Image)
        Me.btnListJizenKozameigikana.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenKozameigikana.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenKozameigikana.Name = "btnListJizenKozameigikana"
        Me.btnListJizenKozameigikana.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenKozameigikana.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenKozameigikana.TabIndex = 1
        Me.btnListJizenKozameigikana.Text = "リスト出力"
        Me.btnListJizenKozameigikana.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenKozameigikana.UseVisualStyleBackColor = true
        '
        'Label354
        '
        Me.Label354.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label354.Location = New System.Drawing.Point(8, 8)
        Me.Label354.Name = "Label354"
        Me.Label354.Size = New System.Drawing.Size(321, 16)
        Me.Label354.TabIndex = 0
        Me.Label354.Text = "使用不可文字が含まれている口座名義カナデータのリスト出力"
        Me.Label354.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label354.UseCompatibleTextRendering = true
        '
        'Label356
        '
        Me.Label356.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label356.Location = New System.Drawing.Point(307, 31)
        Me.Label356.Name = "Label356"
        Me.Label356.Size = New System.Drawing.Size(22, 16)
        Me.Label356.TabIndex = 4
        Me.Label356.Text = " )"
        Me.Label356.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label357
        '
        Me.Label357.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label357.Location = New System.Drawing.Point(131, 31)
        Me.Label357.Name = "Label357"
        Me.Label357.Size = New System.Drawing.Size(92, 16)
        Me.Label357.TabIndex = 2
        Me.Label357.Text = "( 抽出件数 = "
        Me.Label357.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenKozameigikana
        '
        Me.lblCntJizenKozameigikana.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenKozameigikana.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenKozameigikana.Name = "lblCntJizenKozameigikana"
        Me.lblCntJizenKozameigikana.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenKozameigikana.TabIndex = 3
        Me.lblCntJizenKozameigikana.Text = "9,999,999"
        Me.lblCntJizenKozameigikana.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel21
        '
        Me.Panel21.Controls.Add(Me.chkJizenNonJisyaKoza)
        Me.Panel21.Controls.Add(Me.Label365)
        Me.Panel21.Controls.Add(Me.Panel27)
        Me.Panel21.Location = New System.Drawing.Point(6, 18)
        Me.Panel21.Name = "Panel21"
        Me.Panel21.Size = New System.Drawing.Size(400, 206)
        Me.Panel21.TabIndex = 0
        '
        'chkJizenNonJisyaKoza
        '
        Me.chkJizenNonJisyaKoza.AutoSize = true
        Me.chkJizenNonJisyaKoza.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenNonJisyaKoza.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenNonJisyaKoza.Name = "chkJizenNonJisyaKoza"
        Me.chkJizenNonJisyaKoza.Size = New System.Drawing.Size(231, 22)
        Me.chkJizenNonJisyaKoza.TabIndex = 0
        Me.chkJizenNonJisyaKoza.Text = "口座実情報未登録データの確認・調整"
        Me.chkJizenNonJisyaKoza.UseVisualStyleBackColor = true
        '
        'Label365
        '
        Me.Label365.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label365.Location = New System.Drawing.Point(36, 29)
        Me.Label365.Name = "Label365"
        Me.Label365.Size = New System.Drawing.Size(340, 107)
        Me.Label365.TabIndex = 1
        Me.Label365.Text = "10の自社口座では、実口座情報の登録が必須となります(口座名のみで口座情報なしの登録はできません)。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"また、10の口座振替・総合振込では自社口座、家賃入金口座で"& _ 
    "は自社口座または家主口座を紐付ける必要があります。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"その為、実口座情報が未登録であるデータについて、事前に実口座情報を登録しておく必要があります。"
        Me.Label365.UseCompatibleTextRendering = true
        '
        'Panel27
        '
        Me.Panel27.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel27.Controls.Add(Me.btnListJizenNonJisyaKoza)
        Me.Panel27.Controls.Add(Me.Label366)
        Me.Panel27.Controls.Add(Me.Label367)
        Me.Panel27.Controls.Add(Me.Label368)
        Me.Panel27.Controls.Add(Me.lblCntJizenNonJisyaKoza)
        Me.Panel27.Location = New System.Drawing.Point(36, 136)
        Me.Panel27.Name = "Panel27"
        Me.Panel27.Size = New System.Drawing.Size(340, 60)
        Me.Panel27.TabIndex = 2
        '
        'btnListJizenNonJisyaKoza
        '
        Me.btnListJizenNonJisyaKoza.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenNonJisyaKoza.Image = CType(resources.GetObject("btnListJizenNonJisyaKoza.Image"),System.Drawing.Image)
        Me.btnListJizenNonJisyaKoza.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenNonJisyaKoza.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenNonJisyaKoza.Name = "btnListJizenNonJisyaKoza"
        Me.btnListJizenNonJisyaKoza.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenNonJisyaKoza.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenNonJisyaKoza.TabIndex = 1
        Me.btnListJizenNonJisyaKoza.Text = "リスト出力"
        Me.btnListJizenNonJisyaKoza.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenNonJisyaKoza.UseVisualStyleBackColor = true
        '
        'Label366
        '
        Me.Label366.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label366.Location = New System.Drawing.Point(8, 8)
        Me.Label366.Name = "Label366"
        Me.Label366.Size = New System.Drawing.Size(321, 16)
        Me.Label366.TabIndex = 0
        Me.Label366.Text = "口座実情報未登録データのリスト出力"
        Me.Label366.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label366.UseCompatibleTextRendering = true
        '
        'Label367
        '
        Me.Label367.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label367.Location = New System.Drawing.Point(307, 31)
        Me.Label367.Name = "Label367"
        Me.Label367.Size = New System.Drawing.Size(22, 16)
        Me.Label367.TabIndex = 4
        Me.Label367.Text = " )"
        Me.Label367.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label368
        '
        Me.Label368.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label368.Location = New System.Drawing.Point(131, 31)
        Me.Label368.Name = "Label368"
        Me.Label368.Size = New System.Drawing.Size(92, 16)
        Me.Label368.TabIndex = 2
        Me.Label368.Text = "( 抽出件数 = "
        Me.Label368.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenNonJisyaKoza
        '
        Me.lblCntJizenNonJisyaKoza.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenNonJisyaKoza.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenNonJisyaKoza.Name = "lblCntJizenNonJisyaKoza"
        Me.lblCntJizenNonJisyaKoza.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenNonJisyaKoza.TabIndex = 3
        Me.lblCntJizenNonJisyaKoza.Text = "9,999,999"
        Me.lblCntJizenNonJisyaKoza.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label370
        '
        Me.Label370.AutoSize = true
        Me.Label370.Location = New System.Drawing.Point(794, 321)
        Me.Label370.Name = "Label370"
        Me.Label370.Size = New System.Drawing.Size(35, 18)
        Me.Label370.TabIndex = 6
        Me.Label370.Text = "4 / 5"
        '
        'tabPageJizen5
        '
        Me.tabPageJizen5.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJizen5.Controls.Add(Me.GroupBox4)
        Me.tabPageJizen5.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJizen5.Name = "tabPageJizen5"
        Me.tabPageJizen5.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageJizen5.Size = New System.Drawing.Size(847, 363)
        Me.tabPageJizen5.TabIndex = 5
        Me.tabPageJizen5.Text = " 事前作業5"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox4.Controls.Add(Me.Label374)
        Me.GroupBox4.Controls.Add(Me.Panel13)
        Me.GroupBox4.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(6, 15)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(835, 342)
        Me.GroupBox4.TabIndex = 8
        Me.GroupBox4.TabStop = false
        Me.GroupBox4.Text = "【賃貸革命10仕様によるユーザ確認・調整】"
        '
        'Label374
        '
        Me.Label374.AutoSize = true
        Me.Label374.Location = New System.Drawing.Point(794, 321)
        Me.Label374.Name = "Label374"
        Me.Label374.Size = New System.Drawing.Size(35, 18)
        Me.Label374.TabIndex = 6
        Me.Label374.Text = "5 / 5"
        '
        'Panel13
        '
        Me.Panel13.Controls.Add(Me.chkJizenSimeSokin)
        Me.Panel13.Controls.Add(Me.Label149)
        Me.Panel13.Controls.Add(Me.Panel14)
        Me.Panel13.Location = New System.Drawing.Point(18, 24)
        Me.Panel13.Name = "Panel13"
        Me.Panel13.Size = New System.Drawing.Size(400, 302)
        Me.Panel13.TabIndex = 7
        '
        'chkJizenSimeSokin
        '
        Me.chkJizenSimeSokin.AutoSize = true
        Me.chkJizenSimeSokin.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenSimeSokin.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenSimeSokin.Name = "chkJizenSimeSokin"
        Me.chkJizenSimeSokin.Size = New System.Drawing.Size(370, 22)
        Me.chkJizenSimeSokin.TabIndex = 0
        Me.chkJizenSimeSokin.Text = "複数の送金予定設定の相互間隔が1ヶ月以上の設定値確認・調整"
        Me.chkJizenSimeSokin.UseVisualStyleBackColor = true
        '
        'Label149
        '
        Me.Label149.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label149.Location = New System.Drawing.Point(36, 29)
        Me.Label149.Name = "Label149"
        Me.Label149.Size = New System.Drawing.Size(340, 204)
        Me.Label149.TabIndex = 1
        Me.Label149.Text = resources.GetString("Label149.Text")
        Me.Label149.UseCompatibleTextRendering = true
        '
        'Panel14
        '
        Me.Panel14.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel14.Controls.Add(Me.btnListJizenSimeSokin)
        Me.Panel14.Controls.Add(Me.Label310)
        Me.Panel14.Controls.Add(Me.Label348)
        Me.Panel14.Controls.Add(Me.Label351)
        Me.Panel14.Controls.Add(Me.lblCntJizenSimeSokin)
        Me.Panel14.Location = New System.Drawing.Point(32, 236)
        Me.Panel14.Name = "Panel14"
        Me.Panel14.Size = New System.Drawing.Size(340, 60)
        Me.Panel14.TabIndex = 2
        '
        'btnListJizenSimeSokin
        '
        Me.btnListJizenSimeSokin.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenSimeSokin.Image = CType(resources.GetObject("btnListJizenSimeSokin.Image"),System.Drawing.Image)
        Me.btnListJizenSimeSokin.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenSimeSokin.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenSimeSokin.Name = "btnListJizenSimeSokin"
        Me.btnListJizenSimeSokin.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenSimeSokin.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenSimeSokin.TabIndex = 1
        Me.btnListJizenSimeSokin.Text = "リスト出力"
        Me.btnListJizenSimeSokin.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenSimeSokin.UseVisualStyleBackColor = true
        '
        'Label310
        '
        Me.Label310.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label310.Location = New System.Drawing.Point(8, 8)
        Me.Label310.Name = "Label310"
        Me.Label310.Size = New System.Drawing.Size(321, 16)
        Me.Label310.TabIndex = 0
        Me.Label310.Text = "送金予定設定の相互間隔が1ヶ月以上のデータのリスト出力"
        Me.Label310.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label310.UseCompatibleTextRendering = true
        '
        'Label348
        '
        Me.Label348.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label348.Location = New System.Drawing.Point(307, 31)
        Me.Label348.Name = "Label348"
        Me.Label348.Size = New System.Drawing.Size(22, 16)
        Me.Label348.TabIndex = 4
        Me.Label348.Text = " )"
        Me.Label348.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label351
        '
        Me.Label351.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label351.Location = New System.Drawing.Point(131, 31)
        Me.Label351.Name = "Label351"
        Me.Label351.Size = New System.Drawing.Size(92, 16)
        Me.Label351.TabIndex = 2
        Me.Label351.Text = "( 抽出件数 = "
        Me.Label351.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenSimeSokin
        '
        Me.lblCntJizenSimeSokin.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenSimeSokin.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenSimeSokin.Name = "lblCntJizenSimeSokin"
        Me.lblCntJizenSimeSokin.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenSimeSokin.TabIndex = 3
        Me.lblCntJizenSimeSokin.Text = "9,999,999"
        Me.lblCntJizenSimeSokin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabPageHJizen1
        '
        Me.tabPageHJizen1.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHJizen1.Controls.Add(Me.GroupBox5)
        Me.tabPageHJizen1.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHJizen1.Name = "tabPageHJizen1"
        Me.tabPageHJizen1.Size = New System.Drawing.Size(847, 363)
        Me.tabPageHJizen1.TabIndex = 3
        Me.tabPageHJizen1.Text = " 事前作業1(汎用) "
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox5.Controls.Add(Me.pnlHJizenTyukan)
        Me.GroupBox5.Controls.Add(Me.Label355)
        Me.GroupBox5.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(6, 15)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(835, 342)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = false
        Me.GroupBox5.Text = "【ユーザによる事前確認・調整】"
        '
        'pnlHJizenTyukan
        '
        Me.pnlHJizenTyukan.Controls.Add(Me.Label352)
        Me.pnlHJizenTyukan.Controls.Add(Me.chkHJizenTyukan)
        Me.pnlHJizenTyukan.Controls.Add(Me.Label350)
        Me.pnlHJizenTyukan.Controls.Add(Me.Panel26)
        Me.pnlHJizenTyukan.Location = New System.Drawing.Point(6, 18)
        Me.pnlHJizenTyukan.Name = "pnlHJizenTyukan"
        Me.pnlHJizenTyukan.Size = New System.Drawing.Size(400, 305)
        Me.pnlHJizenTyukan.TabIndex = 0
        '
        'Label352
        '
        Me.Label352.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label352.ForeColor = System.Drawing.Color.Red
        Me.Label352.Location = New System.Drawing.Point(36, 234)
        Me.Label352.Name = "Label352"
        Me.Label352.Size = New System.Drawing.Size(340, 66)
        Me.Label352.TabIndex = 3
        Me.Label352.Text = "※中間ファイルのデータ登録が完了したら、必ず、中間ファイル(EXCELファイル)を保存・閉じて下さい。中間ファイルを開いたままコンバートを行った場合、正常にコンバ"& _ 
    "ートができない場合があります。"
        Me.Label352.UseCompatibleTextRendering = true
        '
        'chkHJizenTyukan
        '
        Me.chkHJizenTyukan.AutoSize = true
        Me.chkHJizenTyukan.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkHJizenTyukan.Location = New System.Drawing.Point(16, 4)
        Me.chkHJizenTyukan.Name = "chkHJizenTyukan"
        Me.chkHJizenTyukan.Size = New System.Drawing.Size(147, 22)
        Me.chkHJizenTyukan.TabIndex = 0
        Me.chkHJizenTyukan.Text = "中間ファイル登録作業"
        Me.chkHJizenTyukan.UseVisualStyleBackColor = true
        '
        'Label350
        '
        Me.Label350.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label350.Location = New System.Drawing.Point(36, 29)
        Me.Label350.Name = "Label350"
        Me.Label350.Size = New System.Drawing.Size(340, 85)
        Me.Label350.TabIndex = 1
        Me.Label350.Text = "・中間ファイル(本システムにて用意したEXCELファイル)へ"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　ユーザデータの手動登録を行う必要があります。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"・登録方法については、中間ファイルのヘッダ部分に"& _ 
    "記載"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　している内容、または移行仕様書を参照下さい。"
        Me.Label350.UseCompatibleTextRendering = true
        '
        'Panel26
        '
        Me.Panel26.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel26.Controls.Add(Me.Label322)
        Me.Panel26.Controls.Add(Me.btnHJizenTyukanOpen)
        Me.Panel26.Controls.Add(Me.lblMiddleFile)
        Me.Panel26.Controls.Add(Me.txtMidDirPath)
        Me.Panel26.Controls.Add(Me.btnMidDirSeach)
        Me.Panel26.Location = New System.Drawing.Point(36, 117)
        Me.Panel26.Name = "Panel26"
        Me.Panel26.Size = New System.Drawing.Size(340, 110)
        Me.Panel26.TabIndex = 2
        '
        'Label322
        '
        Me.Label322.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label322.Location = New System.Drawing.Point(8, 8)
        Me.Label322.Name = "Label322"
        Me.Label322.Size = New System.Drawing.Size(325, 20)
        Me.Label322.TabIndex = 0
        Me.Label322.Text = "中間ファイルを開いて、ユーザデータを登録して下さい。"
        Me.Label322.UseCompatibleTextRendering = true
        '
        'btnHJizenTyukanOpen
        '
        Me.btnHJizenTyukanOpen.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnHJizenTyukanOpen.Image = CType(resources.GetObject("btnHJizenTyukanOpen.Image"),System.Drawing.Image)
        Me.btnHJizenTyukanOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnHJizenTyukanOpen.Location = New System.Drawing.Point(166, 79)
        Me.btnHJizenTyukanOpen.Name = "btnHJizenTyukanOpen"
        Me.btnHJizenTyukanOpen.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnHJizenTyukanOpen.Size = New System.Drawing.Size(157, 25)
        Me.btnHJizenTyukanOpen.TabIndex = 4
        Me.btnHJizenTyukanOpen.Text = " 中間ファイルを開く"
        Me.btnHJizenTyukanOpen.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnHJizenTyukanOpen.UseVisualStyleBackColor = true
        '
        'lblMiddleFile
        '
        Me.lblMiddleFile.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblMiddleFile.Location = New System.Drawing.Point(29, 28)
        Me.lblMiddleFile.Name = "lblMiddleFile"
        Me.lblMiddleFile.Size = New System.Drawing.Size(133, 20)
        Me.lblMiddleFile.TabIndex = 1
        Me.lblMiddleFile.Text = "中間ファイル格納先"
        Me.lblMiddleFile.UseCompatibleTextRendering = true
        '
        'txtMidDirPath
        '
        Me.txtMidDirPath.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtMidDirPath.Location = New System.Drawing.Point(31, 49)
        Me.txtMidDirPath.Name = "txtMidDirPath"
        Me.txtMidDirPath.Size = New System.Drawing.Size(256, 24)
        Me.txtMidDirPath.TabIndex = 2
        '
        'btnMidDirSeach
        '
        Me.btnMidDirSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMidDirSeach.Location = New System.Drawing.Point(293, 48)
        Me.btnMidDirSeach.Name = "btnMidDirSeach"
        Me.btnMidDirSeach.Size = New System.Drawing.Size(30, 25)
        Me.btnMidDirSeach.TabIndex = 3
        Me.btnMidDirSeach.Text = "..."
        Me.btnMidDirSeach.UseVisualStyleBackColor = true
        '
        'Label355
        '
        Me.Label355.AutoSize = true
        Me.Label355.Location = New System.Drawing.Point(794, 321)
        Me.Label355.Name = "Label355"
        Me.Label355.Size = New System.Drawing.Size(35, 18)
        Me.Label355.TabIndex = 10
        Me.Label355.Text = "1 / 3"
        '
        'lblCautionDescription
        '
        Me.lblCautionDescription.BackColor = System.Drawing.SystemColors.Menu
        Me.lblCautionDescription.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCautionDescription.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblCautionDescription.Location = New System.Drawing.Point(30, 63)
        Me.lblCautionDescription.Name = "lblCautionDescription"
        Me.lblCautionDescription.Size = New System.Drawing.Size(820, 20)
        Me.lblCautionDescription.TabIndex = 1
        Me.lblCautionDescription.Text = "※データコンバートを実行する前に必ず対応下さい。対応しなかった場合、異なる期待結果となる場合があります。"
        Me.lblCautionDescription.UseCompatibleTextRendering = true
        '
        'lblJizenDescription1
        '
        Me.lblJizenDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblJizenDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblJizenDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblJizenDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblJizenDescription1.Name = "lblJizenDescription1"
        Me.lblJizenDescription1.Size = New System.Drawing.Size(820, 80)
        Me.lblJizenDescription1.TabIndex = 0
        Me.lblJizenDescription1.Text = "対象となる賃貸革命V7のデータを、事前に調整する必要があります。以下の内容を確認し、必要な作業を行って下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(完了したら各チェックボックスをONにして下さい"& _ 
    "。全てONにした場合のみデータコンバートを行うことができます)"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)
        Me.lblJizenDescription1.UseCompatibleTextRendering = true
        '
        'tabPageJigo
        '
        Me.tabPageJigo.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJigo.Controls.Add(Me.pnlJigoListPath)
        Me.tabPageJigo.Controls.Add(Me.Label245)
        Me.tabPageJigo.Controls.Add(Me.lblLine3)
        Me.tabPageJigo.Controls.Add(Me.lblHidden4)
        Me.tabPageJigo.Controls.Add(Me.tabCtrlJigo)
        Me.tabPageJigo.Controls.Add(Me.lblJigoDescription1)
        Me.tabPageJigo.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJigo.Name = "tabPageJigo"
        Me.tabPageJigo.Size = New System.Drawing.Size(912, 549)
        Me.tabPageJigo.TabIndex = 14
        Me.tabPageJigo.Text = "事後作業"
        '
        'pnlJigoListPath
        '
        Me.pnlJigoListPath.Controls.Add(Me.btnJigoListDirSeach)
        Me.pnlJigoListPath.Controls.Add(Me.Label248)
        Me.pnlJigoListPath.Controls.Add(Me.txtJigoListPath)
        Me.pnlJigoListPath.Location = New System.Drawing.Point(31, 86)
        Me.pnlJigoListPath.Name = "pnlJigoListPath"
        Me.pnlJigoListPath.Size = New System.Drawing.Size(420, 60)
        Me.pnlJigoListPath.TabIndex = 13
        '
        'btnJigoListDirSeach
        '
        Me.btnJigoListDirSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnJigoListDirSeach.Location = New System.Drawing.Point(383, 27)
        Me.btnJigoListDirSeach.Name = "btnJigoListDirSeach"
        Me.btnJigoListDirSeach.Size = New System.Drawing.Size(30, 25)
        Me.btnJigoListDirSeach.TabIndex = 2
        Me.btnJigoListDirSeach.Text = "..."
        Me.btnJigoListDirSeach.UseVisualStyleBackColor = true
        '
        'Label248
        '
        Me.Label248.AutoSize = true
        Me.Label248.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label248.Location = New System.Drawing.Point(5, 5)
        Me.Label248.Name = "Label248"
        Me.Label248.Size = New System.Drawing.Size(116, 18)
        Me.Label248.TabIndex = 0
        Me.Label248.Text = "出力ファイル格納先"
        '
        'txtJigoListPath
        '
        Me.txtJigoListPath.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtJigoListPath.Location = New System.Drawing.Point(23, 28)
        Me.txtJigoListPath.Name = "txtJigoListPath"
        Me.txtJigoListPath.Size = New System.Drawing.Size(354, 24)
        Me.txtJigoListPath.TabIndex = 1
        '
        'Label245
        '
        Me.Label245.BackColor = System.Drawing.SystemColors.Menu
        Me.Label245.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label245.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label245.Location = New System.Drawing.Point(30, 63)
        Me.Label245.Name = "Label245"
        Me.Label245.Size = New System.Drawing.Size(820, 20)
        Me.Label245.TabIndex = 12
        Me.Label245.Text = "※データコンバートを実行後に必ず確認をして下さい。必要な作業を行わなかった場合、異なる期待結果となる場合があります。"
        Me.Label245.UseCompatibleTextRendering = true
        '
        'lblLine3
        '
        Me.lblLine3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLine3.Location = New System.Drawing.Point(31, 175)
        Me.lblLine3.Name = "lblLine3"
        Me.lblLine3.Size = New System.Drawing.Size(853, 2)
        Me.lblLine3.TabIndex = 11
        '
        'lblHidden4
        '
        Me.lblHidden4.Location = New System.Drawing.Point(6, 150)
        Me.lblHidden4.Name = "lblHidden4"
        Me.lblHidden4.Size = New System.Drawing.Size(10, 26)
        Me.lblHidden4.TabIndex = 10
        Me.lblHidden4.Text = "　"
        '
        'tabCtrlJigo
        '
        Me.tabCtrlJigo.Controls.Add(Me.tabPageJigo1)
        Me.tabCtrlJigo.Controls.Add(Me.tabPageJigo2)
        Me.tabCtrlJigo.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
        Me.tabCtrlJigo.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.tabCtrlJigo.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.tabCtrlJigo.Location = New System.Drawing.Point(30, 150)
        Me.tabCtrlJigo.Name = "tabCtrlJigo"
        Me.tabCtrlJigo.SelectedIndex = 0
        Me.tabCtrlJigo.Size = New System.Drawing.Size(855, 394)
        Me.tabCtrlJigo.TabIndex = 8
        Me.tabCtrlJigo.TabStop = false
        '
        'tabPageJigo1
        '
        Me.tabPageJigo1.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJigo1.Controls.Add(Me.grpJigoUserSagyo)
        Me.tabPageJigo1.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJigo1.Name = "tabPageJigo1"
        Me.tabPageJigo1.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageJigo1.Size = New System.Drawing.Size(847, 363)
        Me.tabPageJigo1.TabIndex = 4
        Me.tabPageJigo1.Text = " 事後作業1 "
        '
        'grpJigoUserSagyo
        '
        Me.grpJigoUserSagyo.BackColor = System.Drawing.SystemColors.Menu
        Me.grpJigoUserSagyo.Controls.Add(Me.pnlKJigoCmtSyudo)
        Me.grpJigoUserSagyo.Controls.Add(Me.pnlHJigoCmtSyudo)
        Me.grpJigoUserSagyo.Controls.Add(Me.pnlJigoCmtSyusi)
        Me.grpJigoUserSagyo.Controls.Add(Me.pnlJizenMinus)
        Me.grpJigoUserSagyo.Controls.Add(Me.pnlJigoCmtCsvSyuturyoku)
        Me.grpJigoUserSagyo.Controls.Add(Me.pnlJigoCmtSyosiki)
        Me.grpJigoUserSagyo.Controls.Add(Me.lblJizenPageNum1)
        Me.grpJigoUserSagyo.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpJigoUserSagyo.Location = New System.Drawing.Point(6, 15)
        Me.grpJigoUserSagyo.Name = "grpJigoUserSagyo"
        Me.grpJigoUserSagyo.Size = New System.Drawing.Size(835, 342)
        Me.grpJigoUserSagyo.TabIndex = 137
        Me.grpJigoUserSagyo.TabStop = false
        Me.grpJigoUserSagyo.Text = "【データ移行可否によるユーザ作業】"
        '
        'pnlKJigoCmtSyudo
        '
        Me.pnlKJigoCmtSyudo.Controls.Add(Me.Label306)
        Me.pnlKJigoCmtSyudo.Controls.Add(Me.Label305)
        Me.pnlKJigoCmtSyudo.Controls.Add(Me.Label143)
        Me.pnlKJigoCmtSyudo.Location = New System.Drawing.Point(6, 18)
        Me.pnlKJigoCmtSyudo.Name = "pnlKJigoCmtSyudo"
        Me.pnlKJigoCmtSyudo.Size = New System.Drawing.Size(400, 80)
        Me.pnlKJigoCmtSyudo.TabIndex = 137
        '
        'Label306
        '
        Me.Label306.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label306.Location = New System.Drawing.Point(32, 4)
        Me.Label306.Name = "Label306"
        Me.Label306.Size = New System.Drawing.Size(330, 20)
        Me.Label306.TabIndex = 117
        Me.Label306.Text = "コンバート後に手動設定が必要な項目について"
        Me.Label306.UseCompatibleTextRendering = true
        '
        'Label305
        '
        Me.Label305.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label305.Location = New System.Drawing.Point(36, 27)
        Me.Label305.Name = "Label305"
        Me.Label305.Size = New System.Drawing.Size(340, 50)
        Me.Label305.TabIndex = 118
        Me.Label305.Text = "クライアントやセキュリティなど一部の項目についてはユーザにて手動設定頂く必要があります。設定が必要な項目については、仕様書をご確認下さい。"
        Me.Label305.UseCompatibleTextRendering = true
        '
        'Label143
        '
        Me.Label143.Location = New System.Drawing.Point(11, 4)
        Me.Label143.Name = "Label143"
        Me.Label143.Size = New System.Drawing.Size(20, 20)
        Me.Label143.TabIndex = 134
        Me.Label143.Text = "・"
        Me.Label143.UseCompatibleTextRendering = true
        '
        'pnlHJigoCmtSyudo
        '
        Me.pnlHJigoCmtSyudo.Controls.Add(Me.Label136)
        Me.pnlHJigoCmtSyudo.Controls.Add(Me.Label317)
        Me.pnlHJigoCmtSyudo.Controls.Add(Me.Label318)
        Me.pnlHJigoCmtSyudo.Location = New System.Drawing.Point(6, 18)
        Me.pnlHJigoCmtSyudo.Name = "pnlHJigoCmtSyudo"
        Me.pnlHJigoCmtSyudo.Size = New System.Drawing.Size(375, 80)
        Me.pnlHJigoCmtSyudo.TabIndex = 140
        '
        'Label136
        '
        Me.Label136.Location = New System.Drawing.Point(11, 4)
        Me.Label136.Name = "Label136"
        Me.Label136.Size = New System.Drawing.Size(20, 20)
        Me.Label136.TabIndex = 137
        Me.Label136.Text = "・"
        Me.Label136.UseCompatibleTextRendering = true
        '
        'Label317
        '
        Me.Label317.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label317.Location = New System.Drawing.Point(32, 4)
        Me.Label317.Name = "Label317"
        Me.Label317.Size = New System.Drawing.Size(340, 20)
        Me.Label317.TabIndex = 130
        Me.Label317.Text = "コンバート後に手動設定が必要な項目について"
        Me.Label317.UseCompatibleTextRendering = true
        '
        'Label318
        '
        Me.Label318.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label318.Location = New System.Drawing.Point(36, 27)
        Me.Label318.Name = "Label318"
        Me.Label318.Size = New System.Drawing.Size(340, 50)
        Me.Label318.TabIndex = 131
        Me.Label318.Text = "クライアントやセキュリティ、初期設定などにつていてはユーザにて手動設定頂く必要があります。設定が必要な項目については、仕様書をご確認下さい。"
        Me.Label318.UseCompatibleTextRendering = true
        '
        'pnlJigoCmtSyusi
        '
        Me.pnlJigoCmtSyusi.Controls.Add(Me.Label304)
        Me.pnlJigoCmtSyusi.Controls.Add(Me.Label303)
        Me.pnlJigoCmtSyusi.Controls.Add(Me.Label146)
        Me.pnlJigoCmtSyusi.Location = New System.Drawing.Point(6, 104)
        Me.pnlJigoCmtSyusi.Name = "pnlJigoCmtSyusi"
        Me.pnlJigoCmtSyusi.Size = New System.Drawing.Size(400, 150)
        Me.pnlJigoCmtSyusi.TabIndex = 140
        '
        'Label304
        '
        Me.Label304.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label304.Location = New System.Drawing.Point(32, 4)
        Me.Label304.Name = "Label304"
        Me.Label304.Size = New System.Drawing.Size(340, 20)
        Me.Label304.TabIndex = 120
        Me.Label304.Text = "年間収支系帳票について"
        Me.Label304.UseCompatibleTextRendering = true
        '
        'Label303
        '
        Me.Label303.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label303.Location = New System.Drawing.Point(36, 27)
        Me.Label303.Name = "Label303"
        Me.Label303.Size = New System.Drawing.Size(340, 118)
        Me.Label303.TabIndex = 121
        Me.Label303.Text = "年間収支系の帳票については、コンバートを境にデータが分かれる為、事前にV7または10にてファイル出力頂き、ユーザにて編集(自作)頂く必要があります。(以下、作業例"& _ 
    ")"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"①コンバート前(または並行稼働終了時)に、事前にV7の必要データをファイル出力しておく。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"②必要時(物件別年間送金一覧作成時など)に、10の帳票をファイ"& _ 
    "ル出力し、V7と10のデータをマージ(編集)する。"
        Me.Label303.UseCompatibleTextRendering = true
        '
        'Label146
        '
        Me.Label146.Location = New System.Drawing.Point(11, 4)
        Me.Label146.Name = "Label146"
        Me.Label146.Size = New System.Drawing.Size(20, 20)
        Me.Label146.TabIndex = 135
        Me.Label146.Text = "・"
        Me.Label146.UseCompatibleTextRendering = true
        '
        'pnlJizenMinus
        '
        Me.pnlJizenMinus.Controls.Add(Me.Label358)
        Me.pnlJizenMinus.Controls.Add(Me.chkJizenMinus)
        Me.pnlJizenMinus.Controls.Add(Me.Label102)
        Me.pnlJizenMinus.Controls.Add(Me.Panel2)
        Me.pnlJizenMinus.Location = New System.Drawing.Point(429, 104)
        Me.pnlJizenMinus.Name = "pnlJizenMinus"
        Me.pnlJizenMinus.Size = New System.Drawing.Size(400, 214)
        Me.pnlJizenMinus.TabIndex = 1
        '
        'Label358
        '
        Me.Label358.Location = New System.Drawing.Point(11, 4)
        Me.Label358.Name = "Label358"
        Me.Label358.Size = New System.Drawing.Size(20, 20)
        Me.Label358.TabIndex = 138
        Me.Label358.Text = "・"
        Me.Label358.UseCompatibleTextRendering = true
        '
        'chkJizenMinus
        '
        Me.chkJizenMinus.AutoSize = true
        Me.chkJizenMinus.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenMinus.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenMinus.Name = "chkJizenMinus"
        Me.chkJizenMinus.Size = New System.Drawing.Size(123, 22)
        Me.chkJizenMinus.TabIndex = 0
        Me.chkJizenMinus.Text = "マイナス金額調整"
        Me.chkJizenMinus.UseVisualStyleBackColor = true
        '
        'Label102
        '
        Me.Label102.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label102.Location = New System.Drawing.Point(36, 29)
        Me.Label102.Name = "Label102"
        Me.Label102.Size = New System.Drawing.Size(340, 117)
        Me.Label102.TabIndex = 1
        Me.Label102.Text = "マイナス金額(敷金差額を除く)は移行できません。賃貸革命10で移行されなかったマイナス請求額の調整を行って下さい(既に入金済みのV7データの調整はできない為、10"& _ 
    "で調整を行う必要があります)。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"例) V7のマイナス請求額が「賃料」を対象として登録されている場合、移行された10の「賃料」からマイナス請求額を差し引いた請求"& _ 
    "額に調整する。"
        Me.Label102.UseCompatibleTextRendering = true
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label93)
        Me.Panel2.Controls.Add(Me.Label182)
        Me.Panel2.Controls.Add(Me.Label183)
        Me.Panel2.Controls.Add(Me.lblCntJizenMinus)
        Me.Panel2.Controls.Add(Me.btnListJizenMinus)
        Me.Panel2.Location = New System.Drawing.Point(36, 146)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(340, 60)
        Me.Panel2.TabIndex = 2
        '
        'Label93
        '
        Me.Label93.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label93.Location = New System.Drawing.Point(8, 8)
        Me.Label93.Name = "Label93"
        Me.Label93.Size = New System.Drawing.Size(321, 16)
        Me.Label93.TabIndex = 0
        Me.Label93.Text = "敷金差額以外のマイナス金額のリスト出力"
        Me.Label93.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label93.UseCompatibleTextRendering = true
        '
        'Label182
        '
        Me.Label182.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label182.Location = New System.Drawing.Point(307, 31)
        Me.Label182.Name = "Label182"
        Me.Label182.Size = New System.Drawing.Size(22, 16)
        Me.Label182.TabIndex = 4
        Me.Label182.Text = " )"
        Me.Label182.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label183
        '
        Me.Label183.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label183.Location = New System.Drawing.Point(131, 31)
        Me.Label183.Name = "Label183"
        Me.Label183.Size = New System.Drawing.Size(92, 16)
        Me.Label183.TabIndex = 2
        Me.Label183.Text = "( 抽出件数 = "
        Me.Label183.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntJizenMinus
        '
        Me.lblCntJizenMinus.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntJizenMinus.Location = New System.Drawing.Point(218, 31)
        Me.lblCntJizenMinus.Name = "lblCntJizenMinus"
        Me.lblCntJizenMinus.Size = New System.Drawing.Size(83, 16)
        Me.lblCntJizenMinus.TabIndex = 3
        Me.lblCntJizenMinus.Text = "9,999,999"
        Me.lblCntJizenMinus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnListJizenMinus
        '
        Me.btnListJizenMinus.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListJizenMinus.Image = CType(resources.GetObject("btnListJizenMinus.Image"),System.Drawing.Image)
        Me.btnListJizenMinus.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListJizenMinus.Location = New System.Drawing.Point(8, 27)
        Me.btnListJizenMinus.Name = "btnListJizenMinus"
        Me.btnListJizenMinus.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListJizenMinus.Size = New System.Drawing.Size(100, 25)
        Me.btnListJizenMinus.TabIndex = 1
        Me.btnListJizenMinus.Text = "リスト出力"
        Me.btnListJizenMinus.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListJizenMinus.UseVisualStyleBackColor = true
        '
        'pnlJigoCmtCsvSyuturyoku
        '
        Me.pnlJigoCmtCsvSyuturyoku.Controls.Add(Me.Label147)
        Me.pnlJigoCmtCsvSyuturyoku.Controls.Add(Me.Label334)
        Me.pnlJigoCmtCsvSyuturyoku.Controls.Add(Me.Label333)
        Me.pnlJigoCmtCsvSyuturyoku.Location = New System.Drawing.Point(6, 260)
        Me.pnlJigoCmtCsvSyuturyoku.Name = "pnlJigoCmtCsvSyuturyoku"
        Me.pnlJigoCmtCsvSyuturyoku.Size = New System.Drawing.Size(400, 65)
        Me.pnlJigoCmtCsvSyuturyoku.TabIndex = 138
        '
        'Label147
        '
        Me.Label147.Location = New System.Drawing.Point(11, 4)
        Me.Label147.Name = "Label147"
        Me.Label147.Size = New System.Drawing.Size(20, 20)
        Me.Label147.TabIndex = 136
        Me.Label147.Text = "・"
        Me.Label147.UseCompatibleTextRendering = true
        '
        'Label334
        '
        Me.Label334.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label334.Location = New System.Drawing.Point(32, 4)
        Me.Label334.Name = "Label334"
        Me.Label334.Size = New System.Drawing.Size(330, 20)
        Me.Label334.TabIndex = 132
        Me.Label334.Text = "CSV出力項目設定について"
        Me.Label334.UseCompatibleTextRendering = true
        '
        'Label333
        '
        Me.Label333.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label333.Location = New System.Drawing.Point(36, 27)
        Me.Label333.Name = "Label333"
        Me.Label333.Size = New System.Drawing.Size(340, 31)
        Me.Label333.TabIndex = 133
        Me.Label333.Text = "CSV出力項目定義は移行できません。ユーザにて手動設定頂く必要があります。"
        Me.Label333.UseCompatibleTextRendering = true
        '
        'pnlJigoCmtSyosiki
        '
        Me.pnlJigoCmtSyosiki.Controls.Add(Me.Label148)
        Me.pnlJigoCmtSyosiki.Controls.Add(Me.Label337)
        Me.pnlJigoCmtSyosiki.Controls.Add(Me.Label336)
        Me.pnlJigoCmtSyosiki.Location = New System.Drawing.Point(429, 18)
        Me.pnlJigoCmtSyosiki.Name = "pnlJigoCmtSyosiki"
        Me.pnlJigoCmtSyosiki.Size = New System.Drawing.Size(400, 80)
        Me.pnlJigoCmtSyosiki.TabIndex = 139
        '
        'Label148
        '
        Me.Label148.Location = New System.Drawing.Point(11, 4)
        Me.Label148.Name = "Label148"
        Me.Label148.Size = New System.Drawing.Size(20, 20)
        Me.Label148.TabIndex = 137
        Me.Label148.Text = "・"
        Me.Label148.UseCompatibleTextRendering = true
        '
        'Label337
        '
        Me.Label337.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label337.Location = New System.Drawing.Point(32, 4)
        Me.Label337.Name = "Label337"
        Me.Label337.Size = New System.Drawing.Size(340, 20)
        Me.Label337.TabIndex = 130
        Me.Label337.Text = "書式(帳票)について"
        Me.Label337.UseCompatibleTextRendering = true
        '
        'Label336
        '
        Me.Label336.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label336.Location = New System.Drawing.Point(36, 27)
        Me.Label336.Name = "Label336"
        Me.Label336.Size = New System.Drawing.Size(340, 50)
        Me.Label336.TabIndex = 131
        Me.Label336.Text = "書式の移行はできません。その為、賃貸革命10の書式作成機能で新規に作成、または10で用意されている帳票で代替頂く必要があります。"
        Me.Label336.UseCompatibleTextRendering = true
        '
        'lblJizenPageNum1
        '
        Me.lblJizenPageNum1.AutoSize = true
        Me.lblJizenPageNum1.Location = New System.Drawing.Point(794, 321)
        Me.lblJizenPageNum1.Name = "lblJizenPageNum1"
        Me.lblJizenPageNum1.Size = New System.Drawing.Size(35, 18)
        Me.lblJizenPageNum1.TabIndex = 138
        Me.lblJizenPageNum1.Text = "1 / 2"
        '
        'tabPageJigo2
        '
        Me.tabPageJigo2.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJigo2.Controls.Add(Me.grpJigoCmtTyuui)
        Me.tabPageJigo2.Controls.Add(Me.grpJigoDonyuji)
        Me.tabPageJigo2.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJigo2.Name = "tabPageJigo2"
        Me.tabPageJigo2.Size = New System.Drawing.Size(847, 363)
        Me.tabPageJigo2.TabIndex = 5
        Me.tabPageJigo2.Text = " 事後作業2 "
        '
        'grpJigoCmtTyuui
        '
        Me.grpJigoCmtTyuui.Controls.Add(Me.pnlJigoCmtHeiko)
        Me.grpJigoCmtTyuui.Controls.Add(Me.lblJizenPageNum2)
        Me.grpJigoCmtTyuui.Location = New System.Drawing.Point(436, 15)
        Me.grpJigoCmtTyuui.Name = "grpJigoCmtTyuui"
        Me.grpJigoCmtTyuui.Size = New System.Drawing.Size(405, 342)
        Me.grpJigoCmtTyuui.TabIndex = 138
        Me.grpJigoCmtTyuui.TabStop = false
        Me.grpJigoCmtTyuui.Text = "【注意】"
        '
        'pnlJigoCmtHeiko
        '
        Me.pnlJigoCmtHeiko.Controls.Add(Me.Label23)
        Me.pnlJigoCmtHeiko.Controls.Add(Me.Label115)
        Me.pnlJigoCmtHeiko.Location = New System.Drawing.Point(20, 25)
        Me.pnlJigoCmtHeiko.Name = "pnlJigoCmtHeiko"
        Me.pnlJigoCmtHeiko.Size = New System.Drawing.Size(375, 67)
        Me.pnlJigoCmtHeiko.TabIndex = 136
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(5, 4)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(20, 20)
        Me.Label23.TabIndex = 132
        Me.Label23.Text = "・"
        Me.Label23.UseCompatibleTextRendering = true
        '
        'Label115
        '
        Me.Label115.Location = New System.Drawing.Point(23, 4)
        Me.Label115.Name = "Label115"
        Me.Label115.Size = New System.Drawing.Size(340, 56)
        Me.Label115.TabIndex = 131
        Me.Label115.Text = "並行稼働期間中は、賃貸革命V7と10のデータの同期をとる必要があります。(どちらか片方のデータを更新した場合、もう片方のデータも同様に更新する必要があります)"
        Me.Label115.UseCompatibleTextRendering = true
        '
        'lblJizenPageNum2
        '
        Me.lblJizenPageNum2.AutoSize = true
        Me.lblJizenPageNum2.Location = New System.Drawing.Point(364, 321)
        Me.lblJizenPageNum2.Name = "lblJizenPageNum2"
        Me.lblJizenPageNum2.Size = New System.Drawing.Size(35, 18)
        Me.lblJizenPageNum2.TabIndex = 133
        Me.lblJizenPageNum2.Text = "2 / 2"
        '
        'grpJigoDonyuji
        '
        Me.grpJigoDonyuji.BackColor = System.Drawing.SystemColors.Menu
        Me.grpJigoDonyuji.Controls.Add(Me.pnlJigoCmtSoKotiku)
        Me.grpJigoDonyuji.Controls.Add(Me.pnlJigoCmtNkNyuryoku)
        Me.grpJigoDonyuji.Controls.Add(Me.pnlJigoCmtSqKotiku)
        Me.grpJigoDonyuji.Cursor = System.Windows.Forms.Cursors.Default
        Me.grpJigoDonyuji.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpJigoDonyuji.Location = New System.Drawing.Point(6, 15)
        Me.grpJigoDonyuji.Name = "grpJigoDonyuji"
        Me.grpJigoDonyuji.Size = New System.Drawing.Size(405, 342)
        Me.grpJigoDonyuji.TabIndex = 137
        Me.grpJigoDonyuji.TabStop = false
        Me.grpJigoDonyuji.Text = "【賃貸革命導入時通常作業】"
        '
        'pnlJigoCmtSoKotiku
        '
        Me.pnlJigoCmtSoKotiku.Controls.Add(Me.Label109)
        Me.pnlJigoCmtSoKotiku.Controls.Add(Me.Label108)
        Me.pnlJigoCmtSoKotiku.Controls.Add(Me.Label142)
        Me.pnlJigoCmtSoKotiku.Location = New System.Drawing.Point(20, 173)
        Me.pnlJigoCmtSoKotiku.Name = "pnlJigoCmtSoKotiku"
        Me.pnlJigoCmtSoKotiku.Size = New System.Drawing.Size(375, 65)
        Me.pnlJigoCmtSoKotiku.TabIndex = 140
        '
        'Label109
        '
        Me.Label109.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label109.ForeColor = System.Drawing.Color.Red
        Me.Label109.Location = New System.Drawing.Point(27, 4)
        Me.Label109.Name = "Label109"
        Me.Label109.Size = New System.Drawing.Size(340, 20)
        Me.Label109.TabIndex = 123
        Me.Label109.Text = "送金データ一括構築 (送金処理)"
        Me.Label109.UseCompatibleTextRendering = true
        '
        'Label108
        '
        Me.Label108.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label108.Location = New System.Drawing.Point(23, 28)
        Me.Label108.Name = "Label108"
        Me.Label108.Size = New System.Drawing.Size(340, 36)
        Me.Label108.TabIndex = 124
        Me.Label108.Text = "入金処理後、送金データ一括構築を実施して下さい。データ構築を行わないと送金データが作成されません。"
        Me.Label108.UseCompatibleTextRendering = true
        '
        'Label142
        '
        Me.Label142.ForeColor = System.Drawing.Color.Red
        Me.Label142.Location = New System.Drawing.Point(7, 4)
        Me.Label142.Name = "Label142"
        Me.Label142.Size = New System.Drawing.Size(20, 20)
        Me.Label142.TabIndex = 135
        Me.Label142.Text = "・"
        Me.Label142.UseCompatibleTextRendering = true
        '
        'pnlJigoCmtNkNyuryoku
        '
        Me.pnlJigoCmtNkNyuryoku.Controls.Add(Me.Label99)
        Me.pnlJigoCmtNkNyuryoku.Controls.Add(Me.Label111)
        Me.pnlJigoCmtNkNyuryoku.Controls.Add(Me.Label110)
        Me.pnlJigoCmtNkNyuryoku.Location = New System.Drawing.Point(20, 104)
        Me.pnlJigoCmtNkNyuryoku.Name = "pnlJigoCmtNkNyuryoku"
        Me.pnlJigoCmtNkNyuryoku.Size = New System.Drawing.Size(375, 55)
        Me.pnlJigoCmtNkNyuryoku.TabIndex = 139
        '
        'Label99
        '
        Me.Label99.ForeColor = System.Drawing.Color.Red
        Me.Label99.Location = New System.Drawing.Point(3, 4)
        Me.Label99.Name = "Label99"
        Me.Label99.Size = New System.Drawing.Size(20, 20)
        Me.Label99.TabIndex = 134
        Me.Label99.Text = "・"
        Me.Label99.UseCompatibleTextRendering = true
        '
        'Label111
        '
        Me.Label111.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label111.ForeColor = System.Drawing.Color.Red
        Me.Label111.Location = New System.Drawing.Point(23, 4)
        Me.Label111.Name = "Label111"
        Me.Label111.Size = New System.Drawing.Size(340, 20)
        Me.Label111.TabIndex = 120
        Me.Label111.Text = "入金データ入力作業 (入金処理)"
        Me.Label111.UseCompatibleTextRendering = true
        '
        'Label110
        '
        Me.Label110.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label110.Location = New System.Drawing.Point(23, 28)
        Me.Label110.Name = "Label110"
        Me.Label110.Size = New System.Drawing.Size(340, 20)
        Me.Label110.TabIndex = 121
        Me.Label110.Text = "必要な入金処理を行って下さい。"
        Me.Label110.UseCompatibleTextRendering = true
        '
        'pnlJigoCmtSqKotiku
        '
        Me.pnlJigoCmtSqKotiku.Controls.Add(Me.Label113)
        Me.pnlJigoCmtSqKotiku.Controls.Add(Me.Label112)
        Me.pnlJigoCmtSqKotiku.Controls.Add(Me.Label24)
        Me.pnlJigoCmtSqKotiku.Location = New System.Drawing.Point(20, 25)
        Me.pnlJigoCmtSqKotiku.Name = "pnlJigoCmtSqKotiku"
        Me.pnlJigoCmtSqKotiku.Size = New System.Drawing.Size(375, 65)
        Me.pnlJigoCmtSqKotiku.TabIndex = 136
        '
        'Label113
        '
        Me.Label113.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label113.ForeColor = System.Drawing.Color.Red
        Me.Label113.Location = New System.Drawing.Point(23, 4)
        Me.Label113.Name = "Label113"
        Me.Label113.Size = New System.Drawing.Size(330, 20)
        Me.Label113.TabIndex = 117
        Me.Label113.Text = "請求データ一括構築 (請求データ作成)"
        Me.Label113.UseCompatibleTextRendering = true
        '
        'Label112
        '
        Me.Label112.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label112.Location = New System.Drawing.Point(23, 28)
        Me.Label112.Name = "Label112"
        Me.Label112.Size = New System.Drawing.Size(340, 36)
        Me.Label112.TabIndex = 118
        Me.Label112.Text = "コンバート後、請求データ一括構築を実行して下さい。データ構築を行わないと請求データが作成されません。"
        Me.Label112.UseCompatibleTextRendering = true
        '
        'Label24
        '
        Me.Label24.ForeColor = System.Drawing.Color.Red
        Me.Label24.Location = New System.Drawing.Point(3, 4)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(20, 20)
        Me.Label24.TabIndex = 133
        Me.Label24.Text = "・"
        Me.Label24.UseCompatibleTextRendering = true
        '
        'lblJigoDescription1
        '
        Me.lblJigoDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblJigoDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblJigoDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblJigoDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblJigoDescription1.Name = "lblJigoDescription1"
        Me.lblJigoDescription1.Size = New System.Drawing.Size(820, 80)
        Me.lblJigoDescription1.TabIndex = 6
        Me.lblJigoDescription1.Text = "コンバート後の作業を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"以下の内容を確認し、必要な作業を行って下さい。"
        Me.lblJigoDescription1.UseCompatibleTextRendering = true
        '
        'tabPageHojyo
        '
        Me.tabPageHojyo.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHojyo.Controls.Add(Me.pnlKensyoListPath)
        Me.tabPageHojyo.Controls.Add(Me.lblLine4)
        Me.tabPageHojyo.Controls.Add(Me.lblHidden5)
        Me.tabPageHojyo.Controls.Add(Me.tabCtrlHojyo)
        Me.tabPageHojyo.Controls.Add(Me.lblHojyoDescription1)
        Me.tabPageHojyo.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHojyo.Name = "tabPageHojyo"
        Me.tabPageHojyo.Size = New System.Drawing.Size(912, 549)
        Me.tabPageHojyo.TabIndex = 15
        Me.tabPageHojyo.Text = "検証用"
        '
        'pnlKensyoListPath
        '
        Me.pnlKensyoListPath.Controls.Add(Me.btnKensyoListDirSeach)
        Me.pnlKensyoListPath.Controls.Add(Me.Label257)
        Me.pnlKensyoListPath.Controls.Add(Me.txtKensyoListPath)
        Me.pnlKensyoListPath.Location = New System.Drawing.Point(31, 86)
        Me.pnlKensyoListPath.Name = "pnlKensyoListPath"
        Me.pnlKensyoListPath.Size = New System.Drawing.Size(420, 60)
        Me.pnlKensyoListPath.TabIndex = 48
        '
        'btnKensyoListDirSeach
        '
        Me.btnKensyoListDirSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnKensyoListDirSeach.Location = New System.Drawing.Point(383, 27)
        Me.btnKensyoListDirSeach.Name = "btnKensyoListDirSeach"
        Me.btnKensyoListDirSeach.Size = New System.Drawing.Size(30, 25)
        Me.btnKensyoListDirSeach.TabIndex = 2
        Me.btnKensyoListDirSeach.Text = "..."
        Me.btnKensyoListDirSeach.UseVisualStyleBackColor = true
        '
        'Label257
        '
        Me.Label257.AutoSize = true
        Me.Label257.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label257.Location = New System.Drawing.Point(5, 5)
        Me.Label257.Name = "Label257"
        Me.Label257.Size = New System.Drawing.Size(116, 18)
        Me.Label257.TabIndex = 3
        Me.Label257.Text = "出力ファイル格納先"
        '
        'txtKensyoListPath
        '
        Me.txtKensyoListPath.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtKensyoListPath.Location = New System.Drawing.Point(23, 28)
        Me.txtKensyoListPath.Name = "txtKensyoListPath"
        Me.txtKensyoListPath.Size = New System.Drawing.Size(354, 24)
        Me.txtKensyoListPath.TabIndex = 1
        '
        'lblLine4
        '
        Me.lblLine4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLine4.Location = New System.Drawing.Point(31, 176)
        Me.lblLine4.Name = "lblLine4"
        Me.lblLine4.Size = New System.Drawing.Size(853, 2)
        Me.lblLine4.TabIndex = 16
        '
        'lblHidden5
        '
        Me.lblHidden5.BackColor = System.Drawing.SystemColors.Menu
        Me.lblHidden5.Location = New System.Drawing.Point(8, 150)
        Me.lblHidden5.Name = "lblHidden5"
        Me.lblHidden5.Size = New System.Drawing.Size(16, 26)
        Me.lblHidden5.TabIndex = 15
        Me.lblHidden5.Text = "　"
        '
        'tabCtrlHojyo
        '
        Me.tabCtrlHojyo.Controls.Add(Me.tabPageHojyo1)
        Me.tabCtrlHojyo.Controls.Add(Me.tabPageHojyo2)
        Me.tabCtrlHojyo.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
        Me.tabCtrlHojyo.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.tabCtrlHojyo.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.tabCtrlHojyo.Location = New System.Drawing.Point(30, 150)
        Me.tabCtrlHojyo.Name = "tabCtrlHojyo"
        Me.tabCtrlHojyo.SelectedIndex = 0
        Me.tabCtrlHojyo.Size = New System.Drawing.Size(855, 394)
        Me.tabCtrlHojyo.TabIndex = 12
        Me.tabCtrlHojyo.TabStop = false
        '
        'tabPageHojyo1
        '
        Me.tabPageHojyo1.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHojyo1.Controls.Add(Me.GroupBox27)
        Me.tabPageHojyo1.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHojyo1.Name = "tabPageHojyo1"
        Me.tabPageHojyo1.Size = New System.Drawing.Size(847, 363)
        Me.tabPageHojyo1.TabIndex = 4
        Me.tabPageHojyo1.Text = " 補助機能1 "
        '
        'GroupBox27
        '
        Me.GroupBox27.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox27.Controls.Add(Me.pnlKiOpKys)
        Me.GroupBox27.Controls.Add(Me.pnlKiOpOw)
        Me.GroupBox27.Controls.Add(Me.pnlRelRename)
        Me.GroupBox27.Controls.Add(Me.Label402)
        Me.GroupBox27.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox27.Location = New System.Drawing.Point(6, 15)
        Me.GroupBox27.Name = "GroupBox27"
        Me.GroupBox27.Size = New System.Drawing.Size(835, 342)
        Me.GroupBox27.TabIndex = 0
        Me.GroupBox27.TabStop = false
        Me.GroupBox27.Text = "【補助機能】"
        '
        'pnlKiOpKys
        '
        Me.pnlKiOpKys.Controls.Add(Me.Label291)
        Me.pnlKiOpKys.Controls.Add(Me.chkKiOpKys)
        Me.pnlKiOpKys.Controls.Add(Me.Label184)
        Me.pnlKiOpKys.Controls.Add(Me.Panel18)
        Me.pnlKiOpKys.Location = New System.Drawing.Point(429, 167)
        Me.pnlKiOpKys.Name = "pnlKiOpKys"
        Me.pnlKiOpKys.Size = New System.Drawing.Size(400, 150)
        Me.pnlKiOpKys.TabIndex = 2
        '
        'Label291
        '
        Me.Label291.Location = New System.Drawing.Point(11, 5)
        Me.Label291.Name = "Label291"
        Me.Label291.Size = New System.Drawing.Size(20, 20)
        Me.Label291.TabIndex = 136
        Me.Label291.Text = "・"
        Me.Label291.UseCompatibleTextRendering = true
        '
        'chkKiOpKys
        '
        Me.chkKiOpKys.AutoSize = true
        Me.chkKiOpKys.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiOpKys.Location = New System.Drawing.Point(16, 4)
        Me.chkKiOpKys.Name = "chkKiOpKys"
        Me.chkKiOpKys.Size = New System.Drawing.Size(183, 22)
        Me.chkKiOpKys.TabIndex = 0
        Me.chkKiOpKys.Text = "未使用の契約者データの削除"
        Me.chkKiOpKys.UseVisualStyleBackColor = true
        '
        'Label184
        '
        Me.Label184.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label184.Location = New System.Drawing.Point(36, 29)
        Me.Label184.Name = "Label184"
        Me.Label184.Size = New System.Drawing.Size(340, 20)
        Me.Label184.TabIndex = 1
        Me.Label184.Text = "どの情報にも紐付かない契約者データを削除します。"
        Me.Label184.UseCompatibleTextRendering = true
        '
        'Panel18
        '
        Me.Panel18.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel18.Controls.Add(Me.Label206)
        Me.Panel18.Controls.Add(Me.btnListKiOpKys)
        Me.Panel18.Controls.Add(Me.btnKiOpKys)
        Me.Panel18.Controls.Add(Me.Label187)
        Me.Panel18.Controls.Add(Me.Label191)
        Me.Panel18.Controls.Add(Me.Label193)
        Me.Panel18.Controls.Add(Me.lblCntKiOpKys)
        Me.Panel18.Location = New System.Drawing.Point(36, 52)
        Me.Panel18.Name = "Panel18"
        Me.Panel18.Size = New System.Drawing.Size(340, 90)
        Me.Panel18.TabIndex = 2
        '
        'Label206
        '
        Me.Label206.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label206.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label206.Location = New System.Drawing.Point(122, 51)
        Me.Label206.Name = "Label206"
        Me.Label206.Size = New System.Drawing.Size(228, 32)
        Me.Label206.TabIndex = 6
        Me.Label206.Text = "※削除後は元に戻すことはできません。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　十分注意して実行して下さい。"
        '
        'btnListKiOpKys
        '
        Me.btnListKiOpKys.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListKiOpKys.Image = CType(resources.GetObject("btnListKiOpKys.Image"),System.Drawing.Image)
        Me.btnListKiOpKys.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListKiOpKys.Location = New System.Drawing.Point(8, 27)
        Me.btnListKiOpKys.Name = "btnListKiOpKys"
        Me.btnListKiOpKys.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListKiOpKys.Size = New System.Drawing.Size(100, 25)
        Me.btnListKiOpKys.TabIndex = 1
        Me.btnListKiOpKys.Text = "リスト出力"
        Me.btnListKiOpKys.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListKiOpKys.UseVisualStyleBackColor = true
        '
        'btnKiOpKys
        '
        Me.btnKiOpKys.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnKiOpKys.Image = CType(resources.GetObject("btnKiOpKys.Image"),System.Drawing.Image)
        Me.btnKiOpKys.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnKiOpKys.Location = New System.Drawing.Point(8, 58)
        Me.btnKiOpKys.Name = "btnKiOpKys"
        Me.btnKiOpKys.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnKiOpKys.Size = New System.Drawing.Size(100, 25)
        Me.btnKiOpKys.TabIndex = 5
        Me.btnKiOpKys.Text = "データ削除"
        Me.btnKiOpKys.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnKiOpKys.UseVisualStyleBackColor = true
        '
        'Label187
        '
        Me.Label187.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label187.Location = New System.Drawing.Point(8, 8)
        Me.Label187.Name = "Label187"
        Me.Label187.Size = New System.Drawing.Size(321, 16)
        Me.Label187.TabIndex = 0
        Me.Label187.Text = "未使用と思われる契約者のデータのリスト出力"
        Me.Label187.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label187.UseCompatibleTextRendering = true
        '
        'Label191
        '
        Me.Label191.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label191.Location = New System.Drawing.Point(307, 31)
        Me.Label191.Name = "Label191"
        Me.Label191.Size = New System.Drawing.Size(22, 16)
        Me.Label191.TabIndex = 4
        Me.Label191.Text = " )"
        Me.Label191.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label193
        '
        Me.Label193.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label193.Location = New System.Drawing.Point(131, 31)
        Me.Label193.Name = "Label193"
        Me.Label193.Size = New System.Drawing.Size(92, 16)
        Me.Label193.TabIndex = 2
        Me.Label193.Text = "( 対象件数 = "
        Me.Label193.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntKiOpKys
        '
        Me.lblCntKiOpKys.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntKiOpKys.Location = New System.Drawing.Point(218, 31)
        Me.lblCntKiOpKys.Name = "lblCntKiOpKys"
        Me.lblCntKiOpKys.Size = New System.Drawing.Size(83, 16)
        Me.lblCntKiOpKys.TabIndex = 3
        Me.lblCntKiOpKys.Text = "9,999,999"
        Me.lblCntKiOpKys.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiOpOw
        '
        Me.pnlKiOpOw.Controls.Add(Me.Label287)
        Me.pnlKiOpOw.Controls.Add(Me.chkKiOpOw)
        Me.pnlKiOpOw.Controls.Add(Me.Label169)
        Me.pnlKiOpOw.Controls.Add(Me.Panel16)
        Me.pnlKiOpOw.Location = New System.Drawing.Point(429, 18)
        Me.pnlKiOpOw.Name = "pnlKiOpOw"
        Me.pnlKiOpOw.Size = New System.Drawing.Size(400, 150)
        Me.pnlKiOpOw.TabIndex = 1
        '
        'Label287
        '
        Me.Label287.Location = New System.Drawing.Point(11, 5)
        Me.Label287.Name = "Label287"
        Me.Label287.Size = New System.Drawing.Size(20, 20)
        Me.Label287.TabIndex = 136
        Me.Label287.Text = "・"
        Me.Label287.UseCompatibleTextRendering = true
        '
        'chkKiOpOw
        '
        Me.chkKiOpOw.AutoSize = true
        Me.chkKiOpOw.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiOpOw.Location = New System.Drawing.Point(16, 4)
        Me.chkKiOpOw.Name = "chkKiOpOw"
        Me.chkKiOpOw.Size = New System.Drawing.Size(171, 22)
        Me.chkKiOpOw.TabIndex = 0
        Me.chkKiOpOw.Text = "未使用の家主データの削除"
        Me.chkKiOpOw.UseVisualStyleBackColor = true
        '
        'Label169
        '
        Me.Label169.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label169.Location = New System.Drawing.Point(36, 29)
        Me.Label169.Name = "Label169"
        Me.Label169.Size = New System.Drawing.Size(340, 20)
        Me.Label169.TabIndex = 1
        Me.Label169.Text = "どの情報にも紐付かない家主データを削除します。"
        Me.Label169.UseCompatibleTextRendering = true
        '
        'Panel16
        '
        Me.Panel16.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel16.Controls.Add(Me.btnKiOpOw)
        Me.Panel16.Controls.Add(Me.btnListKiOpOw)
        Me.Panel16.Controls.Add(Me.Label205)
        Me.Panel16.Controls.Add(Me.Label172)
        Me.Panel16.Controls.Add(Me.Label177)
        Me.Panel16.Controls.Add(Me.Label178)
        Me.Panel16.Controls.Add(Me.lblCntKiOpOw)
        Me.Panel16.Location = New System.Drawing.Point(36, 52)
        Me.Panel16.Name = "Panel16"
        Me.Panel16.Size = New System.Drawing.Size(340, 91)
        Me.Panel16.TabIndex = 2
        '
        'btnKiOpOw
        '
        Me.btnKiOpOw.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnKiOpOw.Image = CType(resources.GetObject("btnKiOpOw.Image"),System.Drawing.Image)
        Me.btnKiOpOw.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnKiOpOw.Location = New System.Drawing.Point(8, 58)
        Me.btnKiOpOw.Name = "btnKiOpOw"
        Me.btnKiOpOw.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnKiOpOw.Size = New System.Drawing.Size(100, 25)
        Me.btnKiOpOw.TabIndex = 5
        Me.btnKiOpOw.Text = "データ削除"
        Me.btnKiOpOw.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnKiOpOw.UseVisualStyleBackColor = true
        '
        'btnListKiOpOw
        '
        Me.btnListKiOpOw.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListKiOpOw.Image = CType(resources.GetObject("btnListKiOpOw.Image"),System.Drawing.Image)
        Me.btnListKiOpOw.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListKiOpOw.Location = New System.Drawing.Point(8, 27)
        Me.btnListKiOpOw.Name = "btnListKiOpOw"
        Me.btnListKiOpOw.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListKiOpOw.Size = New System.Drawing.Size(100, 25)
        Me.btnListKiOpOw.TabIndex = 1
        Me.btnListKiOpOw.Text = "リスト出力"
        Me.btnListKiOpOw.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListKiOpOw.UseVisualStyleBackColor = true
        '
        'Label205
        '
        Me.Label205.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label205.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label205.Location = New System.Drawing.Point(122, 52)
        Me.Label205.Name = "Label205"
        Me.Label205.Size = New System.Drawing.Size(228, 32)
        Me.Label205.TabIndex = 6
        Me.Label205.Text = "※削除後は元に戻すことはできません。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　十分注意して実行して下さい。"
        '
        'Label172
        '
        Me.Label172.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label172.Location = New System.Drawing.Point(8, 8)
        Me.Label172.Name = "Label172"
        Me.Label172.Size = New System.Drawing.Size(321, 16)
        Me.Label172.TabIndex = 0
        Me.Label172.Text = "未使用と思われる家主のデータのリスト出力"
        Me.Label172.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label172.UseCompatibleTextRendering = true
        '
        'Label177
        '
        Me.Label177.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label177.Location = New System.Drawing.Point(307, 31)
        Me.Label177.Name = "Label177"
        Me.Label177.Size = New System.Drawing.Size(22, 16)
        Me.Label177.TabIndex = 4
        Me.Label177.Text = " )"
        Me.Label177.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label178
        '
        Me.Label178.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label178.Location = New System.Drawing.Point(131, 31)
        Me.Label178.Name = "Label178"
        Me.Label178.Size = New System.Drawing.Size(92, 16)
        Me.Label178.TabIndex = 2
        Me.Label178.Text = "( 対象件数 = "
        Me.Label178.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntKiOpOw
        '
        Me.lblCntKiOpOw.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntKiOpOw.Location = New System.Drawing.Point(218, 31)
        Me.lblCntKiOpOw.Name = "lblCntKiOpOw"
        Me.lblCntKiOpOw.Size = New System.Drawing.Size(83, 16)
        Me.lblCntKiOpOw.TabIndex = 3
        Me.lblCntKiOpOw.Text = "9,999,999"
        Me.lblCntKiOpOw.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlRelRename
        '
        Me.pnlRelRename.Controls.Add(Me.Label274)
        Me.pnlRelRename.Controls.Add(Me.chkRelRename)
        Me.pnlRelRename.Controls.Add(Me.Panel25)
        Me.pnlRelRename.Controls.Add(Me.Label401)
        Me.pnlRelRename.Location = New System.Drawing.Point(6, 18)
        Me.pnlRelRename.Name = "pnlRelRename"
        Me.pnlRelRename.Size = New System.Drawing.Size(400, 310)
        Me.pnlRelRename.TabIndex = 0
        '
        'Label274
        '
        Me.Label274.Location = New System.Drawing.Point(11, 5)
        Me.Label274.Name = "Label274"
        Me.Label274.Size = New System.Drawing.Size(20, 20)
        Me.Label274.TabIndex = 135
        Me.Label274.Text = "・"
        Me.Label274.UseCompatibleTextRendering = true
        '
        'chkRelRename
        '
        Me.chkRelRename.AutoSize = true
        Me.chkRelRename.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkRelRename.Location = New System.Drawing.Point(16, 4)
        Me.chkRelRename.Name = "chkRelRename"
        Me.chkRelRename.Size = New System.Drawing.Size(171, 22)
        Me.chkRelRename.TabIndex = 0
        Me.chkRelRename.Text = "関連ファイルパス一括置換"
        Me.chkRelRename.UseVisualStyleBackColor = true
        '
        'Panel25
        '
        Me.Panel25.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel25.Controls.Add(Me.Label242)
        Me.Panel25.Controls.Add(Me.Label231)
        Me.Panel25.Controls.Add(Me.Label221)
        Me.Panel25.Controls.Add(Me.lblRelRenameSzen)
        Me.Panel25.Controls.Add(Me.Label229)
        Me.Panel25.Controls.Add(Me.Label217)
        Me.Panel25.Controls.Add(Me.Label215)
        Me.Panel25.Controls.Add(Me.Label212)
        Me.Panel25.Controls.Add(Me.chkRelRenameSzen)
        Me.Panel25.Controls.Add(Me.chkRelRenameClaim)
        Me.Panel25.Controls.Add(Me.btnRNRelRename)
        Me.Panel25.Controls.Add(Me.Label394)
        Me.Panel25.Controls.Add(Me.lblRelRenameClaim)
        Me.Panel25.Controls.Add(Me.Label397)
        Me.Panel25.Controls.Add(Me.txtAfRelRename)
        Me.Panel25.Controls.Add(Me.Label398)
        Me.Panel25.Controls.Add(Me.txtBfRelRename)
        Me.Panel25.Controls.Add(Me.Label399)
        Me.Panel25.Controls.Add(Me.Label400)
        Me.Panel25.Location = New System.Drawing.Point(36, 65)
        Me.Panel25.Name = "Panel25"
        Me.Panel25.Size = New System.Drawing.Size(340, 240)
        Me.Panel25.TabIndex = 2
        '
        'Label242
        '
        Me.Label242.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label242.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label242.Location = New System.Drawing.Point(43, 161)
        Me.Label242.Name = "Label242"
        Me.Label242.Size = New System.Drawing.Size(290, 15)
        Me.Label242.TabIndex = 15
        Me.Label242.Text = "※入力されたパスは存在しない可能性があります。"
        '
        'Label231
        '
        Me.Label231.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label231.Location = New System.Drawing.Point(9, 185)
        Me.Label231.Name = "Label231"
        Me.Label231.Size = New System.Drawing.Size(300, 16)
        Me.Label231.TabIndex = 16
        Me.Label231.Text = "③一括変換ボタンを押して下さい。"
        '
        'Label221
        '
        Me.Label221.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label221.Location = New System.Drawing.Point(104, 113)
        Me.Label221.Name = "Label221"
        Me.Label221.Size = New System.Drawing.Size(120, 16)
        Me.Label221.TabIndex = 10
        Me.Label221.Text = "( 修繕置換件数 ="
        Me.Label221.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRelRenameSzen
        '
        Me.lblRelRenameSzen.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblRelRenameSzen.Location = New System.Drawing.Point(218, 114)
        Me.lblRelRenameSzen.Name = "lblRelRenameSzen"
        Me.lblRelRenameSzen.Size = New System.Drawing.Size(83, 16)
        Me.lblRelRenameSzen.TabIndex = 11
        Me.lblRelRenameSzen.Text = "9,999,999"
        Me.lblRelRenameSzen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label229
        '
        Me.Label229.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label229.Location = New System.Drawing.Point(307, 114)
        Me.Label229.Name = "Label229"
        Me.Label229.Size = New System.Drawing.Size(22, 16)
        Me.Label229.TabIndex = 12
        Me.Label229.Text = " )"
        Me.Label229.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label217
        '
        Me.Label217.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label217.Location = New System.Drawing.Point(104, 97)
        Me.Label217.Name = "Label217"
        Me.Label217.Size = New System.Drawing.Size(120, 16)
        Me.Label217.TabIndex = 7
        Me.Label217.Text = "( クレーム置換件数 = "
        Me.Label217.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label215
        '
        Me.Label215.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label215.Location = New System.Drawing.Point(9, 51)
        Me.Label215.Name = "Label215"
        Me.Label215.Size = New System.Drawing.Size(300, 16)
        Me.Label215.TabIndex = 3
        Me.Label215.Text = "②置換前と置換後の文字列を入力して下さい。"
        '
        'Label212
        '
        Me.Label212.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label212.Location = New System.Drawing.Point(8, 8)
        Me.Label212.Name = "Label212"
        Me.Label212.Size = New System.Drawing.Size(300, 16)
        Me.Label212.TabIndex = 0
        Me.Label212.Text = "①置換する関連ファイル情報にチェックを入れて下さい。"
        '
        'chkRelRenameSzen
        '
        Me.chkRelRenameSzen.AutoSize = true
        Me.chkRelRenameSzen.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkRelRenameSzen.Location = New System.Drawing.Point(197, 27)
        Me.chkRelRenameSzen.Name = "chkRelRenameSzen"
        Me.chkRelRenameSzen.Size = New System.Drawing.Size(49, 21)
        Me.chkRelRenameSzen.TabIndex = 2
        Me.chkRelRenameSzen.Text = "修繕"
        Me.chkRelRenameSzen.UseVisualStyleBackColor = true
        '
        'chkRelRenameClaim
        '
        Me.chkRelRenameClaim.AutoSize = true
        Me.chkRelRenameClaim.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkRelRenameClaim.Location = New System.Drawing.Point(57, 27)
        Me.chkRelRenameClaim.Name = "chkRelRenameClaim"
        Me.chkRelRenameClaim.Size = New System.Drawing.Size(71, 21)
        Me.chkRelRenameClaim.TabIndex = 1
        Me.chkRelRenameClaim.Text = "クレーム"
        Me.chkRelRenameClaim.UseVisualStyleBackColor = true
        '
        'btnRNRelRename
        '
        Me.btnRNRelRename.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnRNRelRename.Image = CType(resources.GetObject("btnRNRelRename.Image"),System.Drawing.Image)
        Me.btnRNRelRename.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRNRelRename.Location = New System.Drawing.Point(25, 208)
        Me.btnRNRelRename.Name = "btnRNRelRename"
        Me.btnRNRelRename.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnRNRelRename.Size = New System.Drawing.Size(86, 25)
        Me.btnRNRelRename.TabIndex = 17
        Me.btnRNRelRename.Text = "一括変換"
        Me.btnRNRelRename.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnRNRelRename.UseVisualStyleBackColor = true
        '
        'Label394
        '
        Me.Label394.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label394.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label394.Location = New System.Drawing.Point(122, 203)
        Me.Label394.Name = "Label394"
        Me.Label394.Size = New System.Drawing.Size(211, 36)
        Me.Label394.TabIndex = 18
        Me.Label394.Text = "※変換後は元に戻すことはできません。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　十分注意して実行して下さい。"
        '
        'lblRelRenameClaim
        '
        Me.lblRelRenameClaim.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblRelRenameClaim.Location = New System.Drawing.Point(218, 98)
        Me.lblRelRenameClaim.Name = "lblRelRenameClaim"
        Me.lblRelRenameClaim.Size = New System.Drawing.Size(83, 16)
        Me.lblRelRenameClaim.TabIndex = 8
        Me.lblRelRenameClaim.Text = "9,999,999"
        Me.lblRelRenameClaim.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label397
        '
        Me.Label397.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label397.Location = New System.Drawing.Point(41, 103)
        Me.Label397.Name = "Label397"
        Me.Label397.Size = New System.Drawing.Size(20, 16)
        Me.Label397.TabIndex = 6
        Me.Label397.Text = "↓"
        '
        'txtAfRelRename
        '
        Me.txtAfRelRename.Location = New System.Drawing.Point(79, 133)
        Me.txtAfRelRename.Name = "txtAfRelRename"
        Me.txtAfRelRename.Size = New System.Drawing.Size(200, 25)
        Me.txtAfRelRename.TabIndex = 14
        '
        'Label398
        '
        Me.Label398.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label398.Location = New System.Drawing.Point(31, 137)
        Me.Label398.Name = "Label398"
        Me.Label398.Size = New System.Drawing.Size(60, 16)
        Me.Label398.TabIndex = 13
        Me.Label398.Text = "置換後"
        '
        'txtBfRelRename
        '
        Me.txtBfRelRename.Location = New System.Drawing.Point(79, 70)
        Me.txtBfRelRename.Name = "txtBfRelRename"
        Me.txtBfRelRename.Size = New System.Drawing.Size(200, 25)
        Me.txtBfRelRename.TabIndex = 5
        '
        'Label399
        '
        Me.Label399.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label399.Location = New System.Drawing.Point(31, 74)
        Me.Label399.Name = "Label399"
        Me.Label399.Size = New System.Drawing.Size(60, 16)
        Me.Label399.TabIndex = 4
        Me.Label399.Text = "置換前"
        '
        'Label400
        '
        Me.Label400.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label400.Location = New System.Drawing.Point(307, 98)
        Me.Label400.Name = "Label400"
        Me.Label400.Size = New System.Drawing.Size(22, 16)
        Me.Label400.TabIndex = 9
        Me.Label400.Text = " )"
        Me.Label400.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label401
        '
        Me.Label401.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label401.Location = New System.Drawing.Point(36, 29)
        Me.Label401.Name = "Label401"
        Me.Label401.Size = New System.Drawing.Size(340, 33)
        Me.Label401.TabIndex = 1
        Me.Label401.Text = "クレーム・修繕の関連ファイルパスの一括置換を行います。(登録データの文字位置の先頭を基準に置換します)"
        Me.Label401.UseCompatibleTextRendering = true
        '
        'Label402
        '
        Me.Label402.AutoSize = true
        Me.Label402.Location = New System.Drawing.Point(794, 321)
        Me.Label402.Name = "Label402"
        Me.Label402.Size = New System.Drawing.Size(35, 18)
        Me.Label402.TabIndex = 11
        Me.Label402.Text = "1 / 2"
        '
        'tabPageHojyo2
        '
        Me.tabPageHojyo2.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHojyo2.Controls.Add(Me.GroupBox25)
        Me.tabPageHojyo2.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHojyo2.Name = "tabPageHojyo2"
        Me.tabPageHojyo2.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHojyo2.Size = New System.Drawing.Size(847, 363)
        Me.tabPageHojyo2.TabIndex = 3
        Me.tabPageHojyo2.Text = " 補助機能2 "
        '
        'GroupBox25
        '
        Me.GroupBox25.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox25.Controls.Add(Me.pnlRelationSet)
        Me.GroupBox25.Controls.Add(Me.pnlSzenKysSorit)
        Me.GroupBox25.Controls.Add(Me.pnlKojyosh)
        Me.GroupBox25.Controls.Add(Me.pnlBunkatumisyu)
        Me.GroupBox25.Controls.Add(Me.Label208)
        Me.GroupBox25.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox25.Location = New System.Drawing.Point(6, 15)
        Me.GroupBox25.Name = "GroupBox25"
        Me.GroupBox25.Size = New System.Drawing.Size(835, 342)
        Me.GroupBox25.TabIndex = 0
        Me.GroupBox25.TabStop = false
        Me.GroupBox25.Text = "【検証用】"
        '
        'pnlRelationSet
        '
        Me.pnlRelationSet.Controls.Add(Me.Label307)
        Me.pnlRelationSet.Controls.Add(Me.chkRelationSet)
        Me.pnlRelationSet.Controls.Add(Me.Label225)
        Me.pnlRelationSet.Controls.Add(Me.Panel17)
        Me.pnlRelationSet.Location = New System.Drawing.Point(429, 204)
        Me.pnlRelationSet.Name = "pnlRelationSet"
        Me.pnlRelationSet.Size = New System.Drawing.Size(400, 115)
        Me.pnlRelationSet.TabIndex = 3
        '
        'Label307
        '
        Me.Label307.Location = New System.Drawing.Point(10, 5)
        Me.Label307.Name = "Label307"
        Me.Label307.Size = New System.Drawing.Size(20, 20)
        Me.Label307.TabIndex = 136
        Me.Label307.Text = "・"
        Me.Label307.UseCompatibleTextRendering = true
        '
        'chkRelationSet
        '
        Me.chkRelationSet.AutoSize = true
        Me.chkRelationSet.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkRelationSet.Location = New System.Drawing.Point(16, 4)
        Me.chkRelationSet.Name = "chkRelationSet"
        Me.chkRelationSet.Size = New System.Drawing.Size(320, 22)
        Me.chkRelationSet.TabIndex = 0
        Me.chkRelationSet.Text = "紐付設定画面で設定したV7と10の紐付設定内容の確認"
        Me.chkRelationSet.UseVisualStyleBackColor = true
        '
        'Label225
        '
        Me.Label225.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label225.Location = New System.Drawing.Point(36, 29)
        Me.Label225.Name = "Label225"
        Me.Label225.Size = New System.Drawing.Size(340, 15)
        Me.Label225.TabIndex = 1
        Me.Label225.Text = "コンバート後の検証用リストを出力します。"
        Me.Label225.UseCompatibleTextRendering = true
        '
        'Panel17
        '
        Me.Panel17.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel17.Controls.Add(Me.btnListRelationSet)
        Me.Panel17.Controls.Add(Me.Label234)
        Me.Panel17.Controls.Add(Me.Label237)
        Me.Panel17.Controls.Add(Me.Label239)
        Me.Panel17.Controls.Add(Me.lblCntRelationSet)
        Me.Panel17.Location = New System.Drawing.Point(36, 48)
        Me.Panel17.Name = "Panel17"
        Me.Panel17.Size = New System.Drawing.Size(340, 60)
        Me.Panel17.TabIndex = 2
        '
        'btnListRelationSet
        '
        Me.btnListRelationSet.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListRelationSet.Image = CType(resources.GetObject("btnListRelationSet.Image"),System.Drawing.Image)
        Me.btnListRelationSet.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListRelationSet.Location = New System.Drawing.Point(8, 27)
        Me.btnListRelationSet.Name = "btnListRelationSet"
        Me.btnListRelationSet.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListRelationSet.Size = New System.Drawing.Size(100, 25)
        Me.btnListRelationSet.TabIndex = 1
        Me.btnListRelationSet.Text = "リスト出力"
        Me.btnListRelationSet.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListRelationSet.UseVisualStyleBackColor = true
        '
        'Label234
        '
        Me.Label234.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label234.Location = New System.Drawing.Point(8, 8)
        Me.Label234.Name = "Label234"
        Me.Label234.Size = New System.Drawing.Size(321, 16)
        Me.Label234.TabIndex = 0
        Me.Label234.Text = "紐付設定内容のリスト出力"
        Me.Label234.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label234.UseCompatibleTextRendering = true
        '
        'Label237
        '
        Me.Label237.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label237.Location = New System.Drawing.Point(307, 31)
        Me.Label237.Name = "Label237"
        Me.Label237.Size = New System.Drawing.Size(22, 16)
        Me.Label237.TabIndex = 4
        Me.Label237.Text = " )"
        Me.Label237.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label239
        '
        Me.Label239.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label239.Location = New System.Drawing.Point(131, 31)
        Me.Label239.Name = "Label239"
        Me.Label239.Size = New System.Drawing.Size(92, 16)
        Me.Label239.TabIndex = 2
        Me.Label239.Text = "( 対応紐付数 = "
        Me.Label239.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntRelationSet
        '
        Me.lblCntRelationSet.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntRelationSet.Location = New System.Drawing.Point(218, 31)
        Me.lblCntRelationSet.Name = "lblCntRelationSet"
        Me.lblCntRelationSet.Size = New System.Drawing.Size(83, 16)
        Me.lblCntRelationSet.TabIndex = 3
        Me.lblCntRelationSet.Text = "9,999,999"
        Me.lblCntRelationSet.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlSzenKysSorit
        '
        Me.pnlSzenKysSorit.Controls.Add(Me.Label297)
        Me.pnlSzenKysSorit.Controls.Add(Me.chkSzenKysSorit)
        Me.pnlSzenKysSorit.Controls.Add(Me.Label194)
        Me.pnlSzenKysSorit.Controls.Add(Me.Panel1)
        Me.pnlSzenKysSorit.Location = New System.Drawing.Point(429, 18)
        Me.pnlSzenKysSorit.Name = "pnlSzenKysSorit"
        Me.pnlSzenKysSorit.Size = New System.Drawing.Size(400, 180)
        Me.pnlSzenKysSorit.TabIndex = 2
        '
        'Label297
        '
        Me.Label297.Location = New System.Drawing.Point(10, 5)
        Me.Label297.Name = "Label297"
        Me.Label297.Size = New System.Drawing.Size(20, 20)
        Me.Label297.TabIndex = 136
        Me.Label297.Text = "・"
        Me.Label297.UseCompatibleTextRendering = true
        '
        'chkSzenKysSorit
        '
        Me.chkSzenKysSorit.AutoSize = true
        Me.chkSzenKysSorit.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkSzenKysSorit.Location = New System.Drawing.Point(16, 4)
        Me.chkSzenKysSorit.Name = "chkSzenKysSorit"
        Me.chkSzenKysSorit.Size = New System.Drawing.Size(231, 22)
        Me.chkSzenKysSorit.TabIndex = 0
        Me.chkSzenKysSorit.Text = "修繕項目毎契約者送金率データの確認"
        Me.chkSzenKysSorit.UseVisualStyleBackColor = true
        '
        'Label194
        '
        Me.Label194.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label194.Location = New System.Drawing.Point(36, 29)
        Me.Label194.Name = "Label194"
        Me.Label194.Size = New System.Drawing.Size(340, 85)
        Me.Label194.TabIndex = 1
        Me.Label194.Text = "修繕の契約者負担分の送金率について、V7では修繕項目毎に設定できるが、10ではできません。その為、契約者負担額の各項目の送金額の和(Σ(負担額*送金率))と契約者"& _ 
    "の負担額の和(Σ負担額)より送金率を算出して移行します。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"コンバート後の検証用リストを出力します。"
        Me.Label194.UseCompatibleTextRendering = true
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.btnListSzenKysSorit)
        Me.Panel1.Controls.Add(Me.Label92)
        Me.Panel1.Controls.Add(Me.Label95)
        Me.Panel1.Controls.Add(Me.Label175)
        Me.Panel1.Controls.Add(Me.lblCntSzenKysSorit)
        Me.Panel1.Location = New System.Drawing.Point(36, 116)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(340, 60)
        Me.Panel1.TabIndex = 2
        '
        'btnListSzenKysSorit
        '
        Me.btnListSzenKysSorit.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListSzenKysSorit.Image = CType(resources.GetObject("btnListSzenKysSorit.Image"),System.Drawing.Image)
        Me.btnListSzenKysSorit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListSzenKysSorit.Location = New System.Drawing.Point(8, 27)
        Me.btnListSzenKysSorit.Name = "btnListSzenKysSorit"
        Me.btnListSzenKysSorit.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListSzenKysSorit.Size = New System.Drawing.Size(100, 25)
        Me.btnListSzenKysSorit.TabIndex = 1
        Me.btnListSzenKysSorit.Text = "リスト出力"
        Me.btnListSzenKysSorit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListSzenKysSorit.UseVisualStyleBackColor = true
        '
        'Label92
        '
        Me.Label92.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label92.Location = New System.Drawing.Point(8, 8)
        Me.Label92.Name = "Label92"
        Me.Label92.Size = New System.Drawing.Size(321, 16)
        Me.Label92.TabIndex = 0
        Me.Label92.Text = "契約者負担分の発生率データのリスト出力"
        Me.Label92.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label92.UseCompatibleTextRendering = true
        '
        'Label95
        '
        Me.Label95.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label95.Location = New System.Drawing.Point(307, 31)
        Me.Label95.Name = "Label95"
        Me.Label95.Size = New System.Drawing.Size(22, 16)
        Me.Label95.TabIndex = 4
        Me.Label95.Text = " )"
        Me.Label95.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label175
        '
        Me.Label175.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label175.Location = New System.Drawing.Point(131, 31)
        Me.Label175.Name = "Label175"
        Me.Label175.Size = New System.Drawing.Size(92, 16)
        Me.Label175.TabIndex = 2
        Me.Label175.Text = "( 抽出件数 = "
        Me.Label175.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntSzenKysSorit
        '
        Me.lblCntSzenKysSorit.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntSzenKysSorit.Location = New System.Drawing.Point(218, 31)
        Me.lblCntSzenKysSorit.Name = "lblCntSzenKysSorit"
        Me.lblCntSzenKysSorit.Size = New System.Drawing.Size(83, 16)
        Me.lblCntSzenKysSorit.TabIndex = 3
        Me.lblCntSzenKysSorit.Text = "9,999,999"
        Me.lblCntSzenKysSorit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKojyosh
        '
        Me.pnlKojyosh.Controls.Add(Me.Label296)
        Me.pnlKojyosh.Controls.Add(Me.chkKojyosh)
        Me.pnlKojyosh.Controls.Add(Me.Label209)
        Me.pnlKojyosh.Controls.Add(Me.Panel10)
        Me.pnlKojyosh.Location = New System.Drawing.Point(6, 154)
        Me.pnlKojyosh.Name = "pnlKojyosh"
        Me.pnlKojyosh.Size = New System.Drawing.Size(400, 163)
        Me.pnlKojyosh.TabIndex = 1
        '
        'Label296
        '
        Me.Label296.Location = New System.Drawing.Point(10, 5)
        Me.Label296.Name = "Label296"
        Me.Label296.Size = New System.Drawing.Size(20, 20)
        Me.Label296.TabIndex = 136
        Me.Label296.Text = "・"
        Me.Label296.UseCompatibleTextRendering = true
        '
        'chkKojyosh
        '
        Me.chkKojyosh.AutoSize = true
        Me.chkKojyosh.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKojyosh.Location = New System.Drawing.Point(16, 4)
        Me.chkKojyosh.Name = "chkKojyosh"
        Me.chkKojyosh.Size = New System.Drawing.Size(147, 22)
        Me.chkKojyosh.TabIndex = 0
        Me.chkKojyosh.Text = "控除支払データの確認"
        Me.chkKojyosh.UseVisualStyleBackColor = true
        '
        'Label209
        '
        Me.Label209.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label209.Location = New System.Drawing.Point(36, 29)
        Me.Label209.Name = "Label209"
        Me.Label209.Size = New System.Drawing.Size(340, 65)
        Me.Label209.TabIndex = 1
        Me.Label209.Text = "V7の控除支払データは、10の自社支払へ移行されます。コンバート後の検証用リストを出力します。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※""9910.一括借上額""項目については自動で支払データとして作"& _ 
    "成される為、リストからは除外します。"
        Me.Label209.UseCompatibleTextRendering = true
        '
        'Panel10
        '
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel10.Controls.Add(Me.btnListKojyosh)
        Me.Panel10.Controls.Add(Me.Label199)
        Me.Panel10.Controls.Add(Me.Label200)
        Me.Panel10.Controls.Add(Me.Label201)
        Me.Panel10.Controls.Add(Me.lblCntKojyosh)
        Me.Panel10.Location = New System.Drawing.Point(36, 96)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(340, 60)
        Me.Panel10.TabIndex = 2
        '
        'btnListKojyosh
        '
        Me.btnListKojyosh.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListKojyosh.Image = CType(resources.GetObject("btnListKojyosh.Image"),System.Drawing.Image)
        Me.btnListKojyosh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListKojyosh.Location = New System.Drawing.Point(8, 27)
        Me.btnListKojyosh.Name = "btnListKojyosh"
        Me.btnListKojyosh.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListKojyosh.Size = New System.Drawing.Size(100, 25)
        Me.btnListKojyosh.TabIndex = 1
        Me.btnListKojyosh.Text = "リスト出力"
        Me.btnListKojyosh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListKojyosh.UseVisualStyleBackColor = true
        '
        'Label199
        '
        Me.Label199.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label199.Location = New System.Drawing.Point(8, 8)
        Me.Label199.Name = "Label199"
        Me.Label199.Size = New System.Drawing.Size(321, 16)
        Me.Label199.TabIndex = 0
        Me.Label199.Text = "一括借上額項目以外の控除支払データのリスト出力"
        Me.Label199.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label199.UseCompatibleTextRendering = true
        '
        'Label200
        '
        Me.Label200.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label200.Location = New System.Drawing.Point(307, 31)
        Me.Label200.Name = "Label200"
        Me.Label200.Size = New System.Drawing.Size(22, 16)
        Me.Label200.TabIndex = 4
        Me.Label200.Text = " )"
        Me.Label200.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label201
        '
        Me.Label201.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label201.Location = New System.Drawing.Point(131, 31)
        Me.Label201.Name = "Label201"
        Me.Label201.Size = New System.Drawing.Size(92, 16)
        Me.Label201.TabIndex = 2
        Me.Label201.Text = "( 抽出件数 = "
        Me.Label201.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntKojyosh
        '
        Me.lblCntKojyosh.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntKojyosh.Location = New System.Drawing.Point(218, 31)
        Me.lblCntKojyosh.Name = "lblCntKojyosh"
        Me.lblCntKojyosh.Size = New System.Drawing.Size(83, 16)
        Me.lblCntKojyosh.TabIndex = 3
        Me.lblCntKojyosh.Text = "9,999,999"
        Me.lblCntKojyosh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlBunkatumisyu
        '
        Me.pnlBunkatumisyu.Controls.Add(Me.Label293)
        Me.pnlBunkatumisyu.Controls.Add(Me.chkBunkatumisyu)
        Me.pnlBunkatumisyu.Controls.Add(Me.Label207)
        Me.pnlBunkatumisyu.Controls.Add(Me.Panel9)
        Me.pnlBunkatumisyu.Location = New System.Drawing.Point(6, 18)
        Me.pnlBunkatumisyu.Name = "pnlBunkatumisyu"
        Me.pnlBunkatumisyu.Size = New System.Drawing.Size(400, 130)
        Me.pnlBunkatumisyu.TabIndex = 0
        '
        'Label293
        '
        Me.Label293.Location = New System.Drawing.Point(10, 5)
        Me.Label293.Name = "Label293"
        Me.Label293.Size = New System.Drawing.Size(20, 20)
        Me.Label293.TabIndex = 136
        Me.Label293.Text = "・"
        Me.Label293.UseCompatibleTextRendering = true
        '
        'chkBunkatumisyu
        '
        Me.chkBunkatumisyu.AutoSize = true
        Me.chkBunkatumisyu.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkBunkatumisyu.Location = New System.Drawing.Point(16, 4)
        Me.chkBunkatumisyu.Name = "chkBunkatumisyu"
        Me.chkBunkatumisyu.Size = New System.Drawing.Size(195, 22)
        Me.chkBunkatumisyu.TabIndex = 0
        Me.chkBunkatumisyu.Text = "分割入金の未収分データの確認"
        Me.chkBunkatumisyu.UseVisualStyleBackColor = true
        '
        'Label207
        '
        Me.Label207.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label207.Location = New System.Drawing.Point(36, 29)
        Me.Label207.Name = "Label207"
        Me.Label207.Size = New System.Drawing.Size(340, 33)
        Me.Label207.TabIndex = 1
        Me.Label207.Text = "一部入金の未収分も移行対象となります。コンバート後の検証用リストを出力します。"
        Me.Label207.UseCompatibleTextRendering = true
        '
        'Panel9
        '
        Me.Panel9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel9.Controls.Add(Me.btnListBunkatumisyu)
        Me.Panel9.Controls.Add(Me.Label195)
        Me.Panel9.Controls.Add(Me.Label196)
        Me.Panel9.Controls.Add(Me.Label197)
        Me.Panel9.Controls.Add(Me.lblCntBunkatumisyu)
        Me.Panel9.Location = New System.Drawing.Point(36, 65)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(340, 60)
        Me.Panel9.TabIndex = 2
        '
        'btnListBunkatumisyu
        '
        Me.btnListBunkatumisyu.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnListBunkatumisyu.Image = CType(resources.GetObject("btnListBunkatumisyu.Image"),System.Drawing.Image)
        Me.btnListBunkatumisyu.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnListBunkatumisyu.Location = New System.Drawing.Point(8, 27)
        Me.btnListBunkatumisyu.Name = "btnListBunkatumisyu"
        Me.btnListBunkatumisyu.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnListBunkatumisyu.Size = New System.Drawing.Size(100, 25)
        Me.btnListBunkatumisyu.TabIndex = 1
        Me.btnListBunkatumisyu.Text = "リスト出力"
        Me.btnListBunkatumisyu.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnListBunkatumisyu.UseVisualStyleBackColor = true
        '
        'Label195
        '
        Me.Label195.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label195.Location = New System.Drawing.Point(8, 8)
        Me.Label195.Name = "Label195"
        Me.Label195.Size = New System.Drawing.Size(321, 16)
        Me.Label195.TabIndex = 0
        Me.Label195.Text = "一部入金の未収分データのリスト出力"
        Me.Label195.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label195.UseCompatibleTextRendering = true
        '
        'Label196
        '
        Me.Label196.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label196.Location = New System.Drawing.Point(307, 31)
        Me.Label196.Name = "Label196"
        Me.Label196.Size = New System.Drawing.Size(22, 16)
        Me.Label196.TabIndex = 4
        Me.Label196.Text = " )"
        Me.Label196.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label197
        '
        Me.Label197.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label197.Location = New System.Drawing.Point(131, 31)
        Me.Label197.Name = "Label197"
        Me.Label197.Size = New System.Drawing.Size(92, 16)
        Me.Label197.TabIndex = 2
        Me.Label197.Text = "( 抽出件数 = "
        Me.Label197.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCntBunkatumisyu
        '
        Me.lblCntBunkatumisyu.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblCntBunkatumisyu.Location = New System.Drawing.Point(218, 31)
        Me.lblCntBunkatumisyu.Name = "lblCntBunkatumisyu"
        Me.lblCntBunkatumisyu.Size = New System.Drawing.Size(83, 16)
        Me.lblCntBunkatumisyu.TabIndex = 3
        Me.lblCntBunkatumisyu.Text = "9,999,999"
        Me.lblCntBunkatumisyu.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label208
        '
        Me.Label208.AutoSize = true
        Me.Label208.Location = New System.Drawing.Point(794, 321)
        Me.Label208.Name = "Label208"
        Me.Label208.Size = New System.Drawing.Size(35, 18)
        Me.Label208.TabIndex = 0
        Me.Label208.Text = "2 / 2"
        '
        'lblHojyoDescription1
        '
        Me.lblHojyoDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblHojyoDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHojyoDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblHojyoDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblHojyoDescription1.Name = "lblHojyoDescription1"
        Me.lblHojyoDescription1.Size = New System.Drawing.Size(820, 80)
        Me.lblHojyoDescription1.TabIndex = 10
        Me.lblHojyoDescription1.Text = "コンバート後の検証用に、賃貸革命V7データのリスト出力を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(必要な場合に行って下さい)"
        Me.lblHojyoDescription1.UseCompatibleTextRendering = true
        '
        'tabPageHajimeni
        '
        Me.tabPageHajimeni.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHajimeni.Controls.Add(Me.lblDatacvHajimeniDescription1)
        Me.tabPageHajimeni.Controls.Add(Me.lblKDatacvHajimeniLabel)
        Me.tabPageHajimeni.Controls.Add(Me.lblHDatacvHajimeniLabel)
        Me.tabPageHajimeni.Controls.Add(Me.btnDoui)
        Me.tabPageHajimeni.Controls.Add(Me.GroupBox10)
        Me.tabPageHajimeni.Controls.Add(Me.PictureBox14)
        Me.tabPageHajimeni.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHajimeni.Name = "tabPageHajimeni"
        Me.tabPageHajimeni.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHajimeni.Size = New System.Drawing.Size(912, 549)
        Me.tabPageHajimeni.TabIndex = 6
        Me.tabPageHajimeni.Text = " はじめに"
        '
        'lblDatacvHajimeniDescription1
        '
        Me.lblDatacvHajimeniDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblDatacvHajimeniDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!)
        Me.lblDatacvHajimeniDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblDatacvHajimeniDescription1.Location = New System.Drawing.Point(35, 49)
        Me.lblDatacvHajimeniDescription1.Name = "lblDatacvHajimeniDescription1"
        Me.lblDatacvHajimeniDescription1.Size = New System.Drawing.Size(540, 35)
        Me.lblDatacvHajimeniDescription1.TabIndex = 1
        Me.lblDatacvHajimeniDescription1.Text = "以下の留意事項を確認の上、同意ボタンを押して下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(同意後「次へ」ボタンを押して下さい。)"
        Me.lblDatacvHajimeniDescription1.UseCompatibleTextRendering = true
        '
        'lblKDatacvHajimeniLabel
        '
        Me.lblKDatacvHajimeniLabel.BackColor = System.Drawing.SystemColors.Menu
        Me.lblKDatacvHajimeniLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKDatacvHajimeniLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.lblKDatacvHajimeniLabel.Location = New System.Drawing.Point(30, 20)
        Me.lblKDatacvHajimeniLabel.Name = "lblKDatacvHajimeniLabel"
        Me.lblKDatacvHajimeniLabel.Size = New System.Drawing.Size(572, 80)
        Me.lblKDatacvHajimeniLabel.TabIndex = 0
        Me.lblKDatacvHajimeniLabel.Text = "「賃貸革命V7」から「賃貸革命10」へデータコンバートを行います。"
        Me.lblKDatacvHajimeniLabel.UseCompatibleTextRendering = true
        '
        'lblHDatacvHajimeniLabel
        '
        Me.lblHDatacvHajimeniLabel.BackColor = System.Drawing.SystemColors.Menu
        Me.lblHDatacvHajimeniLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHDatacvHajimeniLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.lblHDatacvHajimeniLabel.Location = New System.Drawing.Point(30, 20)
        Me.lblHDatacvHajimeniLabel.Name = "lblHDatacvHajimeniLabel"
        Me.lblHDatacvHajimeniLabel.Size = New System.Drawing.Size(587, 80)
        Me.lblHDatacvHajimeniLabel.TabIndex = 139
        Me.lblHDatacvHajimeniLabel.Text = "「中間ファイル」から「賃貸革命10」へデータコンバートを行います。"
        Me.lblHDatacvHajimeniLabel.UseCompatibleTextRendering = true
        '
        'btnDoui
        '
        Me.btnDoui.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnDoui.ForeColor = System.Drawing.Color.Red
        Me.btnDoui.Location = New System.Drawing.Point(632, 53)
        Me.btnDoui.Name = "btnDoui"
        Me.btnDoui.Size = New System.Drawing.Size(251, 43)
        Me.btnDoui.TabIndex = 0
        Me.btnDoui.Text = "留意事項に同意します"
        Me.btnDoui.UseVisualStyleBackColor = true
        '
        'GroupBox10
        '
        Me.GroupBox10.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmtH99)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmtK99)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmtH01)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmtH02)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt11)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmtK01)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt06)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt05)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt07)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt10)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt02)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt08)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt04)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt03)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt01)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmtK02)
        Me.GroupBox10.Controls.Add(Me.pnlDcFstCmt09)
        Me.GroupBox10.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox10.Location = New System.Drawing.Point(24, 115)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(865, 423)
        Me.GroupBox10.TabIndex = 2
        Me.GroupBox10.TabStop = false
        Me.GroupBox10.Text = "【留意事項】"
        '
        'pnlDcFstCmtH99
        '
        Me.pnlDcFstCmtH99.Controls.Add(Me.Label301)
        Me.pnlDcFstCmtH99.Controls.Add(Me.Label302)
        Me.pnlDcFstCmtH99.Location = New System.Drawing.Point(54, 101)
        Me.pnlDcFstCmtH99.Name = "pnlDcFstCmtH99"
        Me.pnlDcFstCmtH99.Size = New System.Drawing.Size(395, 115)
        Me.pnlDcFstCmtH99.TabIndex = 26
        '
        'Label301
        '
        Me.Label301.Location = New System.Drawing.Point(3, 3)
        Me.Label301.Name = "Label301"
        Me.Label301.Size = New System.Drawing.Size(200, 16)
        Me.Label301.TabIndex = 14
        Me.Label301.Text = "<コンバート対象>"
        Me.Label301.UseCompatibleTextRendering = true
        '
        'Label302
        '
        Me.Label302.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label302.Location = New System.Drawing.Point(23, 4)
        Me.Label302.Name = "Label302"
        Me.Label302.Size = New System.Drawing.Size(370, 105)
        Me.Label302.TabIndex = 15
        Me.Label302.Text = ""&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"物件情報、部屋情報"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"契約情報(最新契約のみ対象。契約・更新歴、解約は対象外)"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"家主情報、契約者情報、業者情報、口座情報、自社情報"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※請求・入金・送金、画"& _ 
    "像は対応していません。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※コンバート後に請求構築などの作業を行う必要があります。"
        Me.Label302.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmtK99
        '
        Me.pnlDcFstCmtK99.Controls.Add(Me.Label299)
        Me.pnlDcFstCmtK99.Controls.Add(Me.Label300)
        Me.pnlDcFstCmtK99.Location = New System.Drawing.Point(54, 222)
        Me.pnlDcFstCmtK99.Name = "pnlDcFstCmtK99"
        Me.pnlDcFstCmtK99.Size = New System.Drawing.Size(395, 115)
        Me.pnlDcFstCmtK99.TabIndex = 25
        '
        'Label299
        '
        Me.Label299.Location = New System.Drawing.Point(3, 3)
        Me.Label299.Name = "Label299"
        Me.Label299.Size = New System.Drawing.Size(200, 16)
        Me.Label299.TabIndex = 14
        Me.Label299.Text = "<コンバート対象>"
        Me.Label299.UseCompatibleTextRendering = true
        '
        'Label300
        '
        Me.Label300.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label300.Location = New System.Drawing.Point(23, 4)
        Me.Label300.Name = "Label300"
        Me.Label300.Size = New System.Drawing.Size(370, 105)
        Me.Label300.TabIndex = 15
        Me.Label300.Text = ""&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"物件情報、部屋情報、契約情報、家主情報、契約者情報"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"業者情報、口座情報、自社情報、未収預り金、控除、変動費"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"初期設定、画像、クレーム、修繕、ポータル連動"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※請求・入金・送金は対応していません。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※コンバート後に請求構築などの作業を行う必要があります。"
        Me.Label300.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmtH01
        '
        Me.pnlDcFstCmtH01.Controls.Add(Me.Label158)
        Me.pnlDcFstCmtH01.Controls.Add(Me.Label159)
        Me.pnlDcFstCmtH01.Location = New System.Drawing.Point(455, 175)
        Me.pnlDcFstCmtH01.Name = "pnlDcFstCmtH01"
        Me.pnlDcFstCmtH01.Size = New System.Drawing.Size(395, 28)
        Me.pnlDcFstCmtH01.TabIndex = 24
        '
        'Label158
        '
        Me.Label158.Location = New System.Drawing.Point(3, 3)
        Me.Label158.Name = "Label158"
        Me.Label158.Size = New System.Drawing.Size(14, 16)
        Me.Label158.TabIndex = 4
        Me.Label158.Text = "・"
        Me.Label158.UseCompatibleTextRendering = true
        '
        'Label159
        '
        Me.Label159.Location = New System.Drawing.Point(23, 4)
        Me.Label159.Name = "Label159"
        Me.Label159.Size = New System.Drawing.Size(370, 18)
        Me.Label159.TabIndex = 5
        Me.Label159.Text = "中間ファイルへ事前にデータを登録する必要があります。"
        Me.Label159.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmtH02
        '
        Me.pnlDcFstCmtH02.Controls.Add(Me.Label27)
        Me.pnlDcFstCmtH02.Controls.Add(Me.Label28)
        Me.pnlDcFstCmtH02.Location = New System.Drawing.Point(455, 202)
        Me.pnlDcFstCmtH02.Name = "pnlDcFstCmtH02"
        Me.pnlDcFstCmtH02.Size = New System.Drawing.Size(395, 63)
        Me.pnlDcFstCmtH02.TabIndex = 23
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(3, 3)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(14, 16)
        Me.Label27.TabIndex = 4
        Me.Label27.Text = "・"
        Me.Label27.UseCompatibleTextRendering = true
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(23, 4)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(370, 55)
        Me.Label28.TabIndex = 5
        Me.Label28.Text = "弊社にて用意した中間ファイルを必ずご使用下さい。中間ファイルの加工・それ以外(お客様で用意されたファイルなど)については対応していません。"
        Me.Label28.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt11
        '
        Me.pnlDcFstCmt11.Controls.Add(Me.Label21)
        Me.pnlDcFstCmt11.Controls.Add(Me.Label25)
        Me.pnlDcFstCmt11.Location = New System.Drawing.Point(455, 131)
        Me.pnlDcFstCmt11.Name = "pnlDcFstCmt11"
        Me.pnlDcFstCmt11.Size = New System.Drawing.Size(395, 45)
        Me.pnlDcFstCmt11.TabIndex = 22
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(3, 3)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(14, 16)
        Me.Label21.TabIndex = 4
        Me.Label21.Text = "・"
        Me.Label21.UseCompatibleTextRendering = true
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(23, 4)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(370, 38)
        Me.Label25.TabIndex = 5
        Me.Label25.Text = "運用開始(本稼働)後のコンバートは、整合性が取れなくなる場合がある為、原則行わないで下さい。"
        Me.Label25.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmtK01
        '
        Me.pnlDcFstCmtK01.Controls.Add(Me.Label371)
        Me.pnlDcFstCmtK01.Controls.Add(Me.Label372)
        Me.pnlDcFstCmtK01.Location = New System.Drawing.Point(455, 281)
        Me.pnlDcFstCmtK01.Name = "pnlDcFstCmtK01"
        Me.pnlDcFstCmtK01.Size = New System.Drawing.Size(395, 28)
        Me.pnlDcFstCmtK01.TabIndex = 21
        '
        'Label371
        '
        Me.Label371.Location = New System.Drawing.Point(3, 3)
        Me.Label371.Name = "Label371"
        Me.Label371.Size = New System.Drawing.Size(14, 16)
        Me.Label371.TabIndex = 4
        Me.Label371.Text = "・"
        Me.Label371.UseCompatibleTextRendering = true
        '
        'Label372
        '
        Me.Label372.Location = New System.Drawing.Point(23, 4)
        Me.Label372.Name = "Label372"
        Me.Label372.Size = New System.Drawing.Size(370, 18)
        Me.Label372.TabIndex = 5
        Me.Label372.Text = "カスタマイズ部分のコンバートは対応していません。"
        Me.Label372.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt06
        '
        Me.pnlDcFstCmt06.Controls.Add(Me.Label364)
        Me.pnlDcFstCmt06.Controls.Add(Me.Label369)
        Me.pnlDcFstCmt06.Location = New System.Drawing.Point(15, 264)
        Me.pnlDcFstCmt06.Name = "pnlDcFstCmt06"
        Me.pnlDcFstCmt06.Size = New System.Drawing.Size(395, 28)
        Me.pnlDcFstCmt06.TabIndex = 20
        '
        'Label364
        '
        Me.Label364.Location = New System.Drawing.Point(3, 3)
        Me.Label364.Name = "Label364"
        Me.Label364.Size = New System.Drawing.Size(14, 16)
        Me.Label364.TabIndex = 6
        Me.Label364.Text = "・"
        Me.Label364.UseCompatibleTextRendering = true
        '
        'Label369
        '
        Me.Label369.Location = New System.Drawing.Point(23, 4)
        Me.Label369.Name = "Label369"
        Me.Label369.Size = New System.Drawing.Size(370, 18)
        Me.Label369.TabIndex = 7
        Me.Label369.Text = "暦(こよみ)に存在しない不適切な日付はコンバートできません。"
        Me.Label369.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt05
        '
        Me.pnlDcFstCmt05.Controls.Add(Me.Label362)
        Me.pnlDcFstCmt05.Controls.Add(Me.Label363)
        Me.pnlDcFstCmt05.Location = New System.Drawing.Point(15, 202)
        Me.pnlDcFstCmt05.Name = "pnlDcFstCmt05"
        Me.pnlDcFstCmt05.Size = New System.Drawing.Size(395, 63)
        Me.pnlDcFstCmt05.TabIndex = 15
        '
        'Label362
        '
        Me.Label362.Location = New System.Drawing.Point(3, 3)
        Me.Label362.Name = "Label362"
        Me.Label362.Size = New System.Drawing.Size(14, 16)
        Me.Label362.TabIndex = 10
        Me.Label362.Text = "・"
        Me.Label362.UseCompatibleTextRendering = true
        '
        'Label363
        '
        Me.Label363.Location = New System.Drawing.Point(23, 4)
        Me.Label363.Name = "Label363"
        Me.Label363.Size = New System.Drawing.Size(370, 55)
        Me.Label363.TabIndex = 11
        Me.Label363.Text = "ある情報を一意に識別するコードが重複しているデータの場合、最初に読み込まれるデータのみコンバート致します。(例：複数の同一物件が存在する場合、1 物件のみコンバー"& _ 
    "トします)"
        Me.Label363.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt07
        '
        Me.pnlDcFstCmt07.Controls.Add(Me.Label122)
        Me.pnlDcFstCmt07.Controls.Add(Me.Label104)
        Me.pnlDcFstCmt07.Location = New System.Drawing.Point(15, 291)
        Me.pnlDcFstCmt07.Name = "pnlDcFstCmt07"
        Me.pnlDcFstCmt07.Size = New System.Drawing.Size(395, 45)
        Me.pnlDcFstCmt07.TabIndex = 6
        '
        'Label122
        '
        Me.Label122.Location = New System.Drawing.Point(3, 3)
        Me.Label122.Name = "Label122"
        Me.Label122.Size = New System.Drawing.Size(14, 16)
        Me.Label122.TabIndex = 2
        Me.Label122.Text = "・"
        Me.Label122.UseCompatibleTextRendering = true
        '
        'Label104
        '
        Me.Label104.Location = New System.Drawing.Point(23, 4)
        Me.Label104.Name = "Label104"
        Me.Label104.Size = New System.Drawing.Size(370, 38)
        Me.Label104.TabIndex = 3
        Me.Label104.Text = "機種依存文字・旧字体・制御文字(改行・タブ文字など)文字は正しく変換できない場合があります。"
        Me.Label104.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt10
        '
        Me.pnlDcFstCmt10.Controls.Add(Me.Label123)
        Me.pnlDcFstCmt10.Controls.Add(Me.Label105)
        Me.pnlDcFstCmt10.Location = New System.Drawing.Point(455, 87)
        Me.pnlDcFstCmt10.Name = "pnlDcFstCmt10"
        Me.pnlDcFstCmt10.Size = New System.Drawing.Size(395, 45)
        Me.pnlDcFstCmt10.TabIndex = 8
        '
        'Label123
        '
        Me.Label123.Location = New System.Drawing.Point(3, 3)
        Me.Label123.Name = "Label123"
        Me.Label123.Size = New System.Drawing.Size(14, 16)
        Me.Label123.TabIndex = 4
        Me.Label123.Text = "・"
        Me.Label123.UseCompatibleTextRendering = true
        '
        'Label105
        '
        Me.Label105.Location = New System.Drawing.Point(23, 4)
        Me.Label105.Name = "Label105"
        Me.Label105.Size = New System.Drawing.Size(370, 38)
        Me.Label105.TabIndex = 5
        Me.Label105.Text = "本プログラム処理中は、他の作業(賃貸革命やEXCELの使用)を行わないで下さい。"
        Me.Label105.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt02
        '
        Me.pnlDcFstCmt02.Controls.Add(Me.Label127)
        Me.pnlDcFstCmt02.Controls.Add(Me.Label133)
        Me.pnlDcFstCmt02.Location = New System.Drawing.Point(15, 69)
        Me.pnlDcFstCmt02.Name = "pnlDcFstCmt02"
        Me.pnlDcFstCmt02.Size = New System.Drawing.Size(395, 63)
        Me.pnlDcFstCmt02.TabIndex = 10
        '
        'Label127
        '
        Me.Label127.Location = New System.Drawing.Point(3, 3)
        Me.Label127.Name = "Label127"
        Me.Label127.Size = New System.Drawing.Size(14, 16)
        Me.Label127.TabIndex = 6
        Me.Label127.Text = "・"
        Me.Label127.UseCompatibleTextRendering = true
        '
        'Label133
        '
        Me.Label133.Location = New System.Drawing.Point(23, 4)
        Me.Label133.Name = "Label133"
        Me.Label133.Size = New System.Drawing.Size(370, 55)
        Me.Label133.TabIndex = 7
        Me.Label133.Text = "複数のデータベースから1 つのデータベースへ、または1つのデータベースを複数のデータベースに分けるようなコンバートは対応しておりません。"
        Me.Label133.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt08
        '
        Me.pnlDcFstCmt08.Controls.Add(Me.Label126)
        Me.pnlDcFstCmt08.Controls.Add(Me.Label132)
        Me.pnlDcFstCmt08.Location = New System.Drawing.Point(15, 335)
        Me.pnlDcFstCmt08.Name = "pnlDcFstCmt08"
        Me.pnlDcFstCmt08.Size = New System.Drawing.Size(395, 45)
        Me.pnlDcFstCmt08.TabIndex = 12
        '
        'Label126
        '
        Me.Label126.Location = New System.Drawing.Point(3, 3)
        Me.Label126.Name = "Label126"
        Me.Label126.Size = New System.Drawing.Size(14, 16)
        Me.Label126.TabIndex = 8
        Me.Label126.Text = "・"
        Me.Label126.UseCompatibleTextRendering = true
        '
        'Label132
        '
        Me.Label132.Location = New System.Drawing.Point(23, 4)
        Me.Label132.Name = "Label132"
        Me.Label132.Size = New System.Drawing.Size(370, 33)
        Me.Label132.TabIndex = 9
        Me.Label132.Text = "親情報がない子情報のコンバートはできません。(例：部屋情報がコンバートされていない契約情報は移行できません)"
        Me.Label132.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt04
        '
        Me.pnlDcFstCmt04.Controls.Add(Me.Label119)
        Me.pnlDcFstCmt04.Controls.Add(Me.Label116)
        Me.pnlDcFstCmt04.Location = New System.Drawing.Point(15, 175)
        Me.pnlDcFstCmt04.Name = "pnlDcFstCmt04"
        Me.pnlDcFstCmt04.Size = New System.Drawing.Size(395, 28)
        Me.pnlDcFstCmt04.TabIndex = 14
        '
        'Label119
        '
        Me.Label119.Location = New System.Drawing.Point(3, 3)
        Me.Label119.Name = "Label119"
        Me.Label119.Size = New System.Drawing.Size(14, 16)
        Me.Label119.TabIndex = 10
        Me.Label119.Text = "・"
        Me.Label119.UseCompatibleTextRendering = true
        '
        'Label116
        '
        Me.Label116.Location = New System.Drawing.Point(23, 4)
        Me.Label116.Name = "Label116"
        Me.Label116.Size = New System.Drawing.Size(370, 18)
        Me.Label116.TabIndex = 11
        Me.Label116.Text = "必須項目が不足しているデータはコンバートできません。"
        Me.Label116.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt03
        '
        Me.pnlDcFstCmt03.Controls.Add(Me.Label97)
        Me.pnlDcFstCmt03.Controls.Add(Me.Label101)
        Me.pnlDcFstCmt03.Location = New System.Drawing.Point(15, 131)
        Me.pnlDcFstCmt03.Name = "pnlDcFstCmt03"
        Me.pnlDcFstCmt03.Size = New System.Drawing.Size(395, 45)
        Me.pnlDcFstCmt03.TabIndex = 19
        '
        'Label97
        '
        Me.Label97.Location = New System.Drawing.Point(23, 4)
        Me.Label97.Name = "Label97"
        Me.Label97.Size = New System.Drawing.Size(370, 38)
        Me.Label97.TabIndex = 13
        Me.Label97.Text = "移行仕様書に記載していない、または従っていない登録データはコンバートできません。"
        Me.Label97.UseCompatibleTextRendering = true
        '
        'Label101
        '
        Me.Label101.Location = New System.Drawing.Point(3, 3)
        Me.Label101.Name = "Label101"
        Me.Label101.Size = New System.Drawing.Size(14, 16)
        Me.Label101.TabIndex = 12
        Me.Label101.Text = "・"
        Me.Label101.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt01
        '
        Me.pnlDcFstCmt01.Controls.Add(Me.Label129)
        Me.pnlDcFstCmt01.Controls.Add(Me.Label135)
        Me.pnlDcFstCmt01.Location = New System.Drawing.Point(15, 25)
        Me.pnlDcFstCmt01.Name = "pnlDcFstCmt01"
        Me.pnlDcFstCmt01.Size = New System.Drawing.Size(395, 45)
        Me.pnlDcFstCmt01.TabIndex = 18
        '
        'Label129
        '
        Me.Label129.Location = New System.Drawing.Point(3, 3)
        Me.Label129.Name = "Label129"
        Me.Label129.Size = New System.Drawing.Size(14, 16)
        Me.Label129.TabIndex = 0
        Me.Label129.Text = "・"
        Me.Label129.UseCompatibleTextRendering = true
        '
        'Label135
        '
        Me.Label135.Location = New System.Drawing.Point(23, 4)
        Me.Label135.Name = "Label135"
        Me.Label135.Size = New System.Drawing.Size(370, 38)
        Me.Label135.TabIndex = 1
        Me.Label135.Text = "本プログラムは、賃貸革命10 がインストールされている環境のみで動作します。"
        Me.Label135.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmtK02
        '
        Me.pnlDcFstCmtK02.Controls.Add(Me.Label40)
        Me.pnlDcFstCmtK02.Controls.Add(Me.Label37)
        Me.pnlDcFstCmtK02.Location = New System.Drawing.Point(455, 308)
        Me.pnlDcFstCmtK02.Name = "pnlDcFstCmtK02"
        Me.pnlDcFstCmtK02.Size = New System.Drawing.Size(395, 80)
        Me.pnlDcFstCmtK02.TabIndex = 18
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(3, 3)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(14, 16)
        Me.Label40.TabIndex = 14
        Me.Label40.Text = "・"
        Me.Label40.UseCompatibleTextRendering = true
        '
        'Label37
        '
        Me.Label37.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label37.Location = New System.Drawing.Point(23, 4)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(370, 70)
        Me.Label37.TabIndex = 15
        Me.Label37.Text = "賃貸革命V7からのコンバートのみ対象とします。下位バージョン(V5,V6)の場合、事前にバージョンアップをしておく必要があります。バージョンアップされていない場合"& _ 
    "、お手数ですがサポートへお問い合わせ下さい。"
        Me.Label37.UseCompatibleTextRendering = true
        '
        'pnlDcFstCmt09
        '
        Me.pnlDcFstCmt09.Controls.Add(Me.Label326)
        Me.pnlDcFstCmt09.Controls.Add(Me.Label325)
        Me.pnlDcFstCmt09.Location = New System.Drawing.Point(455, 25)
        Me.pnlDcFstCmt09.Name = "pnlDcFstCmt09"
        Me.pnlDcFstCmt09.Size = New System.Drawing.Size(395, 63)
        Me.pnlDcFstCmt09.TabIndex = 16
        '
        'Label326
        '
        Me.Label326.Location = New System.Drawing.Point(3, 3)
        Me.Label326.Name = "Label326"
        Me.Label326.Size = New System.Drawing.Size(14, 16)
        Me.Label326.TabIndex = 16
        Me.Label326.Text = "・"
        Me.Label326.UseCompatibleTextRendering = true
        '
        'Label325
        '
        Me.Label325.Location = New System.Drawing.Point(23, 4)
        Me.Label325.Name = "Label325"
        Me.Label325.Size = New System.Drawing.Size(370, 55)
        Me.Label325.TabIndex = 17
        Me.Label325.Text = "金融機関・支店・沿線・駅・住所などのマスタ値を参照するデータにおいて、賃貸革命10 へ登録されている設定値以外はコンバートできません。"
        Me.Label325.UseCompatibleTextRendering = true
        '
        'PictureBox14
        '
        Me.PictureBox14.Image = CType(resources.GetObject("PictureBox14.Image"),System.Drawing.Image)
        Me.PictureBox14.Location = New System.Drawing.Point(562, 342)
        Me.PictureBox14.Name = "PictureBox14"
        Me.PictureBox14.Size = New System.Drawing.Size(254, 185)
        Me.PictureBox14.TabIndex = 138
        Me.PictureBox14.TabStop = false
        '
        'tabPageSelect
        '
        Me.tabPageSelect.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageSelect.Controls.Add(Me.grpExistMidToBaseMid)
        Me.tabPageSelect.Controls.Add(Me.btnAllChk)
        Me.tabPageSelect.Controls.Add(Me.lblDatacvSelectCaution)
        Me.tabPageSelect.Controls.Add(Me.lblLine2)
        Me.tabPageSelect.Controls.Add(Me.lblHidden2)
        Me.tabPageSelect.Controls.Add(Me.grpMiddleFile)
        Me.tabPageSelect.Controls.Add(Me.tabCtrlCVItem)
        Me.tabPageSelect.Controls.Add(Me.pnlRekiClear)
        Me.tabPageSelect.Controls.Add(Me.lblDatacvSelectDescription1)
        Me.tabPageSelect.Location = New System.Drawing.Point(4, 27)
        Me.tabPageSelect.Name = "tabPageSelect"
        Me.tabPageSelect.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageSelect.Size = New System.Drawing.Size(912, 549)
        Me.tabPageSelect.TabIndex = 2
        Me.tabPageSelect.Text = " 対象項目選択"
        '
        'grpExistMidToBaseMid
        '
        Me.grpExistMidToBaseMid.Font = New System.Drawing.Font("メイリオ", 8.25!)
        Me.grpExistMidToBaseMid.Location = New System.Drawing.Point(544, -14)
        Me.grpExistMidToBaseMid.Name = "grpExistMidToBaseMid"
        Me.grpExistMidToBaseMid.Size = New System.Drawing.Size(368, 48)
        Me.grpExistMidToBaseMid.TabIndex = 140
        Me.grpExistMidToBaseMid.TabStop = false
        '
        'btnAllChk
        '
        Me.btnAllChk.Image = CType(resources.GetObject("btnAllChk.Image"),System.Drawing.Image)
        Me.btnAllChk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAllChk.Location = New System.Drawing.Point(689, 87)
        Me.btnAllChk.Name = "btnAllChk"
        Me.btnAllChk.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnAllChk.Size = New System.Drawing.Size(200, 30)
        Me.btnAllChk.TabIndex = 3
        Me.btnAllChk.Text = " 画面毎全チェックON/OFF"
        Me.btnAllChk.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnAllChk.UseVisualStyleBackColor = true
        '
        'lblDatacvSelectCaution
        '
        Me.lblDatacvSelectCaution.BackColor = System.Drawing.SystemColors.Menu
        Me.lblDatacvSelectCaution.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblDatacvSelectCaution.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblDatacvSelectCaution.Location = New System.Drawing.Point(30, 59)
        Me.lblDatacvSelectCaution.Name = "lblDatacvSelectCaution"
        Me.lblDatacvSelectCaution.Size = New System.Drawing.Size(490, 58)
        Me.lblDatacvSelectCaution.TabIndex = 1
        Me.lblDatacvSelectCaution.Text = "※選択項目の組み合わせ(項目の親子関係)によっては、正しく移行されない場合"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　があります。項目の親子関係を考慮した上で選択下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(例：物件情報が未移行の場"& _ 
    "合、部屋情報や契約情報の移行はできません)"
        Me.lblDatacvSelectCaution.UseCompatibleTextRendering = true
        '
        'lblLine2
        '
        Me.lblLine2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLine2.Location = New System.Drawing.Point(31, 216)
        Me.lblLine2.Name = "lblLine2"
        Me.lblLine2.Size = New System.Drawing.Size(853, 2)
        Me.lblLine2.TabIndex = 6
        '
        'lblHidden2
        '
        Me.lblHidden2.Location = New System.Drawing.Point(6, 190)
        Me.lblHidden2.Name = "lblHidden2"
        Me.lblHidden2.Size = New System.Drawing.Size(18, 28)
        Me.lblHidden2.TabIndex = 6
        Me.lblHidden2.Text = "　"
        '
        'grpMiddleFile
        '
        Me.grpMiddleFile.Controls.Add(Me.Label349)
        Me.grpMiddleFile.Controls.Add(Me.lblMiddleFile2)
        Me.grpMiddleFile.Controls.Add(Me.txtMidDirPath2)
        Me.grpMiddleFile.Controls.Add(Me.btnMidDirSeach2)
        Me.grpMiddleFile.Controls.Add(Me.btnHJizenTyukanOpen2)
        Me.grpMiddleFile.Controls.Add(Me.Label328)
        Me.grpMiddleFile.Controls.Add(Me.Label124)
        Me.grpMiddleFile.Controls.Add(Me.btnMidFileCheck)
        Me.grpMiddleFile.Controls.Add(Me.btnMidDirLogSeach)
        Me.grpMiddleFile.Controls.Add(Me.txtMidDirLogPath)
        Me.grpMiddleFile.Controls.Add(Me.lblMiddleFileLog)
        Me.grpMiddleFile.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpMiddleFile.Location = New System.Drawing.Point(30, 114)
        Me.grpMiddleFile.Name = "grpMiddleFile"
        Me.grpMiddleFile.Size = New System.Drawing.Size(854, 73)
        Me.grpMiddleFile.TabIndex = 4
        Me.grpMiddleFile.TabStop = false
        '
        'Label349
        '
        Me.Label349.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label349.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label349.Location = New System.Drawing.Point(158, 16)
        Me.Label349.Name = "Label349"
        Me.Label349.Size = New System.Drawing.Size(46, 20)
        Me.Label349.TabIndex = 143
        Me.Label349.Text = "※必須"
        Me.Label349.UseCompatibleTextRendering = true
        '
        'lblMiddleFile2
        '
        Me.lblMiddleFile2.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblMiddleFile2.Location = New System.Drawing.Point(397, 18)
        Me.lblMiddleFile2.Name = "lblMiddleFile2"
        Me.lblMiddleFile2.Size = New System.Drawing.Size(117, 20)
        Me.lblMiddleFile2.TabIndex = 140
        Me.lblMiddleFile2.Text = "中間ファイル格納先"
        Me.lblMiddleFile2.UseCompatibleTextRendering = true
        '
        'txtMidDirPath2
        '
        Me.txtMidDirPath2.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtMidDirPath2.Location = New System.Drawing.Point(520, 15)
        Me.txtMidDirPath2.Name = "txtMidDirPath2"
        Me.txtMidDirPath2.Size = New System.Drawing.Size(186, 24)
        Me.txtMidDirPath2.TabIndex = 141
        '
        'btnMidDirSeach2
        '
        Me.btnMidDirSeach2.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMidDirSeach2.Location = New System.Drawing.Point(712, 14)
        Me.btnMidDirSeach2.Name = "btnMidDirSeach2"
        Me.btnMidDirSeach2.Size = New System.Drawing.Size(30, 25)
        Me.btnMidDirSeach2.TabIndex = 142
        Me.btnMidDirSeach2.Text = "..."
        Me.btnMidDirSeach2.UseVisualStyleBackColor = true
        '
        'btnHJizenTyukanOpen2
        '
        Me.btnHJizenTyukanOpen2.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnHJizenTyukanOpen2.Image = CType(resources.GetObject("btnHJizenTyukanOpen2.Image"),System.Drawing.Image)
        Me.btnHJizenTyukanOpen2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnHJizenTyukanOpen2.Location = New System.Drawing.Point(748, 14)
        Me.btnHJizenTyukanOpen2.Name = "btnHJizenTyukanOpen2"
        Me.btnHJizenTyukanOpen2.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnHJizenTyukanOpen2.Size = New System.Drawing.Size(100, 25)
        Me.btnHJizenTyukanOpen2.TabIndex = 140
        Me.btnHJizenTyukanOpen2.Text = "  開　く"
        Me.btnHJizenTyukanOpen2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnHJizenTyukanOpen2.UseVisualStyleBackColor = true
        '
        'Label328
        '
        Me.Label328.BackColor = System.Drawing.SystemColors.Menu
        Me.Label328.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label328.ForeColor = System.Drawing.Color.Black
        Me.Label328.Location = New System.Drawing.Point(29, 35)
        Me.Label328.Name = "Label328"
        Me.Label328.Size = New System.Drawing.Size(315, 32)
        Me.Label328.TabIndex = 138
        Me.Label328.Text = "選択された項目を対象に、中間ファイルへ登録したデータの"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"型やサイズなどの正当性をチェックします。"
        Me.Label328.UseCompatibleTextRendering = true
        '
        'Label124
        '
        Me.Label124.BackColor = System.Drawing.SystemColors.Menu
        Me.Label124.Font = New System.Drawing.Font("メイリオ", 9!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline),System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label124.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label124.Location = New System.Drawing.Point(15, 15)
        Me.Label124.Name = "Label124"
        Me.Label124.Size = New System.Drawing.Size(173, 19)
        Me.Label124.TabIndex = 7
        Me.Label124.Text = "中間ファイルチェック"
        Me.Label124.UseCompatibleTextRendering = true
        '
        'btnMidFileCheck
        '
        Me.btnMidFileCheck.BackColor = System.Drawing.SystemColors.Menu
        Me.btnMidFileCheck.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMidFileCheck.Image = CType(resources.GetObject("btnMidFileCheck.Image"),System.Drawing.Image)
        Me.btnMidFileCheck.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnMidFileCheck.Location = New System.Drawing.Point(748, 43)
        Me.btnMidFileCheck.Name = "btnMidFileCheck"
        Me.btnMidFileCheck.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnMidFileCheck.Size = New System.Drawing.Size(100, 25)
        Me.btnMidFileCheck.TabIndex = 6
        Me.btnMidFileCheck.Text = " チェック"
        Me.btnMidFileCheck.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnMidFileCheck.UseVisualStyleBackColor = true
        '
        'btnMidDirLogSeach
        '
        Me.btnMidDirLogSeach.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnMidDirLogSeach.Location = New System.Drawing.Point(712, 43)
        Me.btnMidDirLogSeach.Name = "btnMidDirLogSeach"
        Me.btnMidDirLogSeach.Size = New System.Drawing.Size(30, 25)
        Me.btnMidDirLogSeach.TabIndex = 5
        Me.btnMidDirLogSeach.Text = "..."
        Me.btnMidDirLogSeach.UseVisualStyleBackColor = true
        '
        'txtMidDirLogPath
        '
        Me.txtMidDirLogPath.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtMidDirLogPath.Location = New System.Drawing.Point(520, 44)
        Me.txtMidDirLogPath.Name = "txtMidDirLogPath"
        Me.txtMidDirLogPath.Size = New System.Drawing.Size(186, 24)
        Me.txtMidDirLogPath.TabIndex = 4
        '
        'lblMiddleFileLog
        '
        Me.lblMiddleFileLog.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblMiddleFileLog.Location = New System.Drawing.Point(397, 47)
        Me.lblMiddleFileLog.Name = "lblMiddleFileLog"
        Me.lblMiddleFileLog.Size = New System.Drawing.Size(117, 20)
        Me.lblMiddleFileLog.TabIndex = 3
        Me.lblMiddleFileLog.Text = "結果ファイル格納先"
        Me.lblMiddleFileLog.UseCompatibleTextRendering = true
        '
        'tabCtrlCVItem
        '
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageKizon110)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageKizon120)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageKizon130)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase110)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase120)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase130)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase140)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase150)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase160)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase170)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase180)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase190)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase200)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase210)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageBase900)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageHanyo110)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageHanyo120)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageHanyo130)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageHanyo140)
        Me.tabCtrlCVItem.Controls.Add(Me.tabPageHanyo150)
        Me.tabCtrlCVItem.Controls.Add(Me.TabPage1)
        Me.tabCtrlCVItem.Location = New System.Drawing.Point(30, 190)
        Me.tabCtrlCVItem.Name = "tabCtrlCVItem"
        Me.tabCtrlCVItem.SelectedIndex = 0
        Me.tabCtrlCVItem.Size = New System.Drawing.Size(855, 358)
        Me.tabCtrlCVItem.TabIndex = 5
        Me.tabCtrlCVItem.TabStop = false
        '
        'tabPageKizon110
        '
        Me.tabPageKizon110.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageKizon110.Controls.Add(Me.grpKizon1)
        Me.tabPageKizon110.Controls.Add(Me.lblSelectPageCnt1)
        Me.tabPageKizon110.Location = New System.Drawing.Point(4, 27)
        Me.tabPageKizon110.Name = "tabPageKizon110"
        Me.tabPageKizon110.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageKizon110.Size = New System.Drawing.Size(847, 327)
        Me.tabPageKizon110.TabIndex = 9
        Me.tabPageKizon110.Text = " 項目選択1"
        '
        'grpKizon1
        '
        Me.grpKizon1.Controls.Add(Me.pnlKiMstHendo)
        Me.grpKizon1.Controls.Add(Me.pnlKiMstTokuyaku)
        Me.grpKizon1.Controls.Add(Me.pnlKiMstKasyoClaimrui)
        Me.grpKizon1.Controls.Add(Me.pnlKiMstTitle)
        Me.grpKizon1.Controls.Add(Me.pnlKiMstArea)
        Me.grpKizon1.Controls.Add(Me.pnlKiMstSchool)
        Me.grpKizon1.Controls.Add(Me.pnlKiMstHokenrui)
        Me.grpKizon1.Controls.Add(Me.pnlKiMstBus)
        Me.grpKizon1.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpKizon1.Location = New System.Drawing.Point(6, 5)
        Me.grpKizon1.Name = "grpKizon1"
        Me.grpKizon1.Size = New System.Drawing.Size(835, 298)
        Me.grpKizon1.TabIndex = 0
        Me.grpKizon1.TabStop = false
        Me.grpKizon1.Text = "【各マスタ情報】"
        '
        'pnlKiMstHendo
        '
        Me.pnlKiMstHendo.Controls.Add(Me.chkKiMstHendo)
        Me.pnlKiMstHendo.Controls.Add(Me.lblKiMstHendoCnt)
        Me.pnlKiMstHendo.Controls.Add(Me.Label230)
        Me.pnlKiMstHendo.Controls.Add(Me.Label82)
        Me.pnlKiMstHendo.Controls.Add(Me.Label81)
        Me.pnlKiMstHendo.Location = New System.Drawing.Point(557, 20)
        Me.pnlKiMstHendo.Name = "pnlKiMstHendo"
        Me.pnlKiMstHendo.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiMstHendo.TabIndex = 50
        '
        'chkKiMstHendo
        '
        Me.chkKiMstHendo.AutoSize = true
        Me.chkKiMstHendo.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiMstHendo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkKiMstHendo.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstHendo.Name = "chkKiMstHendo"
        Me.chkKiMstHendo.Size = New System.Drawing.Size(134, 27)
        Me.chkKiMstHendo.TabIndex = 30
        Me.chkKiMstHendo.Text = "変動費設定内容"
        Me.chkKiMstHendo.UseVisualStyleBackColor = true
        '
        'lblKiMstHendoCnt
        '
        Me.lblKiMstHendoCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstHendoCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiMstHendoCnt.Name = "lblKiMstHendoCnt"
        Me.lblKiMstHendoCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstHendoCnt.TabIndex = 33
        Me.lblKiMstHendoCnt.Text = "9,999,999"
        Me.lblKiMstHendoCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label230
        '
        Me.Label230.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label230.Location = New System.Drawing.Point(19, 54)
        Me.Label230.Name = "Label230"
        Me.Label230.Size = New System.Drawing.Size(92, 16)
        Me.Label230.TabIndex = 32
        Me.Label230.Text = "( 参考件数 = "
        Me.Label230.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label82
        '
        Me.Label82.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label82.Location = New System.Drawing.Point(195, 54)
        Me.Label82.Name = "Label82"
        Me.Label82.Size = New System.Drawing.Size(22, 16)
        Me.Label82.TabIndex = 34
        Me.Label82.Text = " )"
        Me.Label82.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label81
        '
        Me.Label81.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label81.Location = New System.Drawing.Point(19, 33)
        Me.Label81.Name = "Label81"
        Me.Label81.Size = New System.Drawing.Size(198, 16)
        Me.Label81.TabIndex = 31
        Me.Label81.Text = "口径や料金表などの設定"
        Me.Label81.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiMstTokuyaku
        '
        Me.pnlKiMstTokuyaku.Controls.Add(Me.chkKiMstTokuyaku)
        Me.pnlKiMstTokuyaku.Controls.Add(Me.lblKiMstTokuyakuCnt)
        Me.pnlKiMstTokuyaku.Controls.Add(Me.Label223)
        Me.pnlKiMstTokuyaku.Controls.Add(Me.Label222)
        Me.pnlKiMstTokuyaku.Controls.Add(Me.Label78)
        Me.pnlKiMstTokuyaku.Location = New System.Drawing.Point(286, 112)
        Me.pnlKiMstTokuyaku.Name = "pnlKiMstTokuyaku"
        Me.pnlKiMstTokuyaku.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiMstTokuyaku.TabIndex = 49
        '
        'chkKiMstTokuyaku
        '
        Me.chkKiMstTokuyaku.AutoSize = true
        Me.chkKiMstTokuyaku.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiMstTokuyaku.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstTokuyaku.Name = "chkKiMstTokuyaku"
        Me.chkKiMstTokuyaku.Size = New System.Drawing.Size(104, 27)
        Me.chkKiMstTokuyaku.TabIndex = 20
        Me.chkKiMstTokuyaku.Text = "特約マスタ"
        Me.chkKiMstTokuyaku.UseVisualStyleBackColor = true
        '
        'lblKiMstTokuyakuCnt
        '
        Me.lblKiMstTokuyakuCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstTokuyakuCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiMstTokuyakuCnt.Name = "lblKiMstTokuyakuCnt"
        Me.lblKiMstTokuyakuCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstTokuyakuCnt.TabIndex = 23
        Me.lblKiMstTokuyakuCnt.Text = "9,999,999"
        Me.lblKiMstTokuyakuCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label223
        '
        Me.Label223.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label223.Location = New System.Drawing.Point(19, 54)
        Me.Label223.Name = "Label223"
        Me.Label223.Size = New System.Drawing.Size(92, 16)
        Me.Label223.TabIndex = 22
        Me.Label223.Text = "( 参考件数 = "
        Me.Label223.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label222
        '
        Me.Label222.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label222.Location = New System.Drawing.Point(195, 54)
        Me.Label222.Name = "Label222"
        Me.Label222.Size = New System.Drawing.Size(22, 16)
        Me.Label222.TabIndex = 24
        Me.Label222.Text = " )"
        Me.Label222.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label78
        '
        Me.Label78.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label78.Location = New System.Drawing.Point(19, 33)
        Me.Label78.Name = "Label78"
        Me.Label78.Size = New System.Drawing.Size(198, 16)
        Me.Label78.TabIndex = 21
        Me.Label78.Text = "原状回復特約, 修繕特約, 特約事項"
        Me.Label78.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiMstKasyoClaimrui
        '
        Me.pnlKiMstKasyoClaimrui.Controls.Add(Me.chkKiMstKasyoClaimrui)
        Me.pnlKiMstKasyoClaimrui.Controls.Add(Me.lblKiMstKasyoClaimruiCnt)
        Me.pnlKiMstKasyoClaimrui.Controls.Add(Me.Label228)
        Me.pnlKiMstKasyoClaimrui.Controls.Add(Me.Label227)
        Me.pnlKiMstKasyoClaimrui.Controls.Add(Me.Label226)
        Me.pnlKiMstKasyoClaimrui.Location = New System.Drawing.Point(286, 204)
        Me.pnlKiMstKasyoClaimrui.Name = "pnlKiMstKasyoClaimrui"
        Me.pnlKiMstKasyoClaimrui.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiMstKasyoClaimrui.TabIndex = 49
        '
        'chkKiMstKasyoClaimrui
        '
        Me.chkKiMstKasyoClaimrui.AutoSize = true
        Me.chkKiMstKasyoClaimrui.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiMstKasyoClaimrui.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstKasyoClaimrui.Name = "chkKiMstKasyoClaimrui"
        Me.chkKiMstKasyoClaimrui.Size = New System.Drawing.Size(179, 27)
        Me.chkKiMstKasyoClaimrui.TabIndex = 25
        Me.chkKiMstKasyoClaimrui.Text = "クレーム分類設定内容"
        Me.chkKiMstKasyoClaimrui.UseVisualStyleBackColor = true
        '
        'lblKiMstKasyoClaimruiCnt
        '
        Me.lblKiMstKasyoClaimruiCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstKasyoClaimruiCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiMstKasyoClaimruiCnt.Name = "lblKiMstKasyoClaimruiCnt"
        Me.lblKiMstKasyoClaimruiCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstKasyoClaimruiCnt.TabIndex = 28
        Me.lblKiMstKasyoClaimruiCnt.Text = "9,999,999"
        Me.lblKiMstKasyoClaimruiCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label228
        '
        Me.Label228.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label228.Location = New System.Drawing.Point(19, 54)
        Me.Label228.Name = "Label228"
        Me.Label228.Size = New System.Drawing.Size(92, 16)
        Me.Label228.TabIndex = 27
        Me.Label228.Text = "( 参考件数 = "
        Me.Label228.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label227
        '
        Me.Label227.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label227.Location = New System.Drawing.Point(195, 54)
        Me.Label227.Name = "Label227"
        Me.Label227.Size = New System.Drawing.Size(22, 16)
        Me.Label227.TabIndex = 29
        Me.Label227.Text = " )"
        Me.Label227.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label226
        '
        Me.Label226.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label226.Location = New System.Drawing.Point(19, 33)
        Me.Label226.Name = "Label226"
        Me.Label226.Size = New System.Drawing.Size(198, 16)
        Me.Label226.TabIndex = 26
        Me.Label226.Text = "箇所分類(例：外溝), クレーム分類"
        Me.Label226.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiMstTitle
        '
        Me.pnlKiMstTitle.Controls.Add(Me.Label346)
        Me.pnlKiMstTitle.Controls.Add(Me.Label236)
        Me.pnlKiMstTitle.Controls.Add(Me.Label233)
        Me.pnlKiMstTitle.Controls.Add(Me.lblKiMstGazotitleCnt)
        Me.pnlKiMstTitle.Controls.Add(Me.Label347)
        Me.pnlKiMstTitle.Controls.Add(Me.lblKiMstBikotitleCnt)
        Me.pnlKiMstTitle.Controls.Add(Me.Label235)
        Me.pnlKiMstTitle.Controls.Add(Me.chkKiMstTitle)
        Me.pnlKiMstTitle.Controls.Add(Me.lblKiMstKagititleCnt)
        Me.pnlKiMstTitle.Controls.Add(Me.Label232)
        Me.pnlKiMstTitle.Controls.Add(Me.Label80)
        Me.pnlKiMstTitle.Location = New System.Drawing.Point(557, 112)
        Me.pnlKiMstTitle.Name = "pnlKiMstTitle"
        Me.pnlKiMstTitle.Size = New System.Drawing.Size(265, 120)
        Me.pnlKiMstTitle.TabIndex = 49
        '
        'Label346
        '
        Me.Label346.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label346.Location = New System.Drawing.Point(19, 93)
        Me.Label346.Name = "Label346"
        Me.Label346.Size = New System.Drawing.Size(120, 16)
        Me.Label346.TabIndex = 48
        Me.Label346.Text = "( 画像 参考件数 = "
        Me.Label346.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label236
        '
        Me.Label236.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label236.Location = New System.Drawing.Point(19, 73)
        Me.Label236.Name = "Label236"
        Me.Label236.Size = New System.Drawing.Size(120, 16)
        Me.Label236.TabIndex = 45
        Me.Label236.Text = "( 備考 参考件数 = "
        Me.Label236.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label233
        '
        Me.Label233.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label233.Location = New System.Drawing.Point(19, 54)
        Me.Label233.Name = "Label233"
        Me.Label233.Size = New System.Drawing.Size(120, 16)
        Me.Label233.TabIndex = 37
        Me.Label233.Text = "( 鍵    参考件数 = "
        Me.Label233.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblKiMstGazotitleCnt
        '
        Me.lblKiMstGazotitleCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstGazotitleCnt.Location = New System.Drawing.Point(124, 93)
        Me.lblKiMstGazotitleCnt.Name = "lblKiMstGazotitleCnt"
        Me.lblKiMstGazotitleCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstGazotitleCnt.TabIndex = 49
        Me.lblKiMstGazotitleCnt.Text = "9,999,999"
        Me.lblKiMstGazotitleCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label347
        '
        Me.Label347.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label347.Location = New System.Drawing.Point(213, 93)
        Me.Label347.Name = "Label347"
        Me.Label347.Size = New System.Drawing.Size(22, 16)
        Me.Label347.TabIndex = 50
        Me.Label347.Text = " )"
        Me.Label347.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblKiMstBikotitleCnt
        '
        Me.lblKiMstBikotitleCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstBikotitleCnt.Location = New System.Drawing.Point(124, 73)
        Me.lblKiMstBikotitleCnt.Name = "lblKiMstBikotitleCnt"
        Me.lblKiMstBikotitleCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstBikotitleCnt.TabIndex = 46
        Me.lblKiMstBikotitleCnt.Text = "9,999,999"
        Me.lblKiMstBikotitleCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label235
        '
        Me.Label235.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label235.Location = New System.Drawing.Point(213, 73)
        Me.Label235.Name = "Label235"
        Me.Label235.Size = New System.Drawing.Size(22, 16)
        Me.Label235.TabIndex = 47
        Me.Label235.Text = " )"
        Me.Label235.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkKiMstTitle
        '
        Me.chkKiMstTitle.AutoSize = true
        Me.chkKiMstTitle.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiMstTitle.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstTitle.Name = "chkKiMstTitle"
        Me.chkKiMstTitle.Size = New System.Drawing.Size(134, 27)
        Me.chkKiMstTitle.TabIndex = 35
        Me.chkKiMstTitle.Text = "タイトルマスタ"
        Me.chkKiMstTitle.UseVisualStyleBackColor = true
        '
        'lblKiMstKagititleCnt
        '
        Me.lblKiMstKagititleCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstKagititleCnt.Location = New System.Drawing.Point(124, 54)
        Me.lblKiMstKagititleCnt.Name = "lblKiMstKagititleCnt"
        Me.lblKiMstKagititleCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstKagititleCnt.TabIndex = 38
        Me.lblKiMstKagititleCnt.Text = "9,999,999"
        Me.lblKiMstKagititleCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label232
        '
        Me.Label232.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label232.Location = New System.Drawing.Point(213, 54)
        Me.Label232.Name = "Label232"
        Me.Label232.Size = New System.Drawing.Size(22, 16)
        Me.Label232.TabIndex = 39
        Me.Label232.Text = " )"
        Me.Label232.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label80
        '
        Me.Label80.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label80.Location = New System.Drawing.Point(19, 33)
        Me.Label80.Name = "Label80"
        Me.Label80.Size = New System.Drawing.Size(198, 16)
        Me.Label80.TabIndex = 36
        Me.Label80.Text = "鍵、備考、画像タイトル"
        Me.Label80.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiMstArea
        '
        Me.pnlKiMstArea.Controls.Add(Me.chkKiMstArea)
        Me.pnlKiMstArea.Controls.Add(Me.lblKiMstAreaCnt)
        Me.pnlKiMstArea.Controls.Add(Me.Label216)
        Me.pnlKiMstArea.Controls.Add(Me.Label77)
        Me.pnlKiMstArea.Controls.Add(Me.Label76)
        Me.pnlKiMstArea.Location = New System.Drawing.Point(15, 204)
        Me.pnlKiMstArea.Name = "pnlKiMstArea"
        Me.pnlKiMstArea.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiMstArea.TabIndex = 49
        '
        'chkKiMstArea
        '
        Me.chkKiMstArea.AutoSize = true
        Me.chkKiMstArea.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiMstArea.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstArea.Name = "chkKiMstArea"
        Me.chkKiMstArea.Size = New System.Drawing.Size(119, 27)
        Me.chkKiMstArea.TabIndex = 10
        Me.chkKiMstArea.Text = "エリアマスタ"
        Me.chkKiMstArea.UseVisualStyleBackColor = true
        '
        'lblKiMstAreaCnt
        '
        Me.lblKiMstAreaCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstAreaCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiMstAreaCnt.Name = "lblKiMstAreaCnt"
        Me.lblKiMstAreaCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstAreaCnt.TabIndex = 13
        Me.lblKiMstAreaCnt.Text = "9,999,999"
        Me.lblKiMstAreaCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label216
        '
        Me.Label216.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label216.Location = New System.Drawing.Point(19, 54)
        Me.Label216.Name = "Label216"
        Me.Label216.Size = New System.Drawing.Size(92, 16)
        Me.Label216.TabIndex = 12
        Me.Label216.Text = "( 参考件数 = "
        Me.Label216.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label77
        '
        Me.Label77.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label77.Location = New System.Drawing.Point(195, 54)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(22, 16)
        Me.Label77.TabIndex = 14
        Me.Label77.Text = " )"
        Me.Label77.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label76
        '
        Me.Label76.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label76.Location = New System.Drawing.Point(19, 33)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(198, 16)
        Me.Label76.TabIndex = 11
        Me.Label76.Text = "例：東部, 北部"
        Me.Label76.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiMstSchool
        '
        Me.pnlKiMstSchool.Controls.Add(Me.chkKiMstSchool)
        Me.pnlKiMstSchool.Controls.Add(Me.lblKiMstSchoolCnt)
        Me.pnlKiMstSchool.Controls.Add(Me.Label214)
        Me.pnlKiMstSchool.Controls.Add(Me.Label213)
        Me.pnlKiMstSchool.Controls.Add(Me.Label75)
        Me.pnlKiMstSchool.Location = New System.Drawing.Point(15, 112)
        Me.pnlKiMstSchool.Name = "pnlKiMstSchool"
        Me.pnlKiMstSchool.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiMstSchool.TabIndex = 49
        '
        'chkKiMstSchool
        '
        Me.chkKiMstSchool.AutoSize = true
        Me.chkKiMstSchool.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiMstSchool.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkKiMstSchool.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstSchool.Name = "chkKiMstSchool"
        Me.chkKiMstSchool.Size = New System.Drawing.Size(119, 27)
        Me.chkKiMstSchool.TabIndex = 5
        Me.chkKiMstSchool.Text = "学校区マスタ"
        Me.chkKiMstSchool.UseVisualStyleBackColor = true
        '
        'lblKiMstSchoolCnt
        '
        Me.lblKiMstSchoolCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstSchoolCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiMstSchoolCnt.Name = "lblKiMstSchoolCnt"
        Me.lblKiMstSchoolCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstSchoolCnt.TabIndex = 8
        Me.lblKiMstSchoolCnt.Text = "9,999,999"
        Me.lblKiMstSchoolCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label214
        '
        Me.Label214.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label214.Location = New System.Drawing.Point(19, 54)
        Me.Label214.Name = "Label214"
        Me.Label214.Size = New System.Drawing.Size(92, 16)
        Me.Label214.TabIndex = 7
        Me.Label214.Text = "( 参考件数 = "
        Me.Label214.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label213
        '
        Me.Label213.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label213.Location = New System.Drawing.Point(195, 54)
        Me.Label213.Name = "Label213"
        Me.Label213.Size = New System.Drawing.Size(22, 16)
        Me.Label213.TabIndex = 9
        Me.Label213.Text = " )"
        Me.Label213.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label75
        '
        Me.Label75.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label75.Location = New System.Drawing.Point(19, 33)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(198, 16)
        Me.Label75.TabIndex = 6
        Me.Label75.Text = "小学校, 中学校"
        Me.Label75.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiMstHokenrui
        '
        Me.pnlKiMstHokenrui.Controls.Add(Me.chkKiMstHokenrui)
        Me.pnlKiMstHokenrui.Controls.Add(Me.lblKiMstHokenruiCnt)
        Me.pnlKiMstHokenrui.Controls.Add(Me.Label220)
        Me.pnlKiMstHokenrui.Controls.Add(Me.Label219)
        Me.pnlKiMstHokenrui.Controls.Add(Me.Label218)
        Me.pnlKiMstHokenrui.Location = New System.Drawing.Point(286, 20)
        Me.pnlKiMstHokenrui.Name = "pnlKiMstHokenrui"
        Me.pnlKiMstHokenrui.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiMstHokenrui.TabIndex = 48
        '
        'chkKiMstHokenrui
        '
        Me.chkKiMstHokenrui.AutoSize = true
        Me.chkKiMstHokenrui.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiMstHokenrui.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstHokenrui.Name = "chkKiMstHokenrui"
        Me.chkKiMstHokenrui.Size = New System.Drawing.Size(134, 27)
        Me.chkKiMstHokenrui.TabIndex = 15
        Me.chkKiMstHokenrui.Text = "保険種類マスタ"
        Me.chkKiMstHokenrui.UseVisualStyleBackColor = true
        '
        'lblKiMstHokenruiCnt
        '
        Me.lblKiMstHokenruiCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiMstHokenruiCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiMstHokenruiCnt.Name = "lblKiMstHokenruiCnt"
        Me.lblKiMstHokenruiCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstHokenruiCnt.TabIndex = 18
        Me.lblKiMstHokenruiCnt.Text = "9,999,999"
        Me.lblKiMstHokenruiCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label220
        '
        Me.Label220.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label220.Location = New System.Drawing.Point(19, 54)
        Me.Label220.Name = "Label220"
        Me.Label220.Size = New System.Drawing.Size(92, 16)
        Me.Label220.TabIndex = 17
        Me.Label220.Text = "( 参考件数 = "
        Me.Label220.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label219
        '
        Me.Label219.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label219.Location = New System.Drawing.Point(195, 54)
        Me.Label219.Name = "Label219"
        Me.Label219.Size = New System.Drawing.Size(22, 16)
        Me.Label219.TabIndex = 19
        Me.Label219.Text = " )"
        Me.Label219.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label218
        '
        Me.Label218.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label218.Location = New System.Drawing.Point(19, 33)
        Me.Label218.Name = "Label218"
        Me.Label218.Size = New System.Drawing.Size(198, 16)
        Me.Label218.TabIndex = 16
        Me.Label218.Text = "例：住宅総合保険"
        Me.Label218.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiMstBus
        '
        Me.pnlKiMstBus.Controls.Add(Me.chkKiMstBus)
        Me.pnlKiMstBus.Controls.Add(Me.lblKiMstBusCnt)
        Me.pnlKiMstBus.Controls.Add(Me.Label211)
        Me.pnlKiMstBus.Controls.Add(Me.Label181)
        Me.pnlKiMstBus.Controls.Add(Me.Label210)
        Me.pnlKiMstBus.Location = New System.Drawing.Point(15, 20)
        Me.pnlKiMstBus.Name = "pnlKiMstBus"
        Me.pnlKiMstBus.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiMstBus.TabIndex = 47
        '
        'chkKiMstBus
        '
        Me.chkKiMstBus.AutoSize = true
        Me.chkKiMstBus.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiMstBus.Location = New System.Drawing.Point(3, 3)
        Me.chkKiMstBus.Name = "chkKiMstBus"
        Me.chkKiMstBus.Size = New System.Drawing.Size(134, 27)
        Me.chkKiMstBus.TabIndex = 0
        Me.chkKiMstBus.Text = "バス交通マスタ"
        Me.chkKiMstBus.UseVisualStyleBackColor = true
        '
        'lblKiMstBusCnt
        '
        Me.lblKiMstBusCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiMstBusCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiMstBusCnt.Name = "lblKiMstBusCnt"
        Me.lblKiMstBusCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiMstBusCnt.TabIndex = 3
        Me.lblKiMstBusCnt.Text = "9,999,999"
        Me.lblKiMstBusCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label211
        '
        Me.Label211.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label211.Location = New System.Drawing.Point(19, 54)
        Me.Label211.Name = "Label211"
        Me.Label211.Size = New System.Drawing.Size(92, 16)
        Me.Label211.TabIndex = 2
        Me.Label211.Text = "( 参考件数 = "
        Me.Label211.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label181
        '
        Me.Label181.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label181.Location = New System.Drawing.Point(195, 54)
        Me.Label181.Name = "Label181"
        Me.Label181.Size = New System.Drawing.Size(22, 16)
        Me.Label181.TabIndex = 4
        Me.Label181.Text = " )"
        Me.Label181.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label210
        '
        Me.Label210.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label210.Location = New System.Drawing.Point(19, 33)
        Me.Label210.Name = "Label210"
        Me.Label210.Size = New System.Drawing.Size(198, 16)
        Me.Label210.TabIndex = 1
        Me.Label210.Text = "バス会社, 系統, バス停"
        Me.Label210.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSelectPageCnt1
        '
        Me.lblSelectPageCnt1.AutoSize = true
        Me.lblSelectPageCnt1.Location = New System.Drawing.Point(806, 306)
        Me.lblSelectPageCnt1.Name = "lblSelectPageCnt1"
        Me.lblSelectPageCnt1.Size = New System.Drawing.Size(35, 18)
        Me.lblSelectPageCnt1.TabIndex = 46
        Me.lblSelectPageCnt1.Text = "1 / 3"
        '
        'tabPageKizon120
        '
        Me.tabPageKizon120.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageKizon120.Controls.Add(Me.grpKizon3)
        Me.tabPageKizon120.Controls.Add(Me.grpKizon5)
        Me.tabPageKizon120.Controls.Add(Me.lblSelectPageCnt2)
        Me.tabPageKizon120.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.tabPageKizon120.Location = New System.Drawing.Point(4, 27)
        Me.tabPageKizon120.Name = "tabPageKizon120"
        Me.tabPageKizon120.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageKizon120.Size = New System.Drawing.Size(847, 327)
        Me.tabPageKizon120.TabIndex = 10
        Me.tabPageKizon120.Text = " 項目選択2"
        '
        'grpKizon3
        '
        Me.grpKizon3.Controls.Add(Me.pnlKiGySyuzenBase)
        Me.grpKizon3.Controls.Add(Me.pnlKiGyYatinhosyoBase)
        Me.grpKizon3.Controls.Add(Me.pnlKiGyLifelineBase)
        Me.grpKizon3.Controls.Add(Me.pnlKiGyHokenBase)
        Me.grpKizon3.Controls.Add(Me.pnlKiGySisetuBase)
        Me.grpKizon3.Controls.Add(Me.pnlKiGyCyukaiBase)
        Me.grpKizon3.Controls.Add(Me.pnlKiGySekoBase)
        Me.grpKizon3.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpKizon3.Location = New System.Drawing.Point(304, 5)
        Me.grpKizon3.Name = "grpKizon3"
        Me.grpKizon3.Size = New System.Drawing.Size(537, 298)
        Me.grpKizon3.TabIndex = 1
        Me.grpKizon3.TabStop = false
        Me.grpKizon3.Text = "【業者情報】"
        '
        'pnlKiGySyuzenBase
        '
        Me.pnlKiGySyuzenBase.Controls.Add(Me.chkKiGySyuzenBase)
        Me.pnlKiGySyuzenBase.Controls.Add(Me.lblKiGySyuzenBaseCnt)
        Me.pnlKiGySyuzenBase.Controls.Add(Me.Label241)
        Me.pnlKiGySyuzenBase.Controls.Add(Me.Label240)
        Me.pnlKiGySyuzenBase.Location = New System.Drawing.Point(15, 89)
        Me.pnlKiGySyuzenBase.Name = "pnlKiGySyuzenBase"
        Me.pnlKiGySyuzenBase.Size = New System.Drawing.Size(255, 60)
        Me.pnlKiGySyuzenBase.TabIndex = 19
        '
        'chkKiGySyuzenBase
        '
        Me.chkKiGySyuzenBase.AutoSize = true
        Me.chkKiGySyuzenBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiGySyuzenBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiGySyuzenBase.Name = "chkKiGySyuzenBase"
        Me.chkKiGySyuzenBase.Size = New System.Drawing.Size(119, 27)
        Me.chkKiGySyuzenBase.TabIndex = 4
        Me.chkKiGySyuzenBase.Text = "修繕業者情報"
        Me.chkKiGySyuzenBase.UseVisualStyleBackColor = true
        '
        'lblKiGySyuzenBaseCnt
        '
        Me.lblKiGySyuzenBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiGySyuzenBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiGySyuzenBaseCnt.Name = "lblKiGySyuzenBaseCnt"
        Me.lblKiGySyuzenBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiGySyuzenBaseCnt.TabIndex = 6
        Me.lblKiGySyuzenBaseCnt.Text = "9,999,999"
        Me.lblKiGySyuzenBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label241
        '
        Me.Label241.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label241.Location = New System.Drawing.Point(19, 33)
        Me.Label241.Name = "Label241"
        Me.Label241.Size = New System.Drawing.Size(92, 16)
        Me.Label241.TabIndex = 5
        Me.Label241.Text = "( 参考件数 = "
        Me.Label241.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label240
        '
        Me.Label240.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label240.Location = New System.Drawing.Point(195, 33)
        Me.Label240.Name = "Label240"
        Me.Label240.Size = New System.Drawing.Size(22, 16)
        Me.Label240.TabIndex = 7
        Me.Label240.Text = " )"
        Me.Label240.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiGyYatinhosyoBase
        '
        Me.pnlKiGyYatinhosyoBase.Controls.Add(Me.chkKiGyYatinhosyoBase)
        Me.pnlKiGyYatinhosyoBase.Controls.Add(Me.lblKiGyYatinhosyoBaseCnt)
        Me.pnlKiGyYatinhosyoBase.Controls.Add(Me.Label250)
        Me.pnlKiGyYatinhosyoBase.Controls.Add(Me.Label249)
        Me.pnlKiGyYatinhosyoBase.Location = New System.Drawing.Point(276, 20)
        Me.pnlKiGyYatinhosyoBase.Name = "pnlKiGyYatinhosyoBase"
        Me.pnlKiGyYatinhosyoBase.Size = New System.Drawing.Size(255, 60)
        Me.pnlKiGyYatinhosyoBase.TabIndex = 19
        '
        'chkKiGyYatinhosyoBase
        '
        Me.chkKiGyYatinhosyoBase.AutoSize = true
        Me.chkKiGyYatinhosyoBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiGyYatinhosyoBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiGyYatinhosyoBase.Name = "chkKiGyYatinhosyoBase"
        Me.chkKiGyYatinhosyoBase.Size = New System.Drawing.Size(149, 27)
        Me.chkKiGyYatinhosyoBase.TabIndex = 16
        Me.chkKiGyYatinhosyoBase.Text = "家賃保証業者情報"
        Me.chkKiGyYatinhosyoBase.UseVisualStyleBackColor = true
        '
        'lblKiGyYatinhosyoBaseCnt
        '
        Me.lblKiGyYatinhosyoBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiGyYatinhosyoBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiGyYatinhosyoBaseCnt.Name = "lblKiGyYatinhosyoBaseCnt"
        Me.lblKiGyYatinhosyoBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiGyYatinhosyoBaseCnt.TabIndex = 18
        Me.lblKiGyYatinhosyoBaseCnt.Text = "9,999,999"
        Me.lblKiGyYatinhosyoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label250
        '
        Me.Label250.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label250.Location = New System.Drawing.Point(19, 33)
        Me.Label250.Name = "Label250"
        Me.Label250.Size = New System.Drawing.Size(92, 16)
        Me.Label250.TabIndex = 17
        Me.Label250.Text = "( 参考件数 = "
        Me.Label250.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label249
        '
        Me.Label249.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label249.Location = New System.Drawing.Point(195, 33)
        Me.Label249.Name = "Label249"
        Me.Label249.Size = New System.Drawing.Size(22, 16)
        Me.Label249.TabIndex = 19
        Me.Label249.Text = " )"
        Me.Label249.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiGyLifelineBase
        '
        Me.pnlKiGyLifelineBase.Controls.Add(Me.chkKiGyLifelineBase)
        Me.pnlKiGyLifelineBase.Controls.Add(Me.lblKiGyLifelineBaseCnt)
        Me.pnlKiGyLifelineBase.Controls.Add(Me.Label244)
        Me.pnlKiGyLifelineBase.Controls.Add(Me.Label243)
        Me.pnlKiGyLifelineBase.Location = New System.Drawing.Point(15, 158)
        Me.pnlKiGyLifelineBase.Name = "pnlKiGyLifelineBase"
        Me.pnlKiGyLifelineBase.Size = New System.Drawing.Size(255, 60)
        Me.pnlKiGyLifelineBase.TabIndex = 19
        '
        'chkKiGyLifelineBase
        '
        Me.chkKiGyLifelineBase.AutoSize = true
        Me.chkKiGyLifelineBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiGyLifelineBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiGyLifelineBase.Name = "chkKiGyLifelineBase"
        Me.chkKiGyLifelineBase.Size = New System.Drawing.Size(179, 27)
        Me.chkKiGyLifelineBase.TabIndex = 8
        Me.chkKiGyLifelineBase.Text = "ライフライン業者情報"
        Me.chkKiGyLifelineBase.UseVisualStyleBackColor = true
        '
        'lblKiGyLifelineBaseCnt
        '
        Me.lblKiGyLifelineBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiGyLifelineBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiGyLifelineBaseCnt.Name = "lblKiGyLifelineBaseCnt"
        Me.lblKiGyLifelineBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiGyLifelineBaseCnt.TabIndex = 10
        Me.lblKiGyLifelineBaseCnt.Text = "9,999,999"
        Me.lblKiGyLifelineBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label244
        '
        Me.Label244.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label244.Location = New System.Drawing.Point(19, 33)
        Me.Label244.Name = "Label244"
        Me.Label244.Size = New System.Drawing.Size(92, 16)
        Me.Label244.TabIndex = 9
        Me.Label244.Text = "( 参考件数 = "
        Me.Label244.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label243
        '
        Me.Label243.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label243.Location = New System.Drawing.Point(195, 33)
        Me.Label243.Name = "Label243"
        Me.Label243.Size = New System.Drawing.Size(22, 16)
        Me.Label243.TabIndex = 11
        Me.Label243.Text = " )"
        Me.Label243.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiGyHokenBase
        '
        Me.pnlKiGyHokenBase.Controls.Add(Me.chkKiGyHokenBase)
        Me.pnlKiGyHokenBase.Controls.Add(Me.lblKiGyHokenBaseCnt)
        Me.pnlKiGyHokenBase.Controls.Add(Me.Label247)
        Me.pnlKiGyHokenBase.Controls.Add(Me.Label246)
        Me.pnlKiGyHokenBase.Location = New System.Drawing.Point(15, 227)
        Me.pnlKiGyHokenBase.Name = "pnlKiGyHokenBase"
        Me.pnlKiGyHokenBase.Size = New System.Drawing.Size(255, 60)
        Me.pnlKiGyHokenBase.TabIndex = 19
        '
        'chkKiGyHokenBase
        '
        Me.chkKiGyHokenBase.AutoSize = true
        Me.chkKiGyHokenBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiGyHokenBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiGyHokenBase.Name = "chkKiGyHokenBase"
        Me.chkKiGyHokenBase.Size = New System.Drawing.Size(119, 27)
        Me.chkKiGyHokenBase.TabIndex = 12
        Me.chkKiGyHokenBase.Text = "保険業者情報"
        Me.chkKiGyHokenBase.UseVisualStyleBackColor = true
        '
        'lblKiGyHokenBaseCnt
        '
        Me.lblKiGyHokenBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiGyHokenBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiGyHokenBaseCnt.Name = "lblKiGyHokenBaseCnt"
        Me.lblKiGyHokenBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiGyHokenBaseCnt.TabIndex = 14
        Me.lblKiGyHokenBaseCnt.Text = "9,999,999"
        Me.lblKiGyHokenBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label247
        '
        Me.Label247.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label247.Location = New System.Drawing.Point(19, 33)
        Me.Label247.Name = "Label247"
        Me.Label247.Size = New System.Drawing.Size(92, 16)
        Me.Label247.TabIndex = 13
        Me.Label247.Text = "( 参考件数 = "
        Me.Label247.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label246
        '
        Me.Label246.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label246.Location = New System.Drawing.Point(195, 33)
        Me.Label246.Name = "Label246"
        Me.Label246.Size = New System.Drawing.Size(22, 16)
        Me.Label246.TabIndex = 15
        Me.Label246.Text = " )"
        Me.Label246.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiGySisetuBase
        '
        Me.pnlKiGySisetuBase.Controls.Add(Me.chkKiGySisetuBase)
        Me.pnlKiGySisetuBase.Controls.Add(Me.lblKiGySisetuBaseCnt)
        Me.pnlKiGySisetuBase.Controls.Add(Me.Label253)
        Me.pnlKiGySisetuBase.Controls.Add(Me.Label252)
        Me.pnlKiGySisetuBase.Location = New System.Drawing.Point(276, 89)
        Me.pnlKiGySisetuBase.Name = "pnlKiGySisetuBase"
        Me.pnlKiGySisetuBase.Size = New System.Drawing.Size(255, 60)
        Me.pnlKiGySisetuBase.TabIndex = 19
        '
        'chkKiGySisetuBase
        '
        Me.chkKiGySisetuBase.AutoSize = true
        Me.chkKiGySisetuBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiGySisetuBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiGySisetuBase.Name = "chkKiGySisetuBase"
        Me.chkKiGySisetuBase.Size = New System.Drawing.Size(149, 27)
        Me.chkKiGySisetuBase.TabIndex = 20
        Me.chkKiGySisetuBase.Text = "施設保守業者情報"
        Me.chkKiGySisetuBase.UseVisualStyleBackColor = true
        '
        'lblKiGySisetuBaseCnt
        '
        Me.lblKiGySisetuBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiGySisetuBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiGySisetuBaseCnt.Name = "lblKiGySisetuBaseCnt"
        Me.lblKiGySisetuBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiGySisetuBaseCnt.TabIndex = 22
        Me.lblKiGySisetuBaseCnt.Text = "9,999,999"
        Me.lblKiGySisetuBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label253
        '
        Me.Label253.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label253.Location = New System.Drawing.Point(19, 33)
        Me.Label253.Name = "Label253"
        Me.Label253.Size = New System.Drawing.Size(92, 16)
        Me.Label253.TabIndex = 21
        Me.Label253.Text = "( 参考件数 = "
        Me.Label253.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label252
        '
        Me.Label252.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label252.Location = New System.Drawing.Point(195, 33)
        Me.Label252.Name = "Label252"
        Me.Label252.Size = New System.Drawing.Size(22, 16)
        Me.Label252.TabIndex = 23
        Me.Label252.Text = " )"
        Me.Label252.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiGyCyukaiBase
        '
        Me.pnlKiGyCyukaiBase.Controls.Add(Me.chkKiGyCyukaiBase)
        Me.pnlKiGyCyukaiBase.Controls.Add(Me.lblKiGyCyukaiBaseCnt)
        Me.pnlKiGyCyukaiBase.Controls.Add(Me.Label238)
        Me.pnlKiGyCyukaiBase.Controls.Add(Me.Label87)
        Me.pnlKiGyCyukaiBase.Location = New System.Drawing.Point(15, 20)
        Me.pnlKiGyCyukaiBase.Name = "pnlKiGyCyukaiBase"
        Me.pnlKiGyCyukaiBase.Size = New System.Drawing.Size(255, 60)
        Me.pnlKiGyCyukaiBase.TabIndex = 19
        '
        'chkKiGyCyukaiBase
        '
        Me.chkKiGyCyukaiBase.AutoSize = true
        Me.chkKiGyCyukaiBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiGyCyukaiBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiGyCyukaiBase.Name = "chkKiGyCyukaiBase"
        Me.chkKiGyCyukaiBase.Size = New System.Drawing.Size(164, 27)
        Me.chkKiGyCyukaiBase.TabIndex = 0
        Me.chkKiGyCyukaiBase.Text = "仲介・管理業者情報"
        Me.chkKiGyCyukaiBase.UseVisualStyleBackColor = true
        '
        'lblKiGyCyukaiBaseCnt
        '
        Me.lblKiGyCyukaiBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiGyCyukaiBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiGyCyukaiBaseCnt.Name = "lblKiGyCyukaiBaseCnt"
        Me.lblKiGyCyukaiBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiGyCyukaiBaseCnt.TabIndex = 2
        Me.lblKiGyCyukaiBaseCnt.Text = "9,999,999"
        Me.lblKiGyCyukaiBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label238
        '
        Me.Label238.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label238.Location = New System.Drawing.Point(19, 33)
        Me.Label238.Name = "Label238"
        Me.Label238.Size = New System.Drawing.Size(92, 16)
        Me.Label238.TabIndex = 1
        Me.Label238.Text = "( 参考件数 = "
        Me.Label238.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label87
        '
        Me.Label87.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label87.Location = New System.Drawing.Point(195, 33)
        Me.Label87.Name = "Label87"
        Me.Label87.Size = New System.Drawing.Size(22, 16)
        Me.Label87.TabIndex = 3
        Me.Label87.Text = " )"
        Me.Label87.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiGySekoBase
        '
        Me.pnlKiGySekoBase.Controls.Add(Me.chkKiGySekoBase)
        Me.pnlKiGySekoBase.Controls.Add(Me.lblKiGySekoBaseCnt)
        Me.pnlKiGySekoBase.Controls.Add(Me.Label256)
        Me.pnlKiGySekoBase.Controls.Add(Me.Label255)
        Me.pnlKiGySekoBase.Location = New System.Drawing.Point(276, 158)
        Me.pnlKiGySekoBase.Name = "pnlKiGySekoBase"
        Me.pnlKiGySekoBase.Size = New System.Drawing.Size(255, 60)
        Me.pnlKiGySekoBase.TabIndex = 19
        '
        'chkKiGySekoBase
        '
        Me.chkKiGySekoBase.AutoSize = true
        Me.chkKiGySekoBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiGySekoBase.Location = New System.Drawing.Point(6, 3)
        Me.chkKiGySekoBase.Name = "chkKiGySekoBase"
        Me.chkKiGySekoBase.Size = New System.Drawing.Size(119, 27)
        Me.chkKiGySekoBase.TabIndex = 24
        Me.chkKiGySekoBase.Text = "施工業者情報"
        Me.chkKiGySekoBase.UseVisualStyleBackColor = true
        '
        'lblKiGySekoBaseCnt
        '
        Me.lblKiGySekoBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiGySekoBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiGySekoBaseCnt.Name = "lblKiGySekoBaseCnt"
        Me.lblKiGySekoBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiGySekoBaseCnt.TabIndex = 26
        Me.lblKiGySekoBaseCnt.Text = "9,999,999"
        Me.lblKiGySekoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label256
        '
        Me.Label256.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label256.Location = New System.Drawing.Point(19, 33)
        Me.Label256.Name = "Label256"
        Me.Label256.Size = New System.Drawing.Size(92, 16)
        Me.Label256.TabIndex = 25
        Me.Label256.Text = "( 参考件数 = "
        Me.Label256.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label255
        '
        Me.Label255.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label255.Location = New System.Drawing.Point(195, 33)
        Me.Label255.Name = "Label255"
        Me.Label255.Size = New System.Drawing.Size(22, 16)
        Me.Label255.TabIndex = 27
        Me.Label255.Text = " )"
        Me.Label255.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpKizon5
        '
        Me.grpKizon5.Controls.Add(Me.pnlKiOw)
        Me.grpKizon5.Controls.Add(Me.pnlKiJisya)
        Me.grpKizon5.Controls.Add(Me.pnlKiSyskanriBase)
        Me.grpKizon5.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpKizon5.Location = New System.Drawing.Point(6, 5)
        Me.grpKizon5.Name = "grpKizon5"
        Me.grpKizon5.Size = New System.Drawing.Size(292, 298)
        Me.grpKizon5.TabIndex = 0
        Me.grpKizon5.TabStop = false
        Me.grpKizon5.Text = "【基本情報1】"
        '
        'pnlKiOw
        '
        Me.pnlKiOw.Controls.Add(Me.chkKiOwBase)
        Me.pnlKiOw.Controls.Add(Me.lblKiOwBaseCnt)
        Me.pnlKiOw.Controls.Add(Me.Label261)
        Me.pnlKiOw.Controls.Add(Me.Label268)
        Me.pnlKiOw.Controls.Add(Me.Label267)
        Me.pnlKiOw.Location = New System.Drawing.Point(15, 204)
        Me.pnlKiOw.Name = "pnlKiOw"
        Me.pnlKiOw.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiOw.TabIndex = 17
        '
        'chkKiOwBase
        '
        Me.chkKiOwBase.AutoSize = true
        Me.chkKiOwBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiOwBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiOwBase.Name = "chkKiOwBase"
        Me.chkKiOwBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiOwBase.TabIndex = 10
        Me.chkKiOwBase.Text = "家主情報"
        Me.chkKiOwBase.UseVisualStyleBackColor = true
        '
        'lblKiOwBaseCnt
        '
        Me.lblKiOwBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiOwBaseCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiOwBaseCnt.Name = "lblKiOwBaseCnt"
        Me.lblKiOwBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiOwBaseCnt.TabIndex = 13
        Me.lblKiOwBaseCnt.Text = "9,999,999"
        Me.lblKiOwBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label261
        '
        Me.Label261.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label261.Location = New System.Drawing.Point(19, 33)
        Me.Label261.Name = "Label261"
        Me.Label261.Size = New System.Drawing.Size(198, 16)
        Me.Label261.TabIndex = 11
        Me.Label261.Text = "家主情報 口座"
        Me.Label261.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label268
        '
        Me.Label268.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label268.Location = New System.Drawing.Point(19, 54)
        Me.Label268.Name = "Label268"
        Me.Label268.Size = New System.Drawing.Size(92, 16)
        Me.Label268.TabIndex = 12
        Me.Label268.Text = "( 参考件数 = "
        Me.Label268.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label267
        '
        Me.Label267.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label267.Location = New System.Drawing.Point(195, 54)
        Me.Label267.Name = "Label267"
        Me.Label267.Size = New System.Drawing.Size(22, 16)
        Me.Label267.TabIndex = 14
        Me.Label267.Text = " )"
        Me.Label267.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiJisya
        '
        Me.pnlKiJisya.Controls.Add(Me.chkKiJisyaBase)
        Me.pnlKiJisya.Controls.Add(Me.lblKiJisyaBaseCnt)
        Me.pnlKiJisya.Controls.Add(Me.Label265)
        Me.pnlKiJisya.Controls.Add(Me.Label264)
        Me.pnlKiJisya.Controls.Add(Me.Label263)
        Me.pnlKiJisya.Location = New System.Drawing.Point(15, 112)
        Me.pnlKiJisya.Name = "pnlKiJisya"
        Me.pnlKiJisya.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiJisya.TabIndex = 16
        '
        'chkKiJisyaBase
        '
        Me.chkKiJisyaBase.AutoSize = true
        Me.chkKiJisyaBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiJisyaBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiJisyaBase.Name = "chkKiJisyaBase"
        Me.chkKiJisyaBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiJisyaBase.TabIndex = 5
        Me.chkKiJisyaBase.Text = "自社情報"
        Me.chkKiJisyaBase.UseVisualStyleBackColor = true
        '
        'lblKiJisyaBaseCnt
        '
        Me.lblKiJisyaBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiJisyaBaseCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiJisyaBaseCnt.Name = "lblKiJisyaBaseCnt"
        Me.lblKiJisyaBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiJisyaBaseCnt.TabIndex = 8
        Me.lblKiJisyaBaseCnt.Text = "9,999,999"
        Me.lblKiJisyaBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label265
        '
        Me.Label265.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label265.Location = New System.Drawing.Point(19, 54)
        Me.Label265.Name = "Label265"
        Me.Label265.Size = New System.Drawing.Size(92, 16)
        Me.Label265.TabIndex = 7
        Me.Label265.Text = "( 参考件数 = "
        Me.Label265.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label264
        '
        Me.Label264.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label264.Location = New System.Drawing.Point(195, 54)
        Me.Label264.Name = "Label264"
        Me.Label264.Size = New System.Drawing.Size(22, 16)
        Me.Label264.TabIndex = 9
        Me.Label264.Text = " )"
        Me.Label264.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label263
        '
        Me.Label263.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label263.Location = New System.Drawing.Point(19, 33)
        Me.Label263.Name = "Label263"
        Me.Label263.Size = New System.Drawing.Size(198, 16)
        Me.Label263.TabIndex = 6
        Me.Label263.Text = "自社支店, 口座, 担当者"
        Me.Label263.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiSyskanriBase
        '
        Me.pnlKiSyskanriBase.Controls.Add(Me.chkKiSyskanriBase)
        Me.pnlKiSyskanriBase.Controls.Add(Me.lblKiSyskanriBaseCnt)
        Me.pnlKiSyskanriBase.Controls.Add(Me.Label260)
        Me.pnlKiSyskanriBase.Controls.Add(Me.Label259)
        Me.pnlKiSyskanriBase.Controls.Add(Me.Label258)
        Me.pnlKiSyskanriBase.Location = New System.Drawing.Point(15, 20)
        Me.pnlKiSyskanriBase.Name = "pnlKiSyskanriBase"
        Me.pnlKiSyskanriBase.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiSyskanriBase.TabIndex = 2
        '
        'chkKiSyskanriBase
        '
        Me.chkKiSyskanriBase.AutoSize = true
        Me.chkKiSyskanriBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiSyskanriBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiSyskanriBase.Name = "chkKiSyskanriBase"
        Me.chkKiSyskanriBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiSyskanriBase.TabIndex = 0
        Me.chkKiSyskanriBase.Text = "初期設定"
        Me.chkKiSyskanriBase.UseVisualStyleBackColor = true
        '
        'lblKiSyskanriBaseCnt
        '
        Me.lblKiSyskanriBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiSyskanriBaseCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiSyskanriBaseCnt.Name = "lblKiSyskanriBaseCnt"
        Me.lblKiSyskanriBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiSyskanriBaseCnt.TabIndex = 3
        Me.lblKiSyskanriBaseCnt.Text = "9,999,999"
        Me.lblKiSyskanriBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label260
        '
        Me.Label260.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label260.Location = New System.Drawing.Point(19, 54)
        Me.Label260.Name = "Label260"
        Me.Label260.Size = New System.Drawing.Size(92, 16)
        Me.Label260.TabIndex = 2
        Me.Label260.Text = "( 参考件数 = "
        Me.Label260.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label259
        '
        Me.Label259.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label259.Location = New System.Drawing.Point(195, 54)
        Me.Label259.Name = "Label259"
        Me.Label259.Size = New System.Drawing.Size(22, 16)
        Me.Label259.TabIndex = 4
        Me.Label259.Text = " )"
        Me.Label259.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label258
        '
        Me.Label258.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label258.Location = New System.Drawing.Point(19, 33)
        Me.Label258.Name = "Label258"
        Me.Label258.Size = New System.Drawing.Size(198, 16)
        Me.Label258.TabIndex = 1
        Me.Label258.Text = "管理情報設定, タイトル"
        Me.Label258.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSelectPageCnt2
        '
        Me.lblSelectPageCnt2.AutoSize = true
        Me.lblSelectPageCnt2.Location = New System.Drawing.Point(806, 306)
        Me.lblSelectPageCnt2.Name = "lblSelectPageCnt2"
        Me.lblSelectPageCnt2.Size = New System.Drawing.Size(35, 18)
        Me.lblSelectPageCnt2.TabIndex = 28
        Me.lblSelectPageCnt2.Text = "2 / 3"
        '
        'tabPageKizon130
        '
        Me.tabPageKizon130.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageKizon130.Controls.Add(Me.grpKizon2)
        Me.tabPageKizon130.Controls.Add(Me.Label272)
        Me.tabPageKizon130.Location = New System.Drawing.Point(4, 27)
        Me.tabPageKizon130.Name = "tabPageKizon130"
        Me.tabPageKizon130.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageKizon130.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tabPageKizon130.Size = New System.Drawing.Size(847, 327)
        Me.tabPageKizon130.TabIndex = 11
        Me.tabPageKizon130.Text = " 項目選択3"
        '
        'grpKizon2
        '
        Me.grpKizon2.Controls.Add(Me.Label359)
        Me.grpKizon2.Controls.Add(Me.pnlRendo)
        Me.grpKizon2.Controls.Add(Me.pnlKiSq)
        Me.grpKizon2.Controls.Add(Me.pnlSzen)
        Me.grpKizon2.Controls.Add(Me.pnlKiClaim)
        Me.grpKizon2.Controls.Add(Me.pnlKiKy)
        Me.grpKizon2.Controls.Add(Me.pnlKiKys)
        Me.grpKizon2.Controls.Add(Me.pnlKiHy)
        Me.grpKizon2.Controls.Add(Me.pnlKiBk)
        Me.grpKizon2.Controls.Add(Me.grpKizonKagi)
        Me.grpKizon2.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpKizon2.Location = New System.Drawing.Point(6, 5)
        Me.grpKizon2.Name = "grpKizon2"
        Me.grpKizon2.Size = New System.Drawing.Size(835, 298)
        Me.grpKizon2.TabIndex = 0
        Me.grpKizon2.TabStop = false
        Me.grpKizon2.Text = "【基本情報2】"
        '
        'Label359
        '
        Me.Label359.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label359.Location = New System.Drawing.Point(733, 262)
        Me.Label359.Name = "Label359"
        Me.Label359.Size = New System.Drawing.Size(85, 16)
        Me.Label359.TabIndex = 38
        Me.Label359.Text = "(指定IDを設定)"
        Me.Label359.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlRendo
        '
        Me.pnlRendo.Controls.Add(Me.Label83)
        Me.pnlRendo.Controls.Add(Me.txtRendoID)
        Me.pnlRendo.Controls.Add(Me.Label121)
        Me.pnlRendo.Controls.Add(Me.chkKiRendoBase)
        Me.pnlRendo.Controls.Add(Me.lblKiRendoBaseCnt)
        Me.pnlRendo.Controls.Add(Me.Label314)
        Me.pnlRendo.Controls.Add(Me.Label316)
        Me.pnlRendo.Location = New System.Drawing.Point(557, 189)
        Me.pnlRendo.Name = "pnlRendo"
        Me.pnlRendo.Size = New System.Drawing.Size(265, 99)
        Me.pnlRendo.TabIndex = 45
        '
        'Label83
        '
        Me.Label83.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label83.Location = New System.Drawing.Point(32, 73)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(60, 16)
        Me.Label83.TabIndex = 37
        Me.Label83.Text = "連動ID："
        Me.Label83.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtRendoID
        '
        Me.txtRendoID.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.txtRendoID.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtRendoID.Location = New System.Drawing.Point(94, 70)
        Me.txtRendoID.Name = "txtRendoID"
        Me.txtRendoID.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtRendoID.Size = New System.Drawing.Size(71, 24)
        Me.txtRendoID.TabIndex = 36
        Me.txtRendoID.Text = "830013"
        '
        'Label121
        '
        Me.Label121.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label121.Location = New System.Drawing.Point(19, 33)
        Me.Label121.Name = "Label121"
        Me.Label121.Size = New System.Drawing.Size(198, 16)
        Me.Label121.TabIndex = 35
        Me.Label121.Text = "送信設定, 広告補足, 周辺環境"
        Me.Label121.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkKiRendoBase
        '
        Me.chkKiRendoBase.AutoSize = true
        Me.chkKiRendoBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline)
        Me.chkKiRendoBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiRendoBase.Name = "chkKiRendoBase"
        Me.chkKiRendoBase.Size = New System.Drawing.Size(149, 27)
        Me.chkKiRendoBase.TabIndex = 29
        Me.chkKiRendoBase.Text = "ポータル連動情報"
        Me.chkKiRendoBase.UseVisualStyleBackColor = true
        '
        'lblKiRendoBaseCnt
        '
        Me.lblKiRendoBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.lblKiRendoBaseCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiRendoBaseCnt.Name = "lblKiRendoBaseCnt"
        Me.lblKiRendoBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiRendoBaseCnt.TabIndex = 31
        Me.lblKiRendoBaseCnt.Text = "9,999,999"
        Me.lblKiRendoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label314
        '
        Me.Label314.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label314.Location = New System.Drawing.Point(19, 54)
        Me.Label314.Name = "Label314"
        Me.Label314.Size = New System.Drawing.Size(92, 16)
        Me.Label314.TabIndex = 30
        Me.Label314.Text = "( 参考件数 = "
        Me.Label314.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label316
        '
        Me.Label316.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.Label316.Location = New System.Drawing.Point(195, 54)
        Me.Label316.Name = "Label316"
        Me.Label316.Size = New System.Drawing.Size(22, 16)
        Me.Label316.TabIndex = 32
        Me.Label316.Text = " )"
        Me.Label316.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiSq
        '
        Me.pnlKiSq.Controls.Add(Me.Label308)
        Me.pnlKiSq.Controls.Add(Me.lblKiSqOwKojoBaseCnt)
        Me.pnlKiSq.Controls.Add(Me.Label311)
        Me.pnlKiSq.Controls.Add(Me.Label290)
        Me.pnlKiSq.Controls.Add(Me.Label162)
        Me.pnlKiSq.Controls.Add(Me.chkKiSqBase)
        Me.pnlKiSq.Controls.Add(Me.lblKiSqMiBaseCnt)
        Me.pnlKiSq.Controls.Add(Me.Label289)
        Me.pnlKiSq.Controls.Add(Me.Label288)
        Me.pnlKiSq.Controls.Add(Me.lblKiSqAzBaseCnt)
        Me.pnlKiSq.Controls.Add(Me.Label160)
        Me.pnlKiSq.Location = New System.Drawing.Point(286, 188)
        Me.pnlKiSq.Name = "pnlKiSq"
        Me.pnlKiSq.Size = New System.Drawing.Size(265, 100)
        Me.pnlKiSq.TabIndex = 33
        '
        'Label308
        '
        Me.Label308.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label308.Location = New System.Drawing.Point(19, 24)
        Me.Label308.Name = "Label308"
        Me.Label308.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label308.Size = New System.Drawing.Size(92, 16)
        Me.Label308.TabIndex = 42
        Me.Label308.Text = "( 参考件数 = "
        Me.Label308.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label308.Visible = false
        '
        'lblKiSqOwKojoBaseCnt
        '
        Me.lblKiSqOwKojoBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiSqOwKojoBaseCnt.Location = New System.Drawing.Point(106, 24)
        Me.lblKiSqOwKojoBaseCnt.Name = "lblKiSqOwKojoBaseCnt"
        Me.lblKiSqOwKojoBaseCnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblKiSqOwKojoBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiSqOwKojoBaseCnt.TabIndex = 43
        Me.lblKiSqOwKojoBaseCnt.Text = "9,999,999"
        Me.lblKiSqOwKojoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblKiSqOwKojoBaseCnt.Visible = false
        '
        'Label311
        '
        Me.Label311.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label311.Location = New System.Drawing.Point(195, 24)
        Me.Label311.Name = "Label311"
        Me.Label311.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label311.Size = New System.Drawing.Size(22, 16)
        Me.Label311.TabIndex = 44
        Me.Label311.Text = " )"
        Me.Label311.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label311.Visible = false
        '
        'Label290
        '
        Me.Label290.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label290.Location = New System.Drawing.Point(19, 54)
        Me.Label290.Name = "Label290"
        Me.Label290.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label290.Size = New System.Drawing.Size(135, 16)
        Me.Label290.TabIndex = 26
        Me.Label290.Text = "( 未収滞納 参考件数 = "
        Me.Label290.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label162
        '
        Me.Label162.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label162.Location = New System.Drawing.Point(19, 74)
        Me.Label162.Name = "Label162"
        Me.Label162.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label162.Size = New System.Drawing.Size(135, 16)
        Me.Label162.TabIndex = 39
        Me.Label162.Text = "( 預り金 参考件数 = "
        Me.Label162.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkKiSqBase
        '
        Me.chkKiSqBase.AutoSize = true
        Me.chkKiSqBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiSqBase.Location = New System.Drawing.Point(3, 4)
        Me.chkKiSqBase.Name = "chkKiSqBase"
        Me.chkKiSqBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiSqBase.TabIndex = 24
        Me.chkKiSqBase.Text = "請求情報"
        Me.chkKiSqBase.UseVisualStyleBackColor = true
        '
        'lblKiSqMiBaseCnt
        '
        Me.lblKiSqMiBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiSqMiBaseCnt.Location = New System.Drawing.Point(142, 54)
        Me.lblKiSqMiBaseCnt.Name = "lblKiSqMiBaseCnt"
        Me.lblKiSqMiBaseCnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblKiSqMiBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiSqMiBaseCnt.TabIndex = 27
        Me.lblKiSqMiBaseCnt.Text = "9,999,999"
        Me.lblKiSqMiBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label289
        '
        Me.Label289.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label289.Location = New System.Drawing.Point(231, 54)
        Me.Label289.Name = "Label289"
        Me.Label289.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label289.Size = New System.Drawing.Size(22, 16)
        Me.Label289.TabIndex = 28
        Me.Label289.Text = " )"
        Me.Label289.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label288
        '
        Me.Label288.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label288.Location = New System.Drawing.Point(19, 33)
        Me.Label288.Name = "Label288"
        Me.Label288.Size = New System.Drawing.Size(240, 20)
        Me.Label288.TabIndex = 25
        Me.Label288.Text = "未収滞納, 預り金, その他請求, 変動, 控除"
        Me.Label288.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblKiSqAzBaseCnt
        '
        Me.lblKiSqAzBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiSqAzBaseCnt.Location = New System.Drawing.Point(142, 74)
        Me.lblKiSqAzBaseCnt.Name = "lblKiSqAzBaseCnt"
        Me.lblKiSqAzBaseCnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblKiSqAzBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiSqAzBaseCnt.TabIndex = 40
        Me.lblKiSqAzBaseCnt.Text = "9,999,999"
        Me.lblKiSqAzBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label160
        '
        Me.Label160.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label160.Location = New System.Drawing.Point(231, 74)
        Me.Label160.Name = "Label160"
        Me.Label160.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label160.Size = New System.Drawing.Size(22, 16)
        Me.Label160.TabIndex = 41
        Me.Label160.Text = " )"
        Me.Label160.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlSzen
        '
        Me.pnlSzen.Controls.Add(Me.chkKiSzenBase)
        Me.pnlSzen.Controls.Add(Me.lblKiSzenBaseCnt)
        Me.pnlSzen.Controls.Add(Me.Label295)
        Me.pnlSzen.Controls.Add(Me.Label294)
        Me.pnlSzen.Controls.Add(Me.Label174)
        Me.pnlSzen.Location = New System.Drawing.Point(557, 103)
        Me.pnlSzen.Name = "pnlSzen"
        Me.pnlSzen.Size = New System.Drawing.Size(265, 80)
        Me.pnlSzen.TabIndex = 33
        '
        'chkKiSzenBase
        '
        Me.chkKiSzenBase.AutoSize = true
        Me.chkKiSzenBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiSzenBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiSzenBase.Name = "chkKiSzenBase"
        Me.chkKiSzenBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiSzenBase.TabIndex = 33
        Me.chkKiSzenBase.Text = "修繕情報"
        Me.chkKiSzenBase.UseVisualStyleBackColor = true
        '
        'lblKiSzenBaseCnt
        '
        Me.lblKiSzenBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiSzenBaseCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiSzenBaseCnt.Name = "lblKiSzenBaseCnt"
        Me.lblKiSzenBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiSzenBaseCnt.TabIndex = 36
        Me.lblKiSzenBaseCnt.Text = "9,999,999"
        Me.lblKiSzenBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label295
        '
        Me.Label295.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label295.Location = New System.Drawing.Point(19, 54)
        Me.Label295.Name = "Label295"
        Me.Label295.Size = New System.Drawing.Size(92, 16)
        Me.Label295.TabIndex = 35
        Me.Label295.Text = "( 参考件数 = "
        Me.Label295.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label294
        '
        Me.Label294.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label294.Location = New System.Drawing.Point(195, 54)
        Me.Label294.Name = "Label294"
        Me.Label294.Size = New System.Drawing.Size(22, 16)
        Me.Label294.TabIndex = 37
        Me.Label294.Text = " )"
        Me.Label294.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label174
        '
        Me.Label174.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label174.Location = New System.Drawing.Point(19, 33)
        Me.Label174.Name = "Label174"
        Me.Label174.Size = New System.Drawing.Size(198, 16)
        Me.Label174.TabIndex = 34
        Me.Label174.Text = "原状回復, リフォーム"
        Me.Label174.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiClaim
        '
        Me.pnlKiClaim.Controls.Add(Me.chkKiClaimBase)
        Me.pnlKiClaim.Controls.Add(Me.lblKiClaimBaseCnt)
        Me.pnlKiClaim.Controls.Add(Me.Label292)
        Me.pnlKiClaim.Controls.Add(Me.Label180)
        Me.pnlKiClaim.Location = New System.Drawing.Point(557, 20)
        Me.pnlKiClaim.Name = "pnlKiClaim"
        Me.pnlKiClaim.Size = New System.Drawing.Size(265, 60)
        Me.pnlKiClaim.TabIndex = 24
        '
        'chkKiClaimBase
        '
        Me.chkKiClaimBase.AutoSize = true
        Me.chkKiClaimBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiClaimBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiClaimBase.Name = "chkKiClaimBase"
        Me.chkKiClaimBase.Size = New System.Drawing.Size(119, 27)
        Me.chkKiClaimBase.TabIndex = 29
        Me.chkKiClaimBase.Text = "クレーム情報"
        Me.chkKiClaimBase.UseVisualStyleBackColor = true
        '
        'lblKiClaimBaseCnt
        '
        Me.lblKiClaimBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiClaimBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiClaimBaseCnt.Name = "lblKiClaimBaseCnt"
        Me.lblKiClaimBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiClaimBaseCnt.TabIndex = 31
        Me.lblKiClaimBaseCnt.Text = "9,999,999"
        Me.lblKiClaimBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label292
        '
        Me.Label292.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label292.Location = New System.Drawing.Point(19, 33)
        Me.Label292.Name = "Label292"
        Me.Label292.Size = New System.Drawing.Size(92, 16)
        Me.Label292.TabIndex = 30
        Me.Label292.Text = "( 参考件数 = "
        Me.Label292.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label180
        '
        Me.Label180.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label180.Location = New System.Drawing.Point(195, 33)
        Me.Label180.Name = "Label180"
        Me.Label180.Size = New System.Drawing.Size(22, 16)
        Me.Label180.TabIndex = 32
        Me.Label180.Text = " )"
        Me.Label180.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiKy
        '
        Me.pnlKiKy.Controls.Add(Me.chkKiKyBase)
        Me.pnlKiKy.Controls.Add(Me.lblKiKyBaseCnt)
        Me.pnlKiKy.Controls.Add(Me.Label286)
        Me.pnlKiKy.Controls.Add(Me.Label285)
        Me.pnlKiKy.Controls.Add(Me.Label284)
        Me.pnlKiKy.Location = New System.Drawing.Point(286, 104)
        Me.pnlKiKy.Name = "pnlKiKy"
        Me.pnlKiKy.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiKy.TabIndex = 42
        '
        'chkKiKyBase
        '
        Me.chkKiKyBase.AutoSize = true
        Me.chkKiKyBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiKyBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiKyBase.Name = "chkKiKyBase"
        Me.chkKiKyBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiKyBase.TabIndex = 19
        Me.chkKiKyBase.Text = "契約情報"
        Me.chkKiKyBase.UseVisualStyleBackColor = true
        '
        'lblKiKyBaseCnt
        '
        Me.lblKiKyBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiKyBaseCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiKyBaseCnt.Name = "lblKiKyBaseCnt"
        Me.lblKiKyBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiKyBaseCnt.TabIndex = 22
        Me.lblKiKyBaseCnt.Text = "9,999,999"
        Me.lblKiKyBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label286
        '
        Me.Label286.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label286.Location = New System.Drawing.Point(19, 54)
        Me.Label286.Name = "Label286"
        Me.Label286.Size = New System.Drawing.Size(92, 16)
        Me.Label286.TabIndex = 21
        Me.Label286.Text = "( 参考件数 = "
        Me.Label286.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label285
        '
        Me.Label285.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label285.Location = New System.Drawing.Point(195, 54)
        Me.Label285.Name = "Label285"
        Me.Label285.Size = New System.Drawing.Size(22, 16)
        Me.Label285.TabIndex = 23
        Me.Label285.Text = " )"
        Me.Label285.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label284
        '
        Me.Label284.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label284.Location = New System.Drawing.Point(19, 33)
        Me.Label284.Name = "Label284"
        Me.Label284.Size = New System.Drawing.Size(198, 16)
        Me.Label284.TabIndex = 20
        Me.Label284.Text = "契約, 更新, 解約, 入居者"
        Me.Label284.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiKys
        '
        Me.pnlKiKys.Controls.Add(Me.chkKiKysBase)
        Me.pnlKiKys.Controls.Add(Me.lblKiKysBaseCnt)
        Me.pnlKiKys.Controls.Add(Me.Label282)
        Me.pnlKiKys.Controls.Add(Me.Label281)
        Me.pnlKiKys.Controls.Add(Me.Label280)
        Me.pnlKiKys.Location = New System.Drawing.Point(286, 20)
        Me.pnlKiKys.Name = "pnlKiKys"
        Me.pnlKiKys.Size = New System.Drawing.Size(265, 80)
        Me.pnlKiKys.TabIndex = 2
        '
        'chkKiKysBase
        '
        Me.chkKiKysBase.AutoSize = true
        Me.chkKiKysBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiKysBase.Location = New System.Drawing.Point(3, 5)
        Me.chkKiKysBase.Name = "chkKiKysBase"
        Me.chkKiKysBase.Size = New System.Drawing.Size(104, 27)
        Me.chkKiKysBase.TabIndex = 13
        Me.chkKiKysBase.Text = "契約者情報"
        Me.chkKiKysBase.UseVisualStyleBackColor = true
        '
        'lblKiKysBaseCnt
        '
        Me.lblKiKysBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiKysBaseCnt.Location = New System.Drawing.Point(106, 54)
        Me.lblKiKysBaseCnt.Name = "lblKiKysBaseCnt"
        Me.lblKiKysBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiKysBaseCnt.TabIndex = 16
        Me.lblKiKysBaseCnt.Text = "9,999,999"
        Me.lblKiKysBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label282
        '
        Me.Label282.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label282.Location = New System.Drawing.Point(19, 54)
        Me.Label282.Name = "Label282"
        Me.Label282.Size = New System.Drawing.Size(92, 16)
        Me.Label282.TabIndex = 15
        Me.Label282.Text = "( 参考件数 = "
        Me.Label282.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label281
        '
        Me.Label281.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label281.Location = New System.Drawing.Point(195, 54)
        Me.Label281.Name = "Label281"
        Me.Label281.Size = New System.Drawing.Size(22, 16)
        Me.Label281.TabIndex = 17
        Me.Label281.Text = " )"
        Me.Label281.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label280
        '
        Me.Label280.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label280.Location = New System.Drawing.Point(19, 33)
        Me.Label280.Name = "Label280"
        Me.Label280.Size = New System.Drawing.Size(198, 16)
        Me.Label280.TabIndex = 14
        Me.Label280.Text = "契約者・保証人, 口座"
        Me.Label280.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlKiHy
        '
        Me.pnlKiHy.Controls.Add(Me.chkKiHyBase)
        Me.pnlKiHy.Controls.Add(Me.lblKiHyBaseCnt)
        Me.pnlKiHy.Controls.Add(Me.Label275)
        Me.pnlKiHy.Controls.Add(Me.Label270)
        Me.pnlKiHy.Controls.Add(Me.chkKiHySetubi)
        Me.pnlKiHy.Controls.Add(Me.lblKiHySetubiCnt)
        Me.pnlKiHy.Controls.Add(Me.Label278)
        Me.pnlKiHy.Controls.Add(Me.Label277)
        Me.pnlKiHy.Location = New System.Drawing.Point(15, 84)
        Me.pnlKiHy.Name = "pnlKiHy"
        Me.pnlKiHy.Size = New System.Drawing.Size(265, 115)
        Me.pnlKiHy.TabIndex = 2
        '
        'chkKiHyBase
        '
        Me.chkKiHyBase.AutoSize = true
        Me.chkKiHyBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiHyBase.Location = New System.Drawing.Point(3, 4)
        Me.chkKiHyBase.Name = "chkKiHyBase"
        Me.chkKiHyBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiHyBase.TabIndex = 4
        Me.chkKiHyBase.Text = "部屋情報"
        Me.chkKiHyBase.UseVisualStyleBackColor = true
        '
        'lblKiHyBaseCnt
        '
        Me.lblKiHyBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiHyBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiHyBaseCnt.Name = "lblKiHyBaseCnt"
        Me.lblKiHyBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiHyBaseCnt.TabIndex = 6
        Me.lblKiHyBaseCnt.Text = "9,999,999"
        Me.lblKiHyBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label275
        '
        Me.Label275.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label275.Location = New System.Drawing.Point(19, 33)
        Me.Label275.Name = "Label275"
        Me.Label275.Size = New System.Drawing.Size(92, 16)
        Me.Label275.TabIndex = 5
        Me.Label275.Text = "( 参考件数 = "
        Me.Label275.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label270
        '
        Me.Label270.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label270.Location = New System.Drawing.Point(195, 33)
        Me.Label270.Name = "Label270"
        Me.Label270.Size = New System.Drawing.Size(22, 16)
        Me.Label270.TabIndex = 7
        Me.Label270.Text = " )"
        Me.Label270.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkKiHySetubi
        '
        Me.chkKiHySetubi.AutoSize = true
        Me.chkKiHySetubi.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiHySetubi.Location = New System.Drawing.Point(18, 56)
        Me.chkKiHySetubi.Name = "chkKiHySetubi"
        Me.chkKiHySetubi.Size = New System.Drawing.Size(119, 27)
        Me.chkKiHySetubi.TabIndex = 8
        Me.chkKiHySetubi.Text = "部屋設備情報"
        Me.chkKiHySetubi.UseVisualStyleBackColor = true
        '
        'lblKiHySetubiCnt
        '
        Me.lblKiHySetubiCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiHySetubiCnt.Location = New System.Drawing.Point(118, 86)
        Me.lblKiHySetubiCnt.Name = "lblKiHySetubiCnt"
        Me.lblKiHySetubiCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiHySetubiCnt.TabIndex = 10
        Me.lblKiHySetubiCnt.Text = "9,999,999"
        Me.lblKiHySetubiCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label278
        '
        Me.Label278.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label278.Location = New System.Drawing.Point(34, 86)
        Me.Label278.Name = "Label278"
        Me.Label278.Size = New System.Drawing.Size(92, 16)
        Me.Label278.TabIndex = 9
        Me.Label278.Text = "( 参考件数 = "
        Me.Label278.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label277
        '
        Me.Label277.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label277.Location = New System.Drawing.Point(207, 86)
        Me.Label277.Name = "Label277"
        Me.Label277.Size = New System.Drawing.Size(22, 16)
        Me.Label277.TabIndex = 11
        Me.Label277.Text = " )"
        Me.Label277.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlKiBk
        '
        Me.pnlKiBk.Controls.Add(Me.chkKiBkBase)
        Me.pnlKiBk.Controls.Add(Me.lblKiBkBaseCnt)
        Me.pnlKiBk.Controls.Add(Me.Label273)
        Me.pnlKiBk.Controls.Add(Me.Label271)
        Me.pnlKiBk.Location = New System.Drawing.Point(15, 20)
        Me.pnlKiBk.Name = "pnlKiBk"
        Me.pnlKiBk.Size = New System.Drawing.Size(265, 60)
        Me.pnlKiBk.TabIndex = 1
        '
        'chkKiBkBase
        '
        Me.chkKiBkBase.AutoSize = true
        Me.chkKiBkBase.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkKiBkBase.Location = New System.Drawing.Point(3, 3)
        Me.chkKiBkBase.Name = "chkKiBkBase"
        Me.chkKiBkBase.Size = New System.Drawing.Size(89, 27)
        Me.chkKiBkBase.TabIndex = 0
        Me.chkKiBkBase.Text = "物件情報"
        Me.chkKiBkBase.UseVisualStyleBackColor = true
        '
        'lblKiBkBaseCnt
        '
        Me.lblKiBkBaseCnt.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKiBkBaseCnt.Location = New System.Drawing.Point(106, 33)
        Me.lblKiBkBaseCnt.Name = "lblKiBkBaseCnt"
        Me.lblKiBkBaseCnt.Size = New System.Drawing.Size(83, 16)
        Me.lblKiBkBaseCnt.TabIndex = 2
        Me.lblKiBkBaseCnt.Text = "9,999,999"
        Me.lblKiBkBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label273
        '
        Me.Label273.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label273.Location = New System.Drawing.Point(19, 33)
        Me.Label273.Name = "Label273"
        Me.Label273.Size = New System.Drawing.Size(92, 16)
        Me.Label273.TabIndex = 1
        Me.Label273.Text = "( 参考件数 = "
        Me.Label273.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label271
        '
        Me.Label271.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label271.Location = New System.Drawing.Point(195, 33)
        Me.Label271.Name = "Label271"
        Me.Label271.Size = New System.Drawing.Size(22, 16)
        Me.Label271.TabIndex = 3
        Me.Label271.Text = " )"
        Me.Label271.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpKizonKagi
        '
        Me.grpKizonKagi.Controls.Add(Me.optKiKyKagi)
        Me.grpKizonKagi.Controls.Add(Me.optKiHyKagi)
        Me.grpKizonKagi.Controls.Add(Me.Label86)
        Me.grpKizonKagi.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpKizonKagi.Location = New System.Drawing.Point(18, 205)
        Me.grpKizonKagi.Name = "grpKizonKagi"
        Me.grpKizonKagi.Size = New System.Drawing.Size(262, 87)
        Me.grpKizonKagi.TabIndex = 12
        Me.grpKizonKagi.TabStop = false
        Me.grpKizonKagi.Text = "鍵情報選択"
        '
        'optKiKyKagi
        '
        Me.optKiKyKagi.AutoSize = true
        Me.optKiKyKagi.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.optKiKyKagi.Location = New System.Drawing.Point(150, 54)
        Me.optKiKyKagi.Name = "optKiKyKagi"
        Me.optKiKyKagi.Size = New System.Drawing.Size(103, 27)
        Me.optKiKyKagi.TabIndex = 2
        Me.optKiKyKagi.TabStop = true
        Me.optKiKyKagi.Text = "契約鍵情報"
        Me.optKiKyKagi.UseVisualStyleBackColor = true
        '
        'optKiHyKagi
        '
        Me.optKiHyKagi.AutoSize = true
        Me.optKiHyKagi.Checked = true
        Me.optKiHyKagi.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.optKiHyKagi.Location = New System.Drawing.Point(20, 54)
        Me.optKiHyKagi.Name = "optKiHyKagi"
        Me.optKiHyKagi.Size = New System.Drawing.Size(103, 27)
        Me.optKiHyKagi.TabIndex = 1
        Me.optKiHyKagi.TabStop = true
        Me.optKiHyKagi.Text = "部屋鍵情報"
        Me.optKiHyKagi.UseVisualStyleBackColor = true
        '
        'Label86
        '
        Me.Label86.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label86.Location = New System.Drawing.Point(12, 20)
        Me.Label86.Name = "Label86"
        Me.Label86.Size = New System.Drawing.Size(244, 35)
        Me.Label86.TabIndex = 0
        Me.Label86.Text = "V7では部屋鍵と契約鍵が存在します。どちらを移行対象とするか選択して下さい。"
        '
        'Label272
        '
        Me.Label272.AutoSize = true
        Me.Label272.Location = New System.Drawing.Point(806, 306)
        Me.Label272.Name = "Label272"
        Me.Label272.Size = New System.Drawing.Size(35, 18)
        Me.Label272.TabIndex = 38
        Me.Label272.Text = "3 / 3"
        '
        'tabPageBase110
        '
        Me.tabPageBase110.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase110.Controls.Add(Me.lblRelItemInfo)
        Me.tabPageBase110.Controls.Add(Me.grpMst)
        Me.tabPageBase110.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase110.Name = "tabPageBase110"
        Me.tabPageBase110.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase110.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase110.TabIndex = 0
        Me.tabPageBase110.Text = "各マスタ情報"
        '
        'lblRelItemInfo
        '
        Me.lblRelItemInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblRelItemInfo.Location = New System.Drawing.Point(439, 20)
        Me.lblRelItemInfo.Name = "lblRelItemInfo"
        Me.lblRelItemInfo.Size = New System.Drawing.Size(377, 235)
        Me.lblRelItemInfo.TabIndex = 1
        Me.lblRelItemInfo.Text = ""&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  以下は賃貸革命10のマスター情報と紐付けを行う項目です。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  紐付設定画面にて、ユーザーデータと革命10のマスターデータの"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  紐付けを行って下さい。"& _ 
    ""&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"    ・ 物件分類マスタ  ・ 取引態様マスタ  ・ 部屋分類マスタ"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"    ・ 口座種別マスタ  ・ 構造マスタ        ・ 入金区分マスタ"& _ 
    ""&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"    ・ 設備マスタ        ・ 入金項目マスタ"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)
        '
        'grpMst
        '
        Me.grpMst.Controls.Add(Me.chkMstGazotitle)
        Me.grpMst.Controls.Add(Me.chkMstBikolst)
        Me.grpMst.Controls.Add(Me.chkMstBikotitle)
        Me.grpMst.Controls.Add(Me.chkMstHendoitiran)
        Me.grpMst.Controls.Add(Me.chkMstHendo)
        Me.grpMst.Controls.Add(Me.chkMstKasyoClaimrui)
        Me.grpMst.Controls.Add(Me.chkMstBus)
        Me.grpMst.Controls.Add(Me.chkMstSchool)
        Me.grpMst.Controls.Add(Me.chkMstTokuyaku)
        Me.grpMst.Controls.Add(Me.chkMstHokenrui)
        Me.grpMst.Controls.Add(Me.chkMstBusKotu)
        Me.grpMst.Controls.Add(Me.chkMstKagititle)
        Me.grpMst.Controls.Add(Me.chkMstArea)
        Me.grpMst.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpMst.Location = New System.Drawing.Point(10, 5)
        Me.grpMst.Name = "grpMst"
        Me.grpMst.Size = New System.Drawing.Size(397, 265)
        Me.grpMst.TabIndex = 0
        Me.grpMst.TabStop = false
        Me.grpMst.Text = " 各マスタ情報 "
        '
        'chkMstGazotitle
        '
        Me.chkMstGazotitle.AutoSize = true
        Me.chkMstGazotitle.Location = New System.Drawing.Point(200, 192)
        Me.chkMstGazotitle.Name = "chkMstGazotitle"
        Me.chkMstGazotitle.Size = New System.Drawing.Size(135, 22)
        Me.chkMstGazotitle.TabIndex = 13
        Me.chkMstGazotitle.Text = "画像タイトルマスタ"
        Me.chkMstGazotitle.UseVisualStyleBackColor = true
        '
        'chkMstBikolst
        '
        Me.chkMstBikolst.AutoSize = true
        Me.chkMstBikolst.Location = New System.Drawing.Point(216, 165)
        Me.chkMstBikolst.Name = "chkMstBikolst"
        Me.chkMstBikolst.Size = New System.Drawing.Size(171, 22)
        Me.chkMstBikolst.TabIndex = 12
        Me.chkMstBikolst.Text = "備考入力補助リストマスタ"
        Me.chkMstBikolst.UseVisualStyleBackColor = true
        '
        'chkMstBikotitle
        '
        Me.chkMstBikotitle.AutoSize = true
        Me.chkMstBikotitle.Location = New System.Drawing.Point(200, 136)
        Me.chkMstBikotitle.Name = "chkMstBikotitle"
        Me.chkMstBikotitle.Size = New System.Drawing.Size(135, 22)
        Me.chkMstBikotitle.TabIndex = 11
        Me.chkMstBikotitle.Text = "備考タイトルマスタ"
        Me.chkMstBikotitle.UseVisualStyleBackColor = true
        '
        'chkMstHendoitiran
        '
        Me.chkMstHendoitiran.AutoSize = true
        Me.chkMstHendoitiran.Location = New System.Drawing.Point(215, 109)
        Me.chkMstHendoitiran.Name = "chkMstHendoitiran"
        Me.chkMstHendoitiran.Size = New System.Drawing.Size(123, 22)
        Me.chkMstHendoitiran.TabIndex = 10
        Me.chkMstHendoitiran.Text = "変動費料金単価表"
        Me.chkMstHendoitiran.UseVisualStyleBackColor = true
        '
        'chkMstHendo
        '
        Me.chkMstHendo.AutoSize = true
        Me.chkMstHendo.Location = New System.Drawing.Point(201, 81)
        Me.chkMstHendo.Name = "chkMstHendo"
        Me.chkMstHendo.Size = New System.Drawing.Size(111, 22)
        Me.chkMstHendo.TabIndex = 9
        Me.chkMstHendo.Text = "変動費設定内容"
        Me.chkMstHendo.UseVisualStyleBackColor = true
        '
        'chkMstKasyoClaimrui
        '
        Me.chkMstKasyoClaimrui.AutoSize = true
        Me.chkMstKasyoClaimrui.Location = New System.Drawing.Point(200, 53)
        Me.chkMstKasyoClaimrui.Name = "chkMstKasyoClaimrui"
        Me.chkMstKasyoClaimrui.Size = New System.Drawing.Size(147, 22)
        Me.chkMstKasyoClaimrui.TabIndex = 8
        Me.chkMstKasyoClaimrui.Text = "クレーム分類設定内容"
        Me.chkMstKasyoClaimrui.UseVisualStyleBackColor = true
        '
        'chkMstBus
        '
        Me.chkMstBus.AutoSize = true
        Me.chkMstBus.Location = New System.Drawing.Point(25, 25)
        Me.chkMstBus.Name = "chkMstBus"
        Me.chkMstBus.Size = New System.Drawing.Size(111, 22)
        Me.chkMstBus.TabIndex = 0
        Me.chkMstBus.Text = "バス交通マスタ"
        Me.chkMstBus.UseVisualStyleBackColor = true
        '
        'chkMstSchool
        '
        Me.chkMstSchool.AutoSize = true
        Me.chkMstSchool.Location = New System.Drawing.Point(25, 81)
        Me.chkMstSchool.Name = "chkMstSchool"
        Me.chkMstSchool.Size = New System.Drawing.Size(99, 22)
        Me.chkMstSchool.TabIndex = 2
        Me.chkMstSchool.Text = "学校区マスタ"
        Me.chkMstSchool.UseVisualStyleBackColor = true
        '
        'chkMstTokuyaku
        '
        Me.chkMstTokuyaku.AutoSize = true
        Me.chkMstTokuyaku.Location = New System.Drawing.Point(200, 25)
        Me.chkMstTokuyaku.Name = "chkMstTokuyaku"
        Me.chkMstTokuyaku.Size = New System.Drawing.Size(87, 22)
        Me.chkMstTokuyaku.TabIndex = 7
        Me.chkMstTokuyaku.Text = "特約マスタ"
        Me.chkMstTokuyaku.UseVisualStyleBackColor = true
        '
        'chkMstHokenrui
        '
        Me.chkMstHokenrui.AutoSize = true
        Me.chkMstHokenrui.Location = New System.Drawing.Point(25, 137)
        Me.chkMstHokenrui.Name = "chkMstHokenrui"
        Me.chkMstHokenrui.Size = New System.Drawing.Size(111, 22)
        Me.chkMstHokenrui.TabIndex = 4
        Me.chkMstHokenrui.Text = "保険種類マスタ"
        Me.chkMstHokenrui.UseVisualStyleBackColor = true
        '
        'chkMstBusKotu
        '
        Me.chkMstBusKotu.AutoSize = true
        Me.chkMstBusKotu.Location = New System.Drawing.Point(45, 53)
        Me.chkMstBusKotu.Name = "chkMstBusKotu"
        Me.chkMstBusKotu.Size = New System.Drawing.Size(99, 22)
        Me.chkMstBusKotu.TabIndex = 1
        Me.chkMstBusKotu.Text = "バス停マスタ"
        Me.chkMstBusKotu.UseVisualStyleBackColor = true
        '
        'chkMstKagititle
        '
        Me.chkMstKagititle.AutoSize = true
        Me.chkMstKagititle.Location = New System.Drawing.Point(25, 168)
        Me.chkMstKagititle.Name = "chkMstKagititle"
        Me.chkMstKagititle.Size = New System.Drawing.Size(123, 22)
        Me.chkMstKagititle.TabIndex = 6
        Me.chkMstKagititle.Text = "鍵タイトルマスタ"
        Me.chkMstKagititle.UseVisualStyleBackColor = true
        '
        'chkMstArea
        '
        Me.chkMstArea.AutoSize = true
        Me.chkMstArea.Location = New System.Drawing.Point(25, 109)
        Me.chkMstArea.Name = "chkMstArea"
        Me.chkMstArea.Size = New System.Drawing.Size(99, 22)
        Me.chkMstArea.TabIndex = 3
        Me.chkMstArea.Text = "エリアマスタ"
        Me.chkMstArea.UseVisualStyleBackColor = true
        '
        'tabPageBase120
        '
        Me.tabPageBase120.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase120.Controls.Add(Me.grpGy)
        Me.tabPageBase120.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase120.Name = "tabPageBase120"
        Me.tabPageBase120.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase120.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase120.TabIndex = 1
        Me.tabPageBase120.Text = "業者情報"
        '
        'grpGy
        '
        Me.grpGy.Controls.Add(Me.chkGySyuzenMemo)
        Me.grpGy.Controls.Add(Me.chkGyHokenMemo)
        Me.grpGy.Controls.Add(Me.chkGySyuzenKoza)
        Me.grpGy.Controls.Add(Me.chkGySekoBase)
        Me.grpGy.Controls.Add(Me.chkGySisetuBase)
        Me.grpGy.Controls.Add(Me.chkGySyuzenBase)
        Me.grpGy.Controls.Add(Me.chkGyYatinhosyoMemo)
        Me.grpGy.Controls.Add(Me.chkGyYatinhosyoBase)
        Me.grpGy.Controls.Add(Me.chkGyCyukaiKoza)
        Me.grpGy.Controls.Add(Me.chkGyCyukaiMemo)
        Me.grpGy.Controls.Add(Me.chkGyHokenBase)
        Me.grpGy.Controls.Add(Me.chkGyLifelineBase)
        Me.grpGy.Controls.Add(Me.chkGyHokenKoza)
        Me.grpGy.Controls.Add(Me.chkGyCyukaiBase)
        Me.grpGy.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpGy.Location = New System.Drawing.Point(10, 5)
        Me.grpGy.Name = "grpGy"
        Me.grpGy.Size = New System.Drawing.Size(588, 265)
        Me.grpGy.TabIndex = 0
        Me.grpGy.TabStop = false
        Me.grpGy.Text = " 業者情報 "
        '
        'chkGySyuzenMemo
        '
        Me.chkGySyuzenMemo.AutoSize = true
        Me.chkGySyuzenMemo.Location = New System.Drawing.Point(245, 80)
        Me.chkGySyuzenMemo.Name = "chkGySyuzenMemo"
        Me.chkGySyuzenMemo.Size = New System.Drawing.Size(123, 22)
        Me.chkGySyuzenMemo.TabIndex = 9
        Me.chkGySyuzenMemo.Text = "修繕業者メモ情報"
        Me.chkGySyuzenMemo.UseVisualStyleBackColor = true
        '
        'chkGyHokenMemo
        '
        Me.chkGyHokenMemo.AutoSize = true
        Me.chkGyHokenMemo.Location = New System.Drawing.Point(45, 164)
        Me.chkGyHokenMemo.Name = "chkGyHokenMemo"
        Me.chkGyHokenMemo.Size = New System.Drawing.Size(123, 22)
        Me.chkGyHokenMemo.TabIndex = 5
        Me.chkGyHokenMemo.Text = "保険業者メモ情報"
        Me.chkGyHokenMemo.UseVisualStyleBackColor = true
        '
        'chkGySyuzenKoza
        '
        Me.chkGySyuzenKoza.AutoSize = true
        Me.chkGySyuzenKoza.Location = New System.Drawing.Point(245, 52)
        Me.chkGySyuzenKoza.Name = "chkGySyuzenKoza"
        Me.chkGySyuzenKoza.Size = New System.Drawing.Size(123, 22)
        Me.chkGySyuzenKoza.TabIndex = 8
        Me.chkGySyuzenKoza.Text = "修繕業者口座情報"
        Me.chkGySyuzenKoza.UseVisualStyleBackColor = true
        '
        'chkGySekoBase
        '
        Me.chkGySekoBase.AutoSize = true
        Me.chkGySekoBase.Location = New System.Drawing.Point(425, 52)
        Me.chkGySekoBase.Name = "chkGySekoBase"
        Me.chkGySekoBase.Size = New System.Drawing.Size(99, 22)
        Me.chkGySekoBase.TabIndex = 13
        Me.chkGySekoBase.Text = "施工業者情報"
        Me.chkGySekoBase.UseVisualStyleBackColor = true
        '
        'chkGySisetuBase
        '
        Me.chkGySisetuBase.AutoSize = true
        Me.chkGySisetuBase.Location = New System.Drawing.Point(425, 25)
        Me.chkGySisetuBase.Name = "chkGySisetuBase"
        Me.chkGySisetuBase.Size = New System.Drawing.Size(123, 22)
        Me.chkGySisetuBase.TabIndex = 12
        Me.chkGySisetuBase.Text = "施設保守業者情報"
        Me.chkGySisetuBase.UseVisualStyleBackColor = true
        '
        'chkGySyuzenBase
        '
        Me.chkGySyuzenBase.AutoSize = true
        Me.chkGySyuzenBase.Location = New System.Drawing.Point(225, 25)
        Me.chkGySyuzenBase.Name = "chkGySyuzenBase"
        Me.chkGySyuzenBase.Size = New System.Drawing.Size(123, 22)
        Me.chkGySyuzenBase.TabIndex = 7
        Me.chkGySyuzenBase.Text = "修繕業者基本情報"
        Me.chkGySyuzenBase.UseVisualStyleBackColor = true
        '
        'chkGyYatinhosyoMemo
        '
        Me.chkGyYatinhosyoMemo.AutoSize = true
        Me.chkGyYatinhosyoMemo.Location = New System.Drawing.Point(245, 136)
        Me.chkGyYatinhosyoMemo.Name = "chkGyYatinhosyoMemo"
        Me.chkGyYatinhosyoMemo.Size = New System.Drawing.Size(147, 22)
        Me.chkGyYatinhosyoMemo.TabIndex = 11
        Me.chkGyYatinhosyoMemo.Text = "家賃保証業者メモ情報"
        Me.chkGyYatinhosyoMemo.UseVisualStyleBackColor = true
        '
        'chkGyYatinhosyoBase
        '
        Me.chkGyYatinhosyoBase.AutoSize = true
        Me.chkGyYatinhosyoBase.Location = New System.Drawing.Point(225, 108)
        Me.chkGyYatinhosyoBase.Name = "chkGyYatinhosyoBase"
        Me.chkGyYatinhosyoBase.Size = New System.Drawing.Size(147, 22)
        Me.chkGyYatinhosyoBase.TabIndex = 10
        Me.chkGyYatinhosyoBase.Text = "家賃保証業者基本情報"
        Me.chkGyYatinhosyoBase.UseVisualStyleBackColor = true
        '
        'chkGyCyukaiKoza
        '
        Me.chkGyCyukaiKoza.AutoSize = true
        Me.chkGyCyukaiKoza.Location = New System.Drawing.Point(45, 52)
        Me.chkGyCyukaiKoza.Name = "chkGyCyukaiKoza"
        Me.chkGyCyukaiKoza.Size = New System.Drawing.Size(123, 22)
        Me.chkGyCyukaiKoza.TabIndex = 1
        Me.chkGyCyukaiKoza.Text = "仲介業者口座情報"
        Me.chkGyCyukaiKoza.UseVisualStyleBackColor = true
        '
        'chkGyCyukaiMemo
        '
        Me.chkGyCyukaiMemo.AutoSize = true
        Me.chkGyCyukaiMemo.Location = New System.Drawing.Point(45, 80)
        Me.chkGyCyukaiMemo.Name = "chkGyCyukaiMemo"
        Me.chkGyCyukaiMemo.Size = New System.Drawing.Size(123, 22)
        Me.chkGyCyukaiMemo.TabIndex = 2
        Me.chkGyCyukaiMemo.Text = "仲介業者メモ情報"
        Me.chkGyCyukaiMemo.UseVisualStyleBackColor = true
        '
        'chkGyHokenBase
        '
        Me.chkGyHokenBase.AutoSize = true
        Me.chkGyHokenBase.Location = New System.Drawing.Point(25, 108)
        Me.chkGyHokenBase.Name = "chkGyHokenBase"
        Me.chkGyHokenBase.Size = New System.Drawing.Size(123, 22)
        Me.chkGyHokenBase.TabIndex = 3
        Me.chkGyHokenBase.Text = "保険業者基本情報"
        Me.chkGyHokenBase.UseVisualStyleBackColor = true
        '
        'chkGyLifelineBase
        '
        Me.chkGyLifelineBase.AutoSize = true
        Me.chkGyLifelineBase.Location = New System.Drawing.Point(25, 192)
        Me.chkGyLifelineBase.Name = "chkGyLifelineBase"
        Me.chkGyLifelineBase.Size = New System.Drawing.Size(147, 22)
        Me.chkGyLifelineBase.TabIndex = 6
        Me.chkGyLifelineBase.Text = "ライフライン業者情報"
        Me.chkGyLifelineBase.UseVisualStyleBackColor = true
        '
        'chkGyHokenKoza
        '
        Me.chkGyHokenKoza.AutoSize = true
        Me.chkGyHokenKoza.Location = New System.Drawing.Point(45, 136)
        Me.chkGyHokenKoza.Name = "chkGyHokenKoza"
        Me.chkGyHokenKoza.Size = New System.Drawing.Size(123, 22)
        Me.chkGyHokenKoza.TabIndex = 4
        Me.chkGyHokenKoza.Text = "保険業者口座情報"
        Me.chkGyHokenKoza.UseVisualStyleBackColor = true
        '
        'chkGyCyukaiBase
        '
        Me.chkGyCyukaiBase.AutoSize = true
        Me.chkGyCyukaiBase.Location = New System.Drawing.Point(25, 24)
        Me.chkGyCyukaiBase.Name = "chkGyCyukaiBase"
        Me.chkGyCyukaiBase.Size = New System.Drawing.Size(123, 22)
        Me.chkGyCyukaiBase.TabIndex = 0
        Me.chkGyCyukaiBase.Text = "仲介業者基本情報"
        Me.chkGyCyukaiBase.UseVisualStyleBackColor = true
        '
        'tabPageBase130
        '
        Me.tabPageBase130.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase130.Controls.Add(Me.grpKys)
        Me.tabPageBase130.Controls.Add(Me.grpOw)
        Me.tabPageBase130.Controls.Add(Me.grpJisya)
        Me.tabPageBase130.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase130.Name = "tabPageBase130"
        Me.tabPageBase130.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase130.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase130.TabIndex = 2
        Me.tabPageBase130.Text = "各基本情報"
        '
        'grpKys
        '
        Me.grpKys.Controls.Add(Me.chkKysHosyonin)
        Me.grpKys.Controls.Add(Me.chkKysSyogoKana)
        Me.grpKys.Controls.Add(Me.chkKysKoza)
        Me.grpKys.Controls.Add(Me.chkKysMemo)
        Me.grpKys.Controls.Add(Me.chkKysBase)
        Me.grpKys.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpKys.Location = New System.Drawing.Point(638, 6)
        Me.grpKys.Name = "grpKys"
        Me.grpKys.Size = New System.Drawing.Size(200, 265)
        Me.grpKys.TabIndex = 2
        Me.grpKys.TabStop = false
        Me.grpKys.Text = "契約者情報"
        '
        'chkKysHosyonin
        '
        Me.chkKysHosyonin.AutoSize = true
        Me.chkKysHosyonin.Location = New System.Drawing.Point(45, 108)
        Me.chkKysHosyonin.Name = "chkKysHosyonin"
        Me.chkKysHosyonin.Size = New System.Drawing.Size(123, 22)
        Me.chkKysHosyonin.TabIndex = 3
        Me.chkKysHosyonin.Text = "契約者保証人情報"
        Me.chkKysHosyonin.UseVisualStyleBackColor = true
        '
        'chkKysSyogoKana
        '
        Me.chkKysSyogoKana.AutoSize = true
        Me.chkKysSyogoKana.Location = New System.Drawing.Point(45, 80)
        Me.chkKysSyogoKana.Name = "chkKysSyogoKana"
        Me.chkKysSyogoKana.Size = New System.Drawing.Size(147, 22)
        Me.chkKysSyogoKana.TabIndex = 2
        Me.chkKysSyogoKana.Text = "契約者照合用カナ情報"
        Me.chkKysSyogoKana.UseVisualStyleBackColor = true
        '
        'chkKysKoza
        '
        Me.chkKysKoza.AutoSize = true
        Me.chkKysKoza.Location = New System.Drawing.Point(45, 53)
        Me.chkKysKoza.Name = "chkKysKoza"
        Me.chkKysKoza.Size = New System.Drawing.Size(111, 22)
        Me.chkKysKoza.TabIndex = 1
        Me.chkKysKoza.Text = "契約者口座情報"
        Me.chkKysKoza.UseVisualStyleBackColor = true
        '
        'chkKysMemo
        '
        Me.chkKysMemo.AutoSize = true
        Me.chkKysMemo.Location = New System.Drawing.Point(45, 136)
        Me.chkKysMemo.Name = "chkKysMemo"
        Me.chkKysMemo.Size = New System.Drawing.Size(111, 22)
        Me.chkKysMemo.TabIndex = 4
        Me.chkKysMemo.Text = "契約者メモ情報"
        Me.chkKysMemo.UseVisualStyleBackColor = true
        '
        'chkKysBase
        '
        Me.chkKysBase.AutoSize = true
        Me.chkKysBase.Location = New System.Drawing.Point(25, 25)
        Me.chkKysBase.Name = "chkKysBase"
        Me.chkKysBase.Size = New System.Drawing.Size(111, 22)
        Me.chkKysBase.TabIndex = 0
        Me.chkKysBase.Text = "契約者基本情報"
        Me.chkKysBase.UseVisualStyleBackColor = true
        '
        'grpOw
        '
        Me.grpOw.Controls.Add(Me.chkOwEvent)
        Me.grpOw.Controls.Add(Me.chkOwKoza)
        Me.grpOw.Controls.Add(Me.chkOwMemo)
        Me.grpOw.Controls.Add(Me.chkOwBase)
        Me.grpOw.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpOw.Location = New System.Drawing.Point(432, 6)
        Me.grpOw.Name = "grpOw"
        Me.grpOw.Size = New System.Drawing.Size(200, 265)
        Me.grpOw.TabIndex = 1
        Me.grpOw.TabStop = false
        Me.grpOw.Text = "家主情報"
        '
        'chkOwEvent
        '
        Me.chkOwEvent.AutoSize = true
        Me.chkOwEvent.Location = New System.Drawing.Point(45, 80)
        Me.chkOwEvent.Name = "chkOwEvent"
        Me.chkOwEvent.Size = New System.Drawing.Size(123, 22)
        Me.chkOwEvent.TabIndex = 2
        Me.chkOwEvent.Text = "家主イベント情報"
        Me.chkOwEvent.UseVisualStyleBackColor = true
        '
        'chkOwKoza
        '
        Me.chkOwKoza.AutoSize = true
        Me.chkOwKoza.Location = New System.Drawing.Point(45, 53)
        Me.chkOwKoza.Name = "chkOwKoza"
        Me.chkOwKoza.Size = New System.Drawing.Size(99, 22)
        Me.chkOwKoza.TabIndex = 1
        Me.chkOwKoza.Text = "家主口座情報"
        Me.chkOwKoza.UseVisualStyleBackColor = true
        '
        'chkOwMemo
        '
        Me.chkOwMemo.AutoSize = true
        Me.chkOwMemo.Location = New System.Drawing.Point(45, 108)
        Me.chkOwMemo.Name = "chkOwMemo"
        Me.chkOwMemo.Size = New System.Drawing.Size(99, 22)
        Me.chkOwMemo.TabIndex = 3
        Me.chkOwMemo.Text = "家主メモ情報"
        Me.chkOwMemo.UseVisualStyleBackColor = true
        '
        'chkOwBase
        '
        Me.chkOwBase.AutoSize = true
        Me.chkOwBase.Location = New System.Drawing.Point(25, 25)
        Me.chkOwBase.Name = "chkOwBase"
        Me.chkOwBase.Size = New System.Drawing.Size(99, 22)
        Me.chkOwBase.TabIndex = 0
        Me.chkOwBase.Text = "家主基本情報"
        Me.chkOwBase.UseVisualStyleBackColor = true
        '
        'grpJisya
        '
        Me.grpJisya.Controls.Add(Me.chkFBANSERSetuzoku)
        Me.grpJisya.Controls.Add(Me.chkMstANSERArea)
        Me.grpJisya.Controls.Add(Me.chkMstANSERAccpoint)
        Me.grpJisya.Controls.Add(Me.chkJisyaTanto)
        Me.grpJisya.Controls.Add(Me.chkMstYatinKoza)
        Me.grpJisya.Controls.Add(Me.chkJisyaKoza)
        Me.grpJisya.Controls.Add(Me.chkJisyaBase)
        Me.grpJisya.Controls.Add(Me.chkFBFuriirai)
        Me.grpJisya.Controls.Add(Me.chkFBNsSyutoku)
        Me.grpJisya.Controls.Add(Me.chkJisyaMemo)
        Me.grpJisya.Controls.Add(Me.chkFBKozafurikae)
        Me.grpJisya.Controls.Add(Me.chkFBFuritesuryo)
        Me.grpJisya.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpJisya.Location = New System.Drawing.Point(10, 5)
        Me.grpJisya.Name = "grpJisya"
        Me.grpJisya.Size = New System.Drawing.Size(398, 266)
        Me.grpJisya.TabIndex = 0
        Me.grpJisya.TabStop = false
        Me.grpJisya.Text = " 自社情報 "
        '
        'chkFBANSERSetuzoku
        '
        Me.chkFBANSERSetuzoku.AutoSize = true
        Me.chkFBANSERSetuzoku.Location = New System.Drawing.Point(176, 81)
        Me.chkFBANSERSetuzoku.Name = "chkFBANSERSetuzoku"
        Me.chkFBANSERSetuzoku.Size = New System.Drawing.Size(115, 22)
        Me.chkFBANSERSetuzoku.TabIndex = 17
        Me.chkFBANSERSetuzoku.Text = "ANSER接続情報"
        Me.chkFBANSERSetuzoku.UseVisualStyleBackColor = true
        '
        'chkMstANSERArea
        '
        Me.chkMstANSERArea.AutoSize = true
        Me.chkMstANSERArea.Location = New System.Drawing.Point(176, 53)
        Me.chkMstANSERArea.Name = "chkMstANSERArea"
        Me.chkMstANSERArea.Size = New System.Drawing.Size(127, 22)
        Me.chkMstANSERArea.TabIndex = 16
        Me.chkMstANSERArea.Text = "ANSERエリア情報"
        Me.chkMstANSERArea.UseVisualStyleBackColor = true
        '
        'chkMstANSERAccpoint
        '
        Me.chkMstANSERAccpoint.AutoSize = true
        Me.chkMstANSERAccpoint.Location = New System.Drawing.Point(176, 25)
        Me.chkMstANSERAccpoint.Name = "chkMstANSERAccpoint"
        Me.chkMstANSERAccpoint.Size = New System.Drawing.Size(187, 22)
        Me.chkMstANSERAccpoint.TabIndex = 15
        Me.chkMstANSERAccpoint.Text = "ANSERアクセスポイント情報"
        Me.chkMstANSERAccpoint.UseVisualStyleBackColor = true
        '
        'chkJisyaTanto
        '
        Me.chkJisyaTanto.AutoSize = true
        Me.chkJisyaTanto.Location = New System.Drawing.Point(25, 53)
        Me.chkJisyaTanto.Name = "chkJisyaTanto"
        Me.chkJisyaTanto.Size = New System.Drawing.Size(111, 22)
        Me.chkJisyaTanto.TabIndex = 11
        Me.chkJisyaTanto.Text = "自社担当者情報"
        Me.chkJisyaTanto.UseVisualStyleBackColor = true
        '
        'chkMstYatinKoza
        '
        Me.chkMstYatinKoza.AutoSize = true
        Me.chkMstYatinKoza.Location = New System.Drawing.Point(176, 137)
        Me.chkMstYatinKoza.Name = "chkMstYatinKoza"
        Me.chkMstYatinKoza.Size = New System.Drawing.Size(123, 22)
        Me.chkMstYatinKoza.TabIndex = 10
        Me.chkMstYatinKoza.Text = "家賃入金口座情報"
        Me.chkMstYatinKoza.UseVisualStyleBackColor = true
        '
        'chkJisyaKoza
        '
        Me.chkJisyaKoza.AutoSize = true
        Me.chkJisyaKoza.Location = New System.Drawing.Point(25, 109)
        Me.chkJisyaKoza.Name = "chkJisyaKoza"
        Me.chkJisyaKoza.Size = New System.Drawing.Size(99, 22)
        Me.chkJisyaKoza.TabIndex = 1
        Me.chkJisyaKoza.Text = "自社口座情報"
        Me.chkJisyaKoza.UseVisualStyleBackColor = true
        '
        'chkJisyaBase
        '
        Me.chkJisyaBase.AutoSize = true
        Me.chkJisyaBase.Location = New System.Drawing.Point(25, 25)
        Me.chkJisyaBase.Name = "chkJisyaBase"
        Me.chkJisyaBase.Size = New System.Drawing.Size(99, 22)
        Me.chkJisyaBase.TabIndex = 0
        Me.chkJisyaBase.Text = "自社基本情報"
        Me.chkJisyaBase.UseVisualStyleBackColor = true
        '
        'chkFBFuriirai
        '
        Me.chkFBFuriirai.AutoSize = true
        Me.chkFBFuriirai.Location = New System.Drawing.Point(56, 137)
        Me.chkFBFuriirai.Name = "chkFBFuriirai"
        Me.chkFBFuriirai.Size = New System.Drawing.Size(111, 22)
        Me.chkFBFuriirai.TabIndex = 2
        Me.chkFBFuriirai.Text = "振込依頼人情報"
        Me.chkFBFuriirai.UseVisualStyleBackColor = true
        '
        'chkFBNsSyutoku
        '
        Me.chkFBNsSyutoku.AutoSize = true
        Me.chkFBNsSyutoku.Location = New System.Drawing.Point(56, 221)
        Me.chkFBNsSyutoku.Name = "chkFBNsSyutoku"
        Me.chkFBNsSyutoku.Size = New System.Drawing.Size(111, 22)
        Me.chkFBNsSyutoku.TabIndex = 9
        Me.chkFBNsSyutoku.Text = "入出金取得情報"
        Me.chkFBNsSyutoku.UseVisualStyleBackColor = true
        '
        'chkJisyaMemo
        '
        Me.chkJisyaMemo.AutoSize = true
        Me.chkJisyaMemo.Location = New System.Drawing.Point(25, 81)
        Me.chkJisyaMemo.Name = "chkJisyaMemo"
        Me.chkJisyaMemo.Size = New System.Drawing.Size(99, 22)
        Me.chkJisyaMemo.TabIndex = 4
        Me.chkJisyaMemo.Text = "自社メモ情報"
        Me.chkJisyaMemo.UseVisualStyleBackColor = true
        '
        'chkFBKozafurikae
        '
        Me.chkFBKozafurikae.AutoSize = true
        Me.chkFBKozafurikae.Location = New System.Drawing.Point(56, 165)
        Me.chkFBKozafurikae.Name = "chkFBKozafurikae"
        Me.chkFBKozafurikae.Size = New System.Drawing.Size(99, 22)
        Me.chkFBKozafurikae.TabIndex = 3
        Me.chkFBKozafurikae.Text = "口座振替情報"
        Me.chkFBKozafurikae.UseVisualStyleBackColor = true
        '
        'chkFBFuritesuryo
        '
        Me.chkFBFuritesuryo.AutoSize = true
        Me.chkFBFuritesuryo.Location = New System.Drawing.Point(56, 193)
        Me.chkFBFuritesuryo.Name = "chkFBFuritesuryo"
        Me.chkFBFuritesuryo.Size = New System.Drawing.Size(111, 22)
        Me.chkFBFuritesuryo.TabIndex = 8
        Me.chkFBFuritesuryo.Text = "振込手数料情報"
        Me.chkFBFuritesuryo.UseVisualStyleBackColor = true
        '
        'tabPageBase140
        '
        Me.tabPageBase140.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase140.Controls.Add(Me.grpBk)
        Me.tabPageBase140.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase140.Name = "tabPageBase140"
        Me.tabPageBase140.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase140.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase140.TabIndex = 4
        Me.tabPageBase140.Text = "物件情報"
        '
        'grpBk
        '
        Me.grpBk.Controls.Add(Me.chkOpKys)
        Me.grpBk.Controls.Add(Me.chkOpOw)
        Me.grpBk.Controls.Add(Me.grpKagiSelect)
        Me.grpBk.Controls.Add(Me.chkBkSyo)
        Me.grpBk.Controls.Add(Me.chkBkHendo)
        Me.grpBk.Controls.Add(Me.chkBkKinrincyusyajo)
        Me.grpBk.Controls.Add(Me.chkBkSansyofile)
        Me.grpBk.Controls.Add(Me.chkBkSzeniji)
        Me.grpBk.Controls.Add(Me.chkBkSyuhen)
        Me.grpBk.Controls.Add(Me.chkBkSetudo)
        Me.grpBk.Controls.Add(Me.chkBkKotu)
        Me.grpBk.Controls.Add(Me.chkBkKenri)
        Me.grpBk.Controls.Add(Me.chkBkGomi)
        Me.grpBk.Controls.Add(Me.chkBkSyosai)
        Me.grpBk.Controls.Add(Me.chkBkKagi)
        Me.grpBk.Controls.Add(Me.chkBkMemo)
        Me.grpBk.Controls.Add(Me.chkBkBase)
        Me.grpBk.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpBk.Location = New System.Drawing.Point(10, 5)
        Me.grpBk.Name = "grpBk"
        Me.grpBk.Size = New System.Drawing.Size(817, 265)
        Me.grpBk.TabIndex = 0
        Me.grpBk.TabStop = false
        Me.grpBk.Text = "物件情報"
        '
        'chkOpKys
        '
        Me.chkOpKys.AutoSize = true
        Me.chkOpKys.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOpKys.Location = New System.Drawing.Point(545, 179)
        Me.chkOpKys.Name = "chkOpKys"
        Me.chkOpKys.Size = New System.Drawing.Size(229, 22)
        Me.chkOpKys.TabIndex = 16
        Me.chkOpKys.Text = "(未使用の契約者データは移行しない)"
        Me.chkOpKys.UseVisualStyleBackColor = true
        '
        'chkOpOw
        '
        Me.chkOpOw.AutoSize = true
        Me.chkOpOw.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkOpOw.Location = New System.Drawing.Point(545, 151)
        Me.chkOpOw.Name = "chkOpOw"
        Me.chkOpOw.Size = New System.Drawing.Size(217, 22)
        Me.chkOpOw.TabIndex = 15
        Me.chkOpOw.Text = "(未使用の家主データは移行しない)"
        Me.chkOpOw.UseVisualStyleBackColor = true
        '
        'grpKagiSelect
        '
        Me.grpKagiSelect.Controls.Add(Me.optKyKagi)
        Me.grpKagiSelect.Controls.Add(Me.optHyKagi)
        Me.grpKagiSelect.Location = New System.Drawing.Point(557, 81)
        Me.grpKagiSelect.Name = "grpKagiSelect"
        Me.grpKagiSelect.Size = New System.Drawing.Size(240, 50)
        Me.grpKagiSelect.TabIndex = 14
        Me.grpKagiSelect.TabStop = false
        Me.grpKagiSelect.Text = " 取得する鍵情報 "
        '
        'optKyKagi
        '
        Me.optKyKagi.AutoSize = true
        Me.optKyKagi.Location = New System.Drawing.Point(140, 20)
        Me.optKyKagi.Name = "optKyKagi"
        Me.optKyKagi.Size = New System.Drawing.Size(86, 22)
        Me.optKyKagi.TabIndex = 1
        Me.optKyKagi.TabStop = true
        Me.optKyKagi.Text = "契約鍵情報"
        Me.optKyKagi.UseVisualStyleBackColor = true
        '
        'optHyKagi
        '
        Me.optHyKagi.AutoSize = true
        Me.optHyKagi.Checked = true
        Me.optHyKagi.Location = New System.Drawing.Point(20, 20)
        Me.optHyKagi.Name = "optHyKagi"
        Me.optHyKagi.Size = New System.Drawing.Size(86, 22)
        Me.optHyKagi.TabIndex = 0
        Me.optHyKagi.TabStop = true
        Me.optHyKagi.Text = "部屋鍵情報"
        Me.optHyKagi.UseVisualStyleBackColor = true
        '
        'chkBkSyo
        '
        Me.chkBkSyo.AutoSize = true
        Me.chkBkSyo.Location = New System.Drawing.Point(295, 193)
        Me.chkBkSyo.Name = "chkBkSyo"
        Me.chkBkSyo.Size = New System.Drawing.Size(111, 22)
        Me.chkBkSyo.TabIndex = 12
        Me.chkBkSyo.Text = "物件所有者情報"
        Me.chkBkSyo.UseVisualStyleBackColor = true
        '
        'chkBkHendo
        '
        Me.chkBkHendo.AutoSize = true
        Me.chkBkHendo.Location = New System.Drawing.Point(295, 109)
        Me.chkBkHendo.Name = "chkBkHendo"
        Me.chkBkHendo.Size = New System.Drawing.Size(171, 22)
        Me.chkBkHendo.TabIndex = 9
        Me.chkBkHendo.Text = "物件変動費親メーター情報"
        Me.chkBkHendo.UseVisualStyleBackColor = true
        '
        'chkBkKinrincyusyajo
        '
        Me.chkBkKinrincyusyajo.AutoSize = true
        Me.chkBkKinrincyusyajo.Location = New System.Drawing.Point(295, 137)
        Me.chkBkKinrincyusyajo.Name = "chkBkKinrincyusyajo"
        Me.chkBkKinrincyusyajo.Size = New System.Drawing.Size(135, 22)
        Me.chkBkKinrincyusyajo.TabIndex = 10
        Me.chkBkKinrincyusyajo.Text = "物件近隣駐車場情報"
        Me.chkBkKinrincyusyajo.UseVisualStyleBackColor = true
        '
        'chkBkSansyofile
        '
        Me.chkBkSansyofile.AutoSize = true
        Me.chkBkSansyofile.Location = New System.Drawing.Point(295, 165)
        Me.chkBkSansyofile.Name = "chkBkSansyofile"
        Me.chkBkSansyofile.Size = New System.Drawing.Size(147, 22)
        Me.chkBkSansyofile.TabIndex = 11
        Me.chkBkSansyofile.Text = "物件参照ファイル情報"
        Me.chkBkSansyofile.UseVisualStyleBackColor = true
        '
        'chkBkSzeniji
        '
        Me.chkBkSzeniji.AutoSize = true
        Me.chkBkSzeniji.Location = New System.Drawing.Point(295, 81)
        Me.chkBkSzeniji.Name = "chkBkSzeniji"
        Me.chkBkSzeniji.Size = New System.Drawing.Size(183, 22)
        Me.chkBkSzeniji.TabIndex = 8
        Me.chkBkSzeniji.Text = "物件修繕維持管理連絡先情報"
        Me.chkBkSzeniji.UseVisualStyleBackColor = true
        '
        'chkBkSyuhen
        '
        Me.chkBkSyuhen.AutoSize = true
        Me.chkBkSyuhen.Location = New System.Drawing.Point(45, 81)
        Me.chkBkSyuhen.Name = "chkBkSyuhen"
        Me.chkBkSyuhen.Size = New System.Drawing.Size(99, 22)
        Me.chkBkSyuhen.TabIndex = 2
        Me.chkBkSyuhen.Text = "物件周辺情報"
        Me.chkBkSyuhen.UseVisualStyleBackColor = true
        '
        'chkBkSetudo
        '
        Me.chkBkSetudo.AutoSize = true
        Me.chkBkSetudo.Location = New System.Drawing.Point(45, 193)
        Me.chkBkSetudo.Name = "chkBkSetudo"
        Me.chkBkSetudo.Size = New System.Drawing.Size(99, 22)
        Me.chkBkSetudo.TabIndex = 6
        Me.chkBkSetudo.Text = "物件接道情報"
        Me.chkBkSetudo.UseVisualStyleBackColor = true
        '
        'chkBkKotu
        '
        Me.chkBkKotu.AutoSize = true
        Me.chkBkKotu.Location = New System.Drawing.Point(45, 165)
        Me.chkBkKotu.Name = "chkBkKotu"
        Me.chkBkKotu.Size = New System.Drawing.Size(99, 22)
        Me.chkBkKotu.TabIndex = 5
        Me.chkBkKotu.Text = "物件交通情報"
        Me.chkBkKotu.UseVisualStyleBackColor = true
        '
        'chkBkKenri
        '
        Me.chkBkKenri.AutoSize = true
        Me.chkBkKenri.Location = New System.Drawing.Point(45, 137)
        Me.chkBkKenri.Name = "chkBkKenri"
        Me.chkBkKenri.Size = New System.Drawing.Size(99, 22)
        Me.chkBkKenri.TabIndex = 4
        Me.chkBkKenri.Text = "物件権利情報"
        Me.chkBkKenri.UseVisualStyleBackColor = true
        '
        'chkBkGomi
        '
        Me.chkBkGomi.AutoSize = true
        Me.chkBkGomi.Location = New System.Drawing.Point(45, 109)
        Me.chkBkGomi.Name = "chkBkGomi"
        Me.chkBkGomi.Size = New System.Drawing.Size(99, 22)
        Me.chkBkGomi.TabIndex = 3
        Me.chkBkGomi.Text = "物件ゴミ情報"
        Me.chkBkGomi.UseVisualStyleBackColor = true
        '
        'chkBkSyosai
        '
        Me.chkBkSyosai.AutoSize = true
        Me.chkBkSyosai.Location = New System.Drawing.Point(45, 53)
        Me.chkBkSyosai.Name = "chkBkSyosai"
        Me.chkBkSyosai.Size = New System.Drawing.Size(99, 22)
        Me.chkBkSyosai.TabIndex = 1
        Me.chkBkSyosai.Text = "物件詳細情報"
        Me.chkBkSyosai.UseVisualStyleBackColor = true
        '
        'chkBkKagi
        '
        Me.chkBkKagi.AutoSize = true
        Me.chkBkKagi.Location = New System.Drawing.Point(545, 53)
        Me.chkBkKagi.Name = "chkBkKagi"
        Me.chkBkKagi.Size = New System.Drawing.Size(87, 22)
        Me.chkBkKagi.TabIndex = 13
        Me.chkBkKagi.Text = "物件鍵情報"
        Me.chkBkKagi.UseVisualStyleBackColor = true
        '
        'chkBkMemo
        '
        Me.chkBkMemo.AutoSize = true
        Me.chkBkMemo.Location = New System.Drawing.Point(295, 53)
        Me.chkBkMemo.Name = "chkBkMemo"
        Me.chkBkMemo.Size = New System.Drawing.Size(99, 22)
        Me.chkBkMemo.TabIndex = 7
        Me.chkBkMemo.Text = "物件メモ情報"
        Me.chkBkMemo.UseVisualStyleBackColor = true
        '
        'chkBkBase
        '
        Me.chkBkBase.AutoSize = true
        Me.chkBkBase.Location = New System.Drawing.Point(25, 25)
        Me.chkBkBase.Name = "chkBkBase"
        Me.chkBkBase.Size = New System.Drawing.Size(99, 22)
        Me.chkBkBase.TabIndex = 0
        Me.chkBkBase.Text = "物件基本情報"
        Me.chkBkBase.UseVisualStyleBackColor = true
        '
        'tabPageBase150
        '
        Me.tabPageBase150.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase150.Controls.Add(Me.grpHy)
        Me.tabPageBase150.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase150.Name = "tabPageBase150"
        Me.tabPageBase150.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase150.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase150.TabIndex = 14
        Me.tabPageBase150.Text = "部屋情報"
        '
        'grpHy
        '
        Me.grpHy.Controls.Add(Me.chkHySyo)
        Me.grpHy.Controls.Add(Me.chkHySansyofile)
        Me.grpHy.Controls.Add(Me.chkHyGenjotanka)
        Me.grpHy.Controls.Add(Me.chkHyKenri)
        Me.grpHy.Controls.Add(Me.chkHyConfirm)
        Me.grpHy.Controls.Add(Me.chkHyHendo)
        Me.grpHy.Controls.Add(Me.chkHyCommonsalespoint)
        Me.grpHy.Controls.Add(Me.chkHyMenseki)
        Me.grpHy.Controls.Add(Me.chkHyNkinkomk)
        Me.grpHy.Controls.Add(Me.chkHyMadoriutiwake)
        Me.grpHy.Controls.Add(Me.chkHySzeniji)
        Me.grpHy.Controls.Add(Me.chkHyTokuyaku)
        Me.grpHy.Controls.Add(Me.chkHyParking)
        Me.grpHy.Controls.Add(Me.chkHySyosai)
        Me.grpHy.Controls.Add(Me.chkHyMemo)
        Me.grpHy.Controls.Add(Me.chkHyKagi)
        Me.grpHy.Controls.Add(Me.chkHySetubi)
        Me.grpHy.Controls.Add(Me.chkHyBase)
        Me.grpHy.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpHy.Location = New System.Drawing.Point(10, 5)
        Me.grpHy.Name = "grpHy"
        Me.grpHy.Size = New System.Drawing.Size(817, 265)
        Me.grpHy.TabIndex = 0
        Me.grpHy.TabStop = false
        Me.grpHy.Text = "部屋情報"
        '
        'chkHySyo
        '
        Me.chkHySyo.AutoSize = true
        Me.chkHySyo.Location = New System.Drawing.Point(295, 164)
        Me.chkHySyo.Name = "chkHySyo"
        Me.chkHySyo.Size = New System.Drawing.Size(111, 22)
        Me.chkHySyo.TabIndex = 12
        Me.chkHySyo.Text = "部屋所有者情報"
        Me.chkHySyo.UseVisualStyleBackColor = true
        '
        'chkHySansyofile
        '
        Me.chkHySansyofile.AutoSize = true
        Me.chkHySansyofile.Location = New System.Drawing.Point(545, 81)
        Me.chkHySansyofile.Name = "chkHySansyofile"
        Me.chkHySansyofile.Size = New System.Drawing.Size(147, 22)
        Me.chkHySansyofile.TabIndex = 16
        Me.chkHySansyofile.Text = "部屋参照ファイル情報"
        Me.chkHySansyofile.UseVisualStyleBackColor = true
        '
        'chkHyGenjotanka
        '
        Me.chkHyGenjotanka.AutoSize = true
        Me.chkHyGenjotanka.Location = New System.Drawing.Point(545, 109)
        Me.chkHyGenjotanka.Name = "chkHyGenjotanka"
        Me.chkHyGenjotanka.Size = New System.Drawing.Size(171, 22)
        Me.chkHyGenjotanka.TabIndex = 17
        Me.chkHyGenjotanka.Text = "部屋原状回復目安単価情報"
        Me.chkHyGenjotanka.UseVisualStyleBackColor = true
        '
        'chkHyKenri
        '
        Me.chkHyKenri.AutoSize = true
        Me.chkHyKenri.Location = New System.Drawing.Point(545, 53)
        Me.chkHyKenri.Name = "chkHyKenri"
        Me.chkHyKenri.Size = New System.Drawing.Size(99, 22)
        Me.chkHyKenri.TabIndex = 15
        Me.chkHyKenri.Text = "部屋権利情報"
        Me.chkHyKenri.UseVisualStyleBackColor = true
        '
        'chkHyConfirm
        '
        Me.chkHyConfirm.AutoSize = true
        Me.chkHyConfirm.Location = New System.Drawing.Point(295, 221)
        Me.chkHyConfirm.Name = "chkHyConfirm"
        Me.chkHyConfirm.Size = New System.Drawing.Size(171, 22)
        Me.chkHyConfirm.TabIndex = 14
        Me.chkHyConfirm.Text = "部屋契約解約確認事項情報"
        Me.chkHyConfirm.UseVisualStyleBackColor = true
        '
        'chkHyHendo
        '
        Me.chkHyHendo.AutoSize = true
        Me.chkHyHendo.Location = New System.Drawing.Point(295, 137)
        Me.chkHyHendo.Name = "chkHyHendo"
        Me.chkHyHendo.Size = New System.Drawing.Size(183, 22)
        Me.chkHyHendo.TabIndex = 11
        Me.chkHyHendo.Text = "部屋変動費各戸メーター情報"
        Me.chkHyHendo.UseVisualStyleBackColor = true
        '
        'chkHyCommonsalespoint
        '
        Me.chkHyCommonsalespoint.AutoSize = true
        Me.chkHyCommonsalespoint.Location = New System.Drawing.Point(295, 193)
        Me.chkHyCommonsalespoint.Name = "chkHyCommonsalespoint"
        Me.chkHyCommonsalespoint.Size = New System.Drawing.Size(195, 22)
        Me.chkHyCommonsalespoint.TabIndex = 13
        Me.chkHyCommonsalespoint.Text = "部屋共通セールスポイント情報"
        Me.chkHyCommonsalespoint.UseVisualStyleBackColor = true
        '
        'chkHyMenseki
        '
        Me.chkHyMenseki.AutoSize = true
        Me.chkHyMenseki.Location = New System.Drawing.Point(45, 109)
        Me.chkHyMenseki.Name = "chkHyMenseki"
        Me.chkHyMenseki.Size = New System.Drawing.Size(99, 22)
        Me.chkHyMenseki.TabIndex = 3
        Me.chkHyMenseki.Text = "部屋面積情報"
        Me.chkHyMenseki.UseVisualStyleBackColor = true
        '
        'chkHyNkinkomk
        '
        Me.chkHyNkinkomk.AutoSize = true
        Me.chkHyNkinkomk.Location = New System.Drawing.Point(295, 81)
        Me.chkHyNkinkomk.Name = "chkHyNkinkomk"
        Me.chkHyNkinkomk.Size = New System.Drawing.Size(123, 22)
        Me.chkHyNkinkomk.TabIndex = 9
        Me.chkHyNkinkomk.Text = "部屋入金項目情報"
        Me.chkHyNkinkomk.UseVisualStyleBackColor = true
        '
        'chkHyMadoriutiwake
        '
        Me.chkHyMadoriutiwake.AutoSize = true
        Me.chkHyMadoriutiwake.Location = New System.Drawing.Point(45, 165)
        Me.chkHyMadoriutiwake.Name = "chkHyMadoriutiwake"
        Me.chkHyMadoriutiwake.Size = New System.Drawing.Size(123, 22)
        Me.chkHyMadoriutiwake.TabIndex = 5
        Me.chkHyMadoriutiwake.Text = "部屋間取内訳情報"
        Me.chkHyMadoriutiwake.UseVisualStyleBackColor = true
        '
        'chkHySzeniji
        '
        Me.chkHySzeniji.AutoSize = true
        Me.chkHySzeniji.Location = New System.Drawing.Point(295, 109)
        Me.chkHySzeniji.Name = "chkHySzeniji"
        Me.chkHySzeniji.Size = New System.Drawing.Size(183, 22)
        Me.chkHySzeniji.TabIndex = 10
        Me.chkHySzeniji.Text = "部屋修繕維持管理連絡先情報"
        Me.chkHySzeniji.UseVisualStyleBackColor = true
        '
        'chkHyTokuyaku
        '
        Me.chkHyTokuyaku.AutoSize = true
        Me.chkHyTokuyaku.Location = New System.Drawing.Point(45, 221)
        Me.chkHyTokuyaku.Name = "chkHyTokuyaku"
        Me.chkHyTokuyaku.Size = New System.Drawing.Size(99, 22)
        Me.chkHyTokuyaku.TabIndex = 7
        Me.chkHyTokuyaku.Text = "部屋特約情報"
        Me.chkHyTokuyaku.UseVisualStyleBackColor = true
        '
        'chkHyParking
        '
        Me.chkHyParking.AutoSize = true
        Me.chkHyParking.Location = New System.Drawing.Point(45, 193)
        Me.chkHyParking.Name = "chkHyParking"
        Me.chkHyParking.Size = New System.Drawing.Size(111, 22)
        Me.chkHyParking.TabIndex = 6
        Me.chkHyParking.Text = "部屋駐車場情報"
        Me.chkHyParking.UseVisualStyleBackColor = true
        '
        'chkHySyosai
        '
        Me.chkHySyosai.AutoSize = true
        Me.chkHySyosai.Location = New System.Drawing.Point(45, 53)
        Me.chkHySyosai.Name = "chkHySyosai"
        Me.chkHySyosai.Size = New System.Drawing.Size(99, 22)
        Me.chkHySyosai.TabIndex = 1
        Me.chkHySyosai.Text = "部屋詳細情報"
        Me.chkHySyosai.UseVisualStyleBackColor = true
        '
        'chkHyMemo
        '
        Me.chkHyMemo.AutoSize = true
        Me.chkHyMemo.Location = New System.Drawing.Point(45, 137)
        Me.chkHyMemo.Name = "chkHyMemo"
        Me.chkHyMemo.Size = New System.Drawing.Size(99, 22)
        Me.chkHyMemo.TabIndex = 4
        Me.chkHyMemo.Text = "部屋メモ情報"
        Me.chkHyMemo.UseVisualStyleBackColor = true
        '
        'chkHyKagi
        '
        Me.chkHyKagi.AutoSize = true
        Me.chkHyKagi.Location = New System.Drawing.Point(295, 53)
        Me.chkHyKagi.Name = "chkHyKagi"
        Me.chkHyKagi.Size = New System.Drawing.Size(87, 22)
        Me.chkHyKagi.TabIndex = 8
        Me.chkHyKagi.Text = "部屋鍵情報"
        Me.chkHyKagi.UseVisualStyleBackColor = true
        '
        'chkHySetubi
        '
        Me.chkHySetubi.AutoSize = true
        Me.chkHySetubi.Location = New System.Drawing.Point(45, 81)
        Me.chkHySetubi.Name = "chkHySetubi"
        Me.chkHySetubi.Size = New System.Drawing.Size(99, 22)
        Me.chkHySetubi.TabIndex = 2
        Me.chkHySetubi.Text = "部屋設備情報"
        Me.chkHySetubi.UseVisualStyleBackColor = true
        '
        'chkHyBase
        '
        Me.chkHyBase.AutoSize = true
        Me.chkHyBase.Location = New System.Drawing.Point(25, 25)
        Me.chkHyBase.Name = "chkHyBase"
        Me.chkHyBase.Size = New System.Drawing.Size(99, 22)
        Me.chkHyBase.TabIndex = 0
        Me.chkHyBase.Text = "部屋基本情報"
        Me.chkHyBase.UseVisualStyleBackColor = true
        '
        'tabPageBase160
        '
        Me.tabPageBase160.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase160.Controls.Add(Me.grpSorule)
        Me.tabPageBase160.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase160.Name = "tabPageBase160"
        Me.tabPageBase160.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase160.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase160.TabIndex = 5
        Me.tabPageBase160.Text = "送金ルール"
        '
        'grpSorule
        '
        Me.grpSorule.Controls.Add(Me.chkSoruleSosaki)
        Me.grpSorule.Controls.Add(Me.chkSoruleKojo)
        Me.grpSorule.Controls.Add(Me.chkSoruleNkin)
        Me.grpSorule.Controls.Add(Me.chkSoruleBase)
        Me.grpSorule.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpSorule.Location = New System.Drawing.Point(10, 5)
        Me.grpSorule.Name = "grpSorule"
        Me.grpSorule.Size = New System.Drawing.Size(282, 265)
        Me.grpSorule.TabIndex = 0
        Me.grpSorule.TabStop = false
        Me.grpSorule.Text = "送金ルール情報"
        '
        'chkSoruleSosaki
        '
        Me.chkSoruleSosaki.AutoSize = true
        Me.chkSoruleSosaki.Location = New System.Drawing.Point(45, 53)
        Me.chkSoruleSosaki.Name = "chkSoruleSosaki"
        Me.chkSoruleSosaki.Size = New System.Drawing.Size(147, 22)
        Me.chkSoruleSosaki.TabIndex = 1
        Me.chkSoruleSosaki.Text = "送金ルール送金先情報"
        Me.chkSoruleSosaki.UseVisualStyleBackColor = true
        '
        'chkSoruleKojo
        '
        Me.chkSoruleKojo.AutoSize = true
        Me.chkSoruleKojo.Location = New System.Drawing.Point(45, 109)
        Me.chkSoruleKojo.Name = "chkSoruleKojo"
        Me.chkSoruleKojo.Size = New System.Drawing.Size(159, 22)
        Me.chkSoruleKojo.TabIndex = 3
        Me.chkSoruleKojo.Text = "送金ルール控除項目情報"
        Me.chkSoruleKojo.UseVisualStyleBackColor = true
        '
        'chkSoruleNkin
        '
        Me.chkSoruleNkin.AutoSize = true
        Me.chkSoruleNkin.Location = New System.Drawing.Point(45, 81)
        Me.chkSoruleNkin.Name = "chkSoruleNkin"
        Me.chkSoruleNkin.Size = New System.Drawing.Size(159, 22)
        Me.chkSoruleNkin.TabIndex = 2
        Me.chkSoruleNkin.Text = "送金ルール入金項目情報"
        Me.chkSoruleNkin.UseVisualStyleBackColor = true
        '
        'chkSoruleBase
        '
        Me.chkSoruleBase.AutoSize = true
        Me.chkSoruleBase.Location = New System.Drawing.Point(25, 25)
        Me.chkSoruleBase.Name = "chkSoruleBase"
        Me.chkSoruleBase.Size = New System.Drawing.Size(135, 22)
        Me.chkSoruleBase.TabIndex = 0
        Me.chkSoruleBase.Text = "送金ルール基本情報"
        Me.chkSoruleBase.UseVisualStyleBackColor = true
        '
        'tabPageBase170
        '
        Me.tabPageBase170.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase170.Controls.Add(Me.grpKy)
        Me.tabPageBase170.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase170.Name = "tabPageBase170"
        Me.tabPageBase170.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase170.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase170.TabIndex = 6
        Me.tabPageBase170.Text = "契約情報"
        '
        'grpKy
        '
        Me.grpKy.Controls.Add(Me.chkKySzenmeisai)
        Me.grpKy.Controls.Add(Me.chkKySzen)
        Me.grpKy.Controls.Add(Me.chkKyKai)
        Me.grpKy.Controls.Add(Me.CheckBox14)
        Me.grpKy.Controls.Add(Me.CheckBox24)
        Me.grpKy.Controls.Add(Me.CheckBox23)
        Me.grpKy.Controls.Add(Me.chkKyHosyonin)
        Me.grpKy.Controls.Add(Me.chkKyMemo)
        Me.grpKy.Controls.Add(Me.chkKyTokuyaku)
        Me.grpKy.Controls.Add(Me.CheckBox18)
        Me.grpKy.Controls.Add(Me.chkKySorule)
        Me.grpKy.Controls.Add(Me.CheckBox16)
        Me.grpKy.Controls.Add(Me.CheckBox15)
        Me.grpKy.Controls.Add(Me.chkKyNyukyo)
        Me.grpKy.Controls.Add(Me.chkKyNkinkomkNx)
        Me.grpKy.Controls.Add(Me.chkKyNkinkomk)
        Me.grpKy.Controls.Add(Me.chkKyHoken)
        Me.grpKy.Controls.Add(Me.chkKyKojoRule)
        Me.grpKy.Controls.Add(Me.chkKyKys)
        Me.grpKy.Controls.Add(Me.CheckBox5)
        Me.grpKy.Controls.Add(Me.chkKyRireki)
        Me.grpKy.Controls.Add(Me.chkKyHendo)
        Me.grpKy.Controls.Add(Me.chkKyCar)
        Me.grpKy.Controls.Add(Me.chkKyBase)
        Me.grpKy.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpKy.Location = New System.Drawing.Point(10, 5)
        Me.grpKy.Name = "grpKy"
        Me.grpKy.Size = New System.Drawing.Size(819, 265)
        Me.grpKy.TabIndex = 0
        Me.grpKy.TabStop = false
        Me.grpKy.Text = "契約情報"
        '
        'chkKySzenmeisai
        '
        Me.chkKySzenmeisai.AutoSize = true
        Me.chkKySzenmeisai.Location = New System.Drawing.Point(400, 80)
        Me.chkKySzenmeisai.Name = "chkKySzenmeisai"
        Me.chkKySzenmeisai.Size = New System.Drawing.Size(147, 22)
        Me.chkKySzenmeisai.TabIndex = 23
        Me.chkKySzenmeisai.Text = "契約修繕見積詳細情報"
        Me.chkKySzenmeisai.UseVisualStyleBackColor = true
        '
        'chkKySzen
        '
        Me.chkKySzen.AutoSize = true
        Me.chkKySzen.Location = New System.Drawing.Point(400, 48)
        Me.chkKySzen.Name = "chkKySzen"
        Me.chkKySzen.Size = New System.Drawing.Size(123, 22)
        Me.chkKySzen.TabIndex = 22
        Me.chkKySzen.Text = "契約修繕見積情報"
        Me.chkKySzen.UseVisualStyleBackColor = true
        '
        'chkKyKai
        '
        Me.chkKyKai.AutoSize = true
        Me.chkKyKai.Location = New System.Drawing.Point(220, 221)
        Me.chkKyKai.Name = "chkKyKai"
        Me.chkKyKai.Size = New System.Drawing.Size(99, 22)
        Me.chkKyKai.TabIndex = 14
        Me.chkKyKai.Text = "契約解約情報"
        Me.chkKyKai.UseVisualStyleBackColor = true
        '
        'CheckBox14
        '
        Me.CheckBox14.AutoSize = true
        Me.CheckBox14.Location = New System.Drawing.Point(576, 104)
        Me.CheckBox14.Name = "CheckBox14"
        Me.CheckBox14.Size = New System.Drawing.Size(171, 22)
        Me.CheckBox14.TabIndex = 17
        Me.CheckBox14.Text = "契約変動費親メーター情報"
        Me.CheckBox14.UseVisualStyleBackColor = true
        '
        'CheckBox24
        '
        Me.CheckBox24.AutoSize = true
        Me.CheckBox24.Location = New System.Drawing.Point(576, 48)
        Me.CheckBox24.Name = "CheckBox24"
        Me.CheckBox24.Size = New System.Drawing.Size(147, 22)
        Me.CheckBox24.TabIndex = 15
        Me.CheckBox24.Text = "契約解約確認事項情報"
        Me.CheckBox24.UseVisualStyleBackColor = true
        '
        'CheckBox23
        '
        Me.CheckBox23.AutoSize = true
        Me.CheckBox23.Location = New System.Drawing.Point(576, 76)
        Me.CheckBox23.Name = "CheckBox23"
        Me.CheckBox23.Size = New System.Drawing.Size(123, 22)
        Me.CheckBox23.TabIndex = 16
        Me.CheckBox23.Text = "契約同時契約情報"
        Me.CheckBox23.UseVisualStyleBackColor = true
        '
        'chkKyHosyonin
        '
        Me.chkKyHosyonin.AutoSize = true
        Me.chkKyHosyonin.Location = New System.Drawing.Point(45, 137)
        Me.chkKyHosyonin.Name = "chkKyHosyonin"
        Me.chkKyHosyonin.Size = New System.Drawing.Size(111, 22)
        Me.chkKyHosyonin.TabIndex = 4
        Me.chkKyHosyonin.Text = "契約保証人情報"
        Me.chkKyHosyonin.UseVisualStyleBackColor = true
        '
        'chkKyMemo
        '
        Me.chkKyMemo.AutoSize = true
        Me.chkKyMemo.Location = New System.Drawing.Point(220, 53)
        Me.chkKyMemo.Name = "chkKyMemo"
        Me.chkKyMemo.Size = New System.Drawing.Size(99, 22)
        Me.chkKyMemo.TabIndex = 8
        Me.chkKyMemo.Text = "契約メモ情報"
        Me.chkKyMemo.UseVisualStyleBackColor = true
        '
        'chkKyTokuyaku
        '
        Me.chkKyTokuyaku.AutoSize = true
        Me.chkKyTokuyaku.Location = New System.Drawing.Point(45, 221)
        Me.chkKyTokuyaku.Name = "chkKyTokuyaku"
        Me.chkKyTokuyaku.Size = New System.Drawing.Size(123, 22)
        Me.chkKyTokuyaku.TabIndex = 7
        Me.chkKyTokuyaku.Text = "契約特約事項情報"
        Me.chkKyTokuyaku.UseVisualStyleBackColor = true
        '
        'CheckBox18
        '
        Me.CheckBox18.AutoSize = true
        Me.CheckBox18.Location = New System.Drawing.Point(576, 216)
        Me.CheckBox18.Name = "CheckBox18"
        Me.CheckBox18.Size = New System.Drawing.Size(171, 22)
        Me.CheckBox18.TabIndex = 21
        Me.CheckBox18.Text = "契約原状回復目安単価情報"
        Me.CheckBox18.UseVisualStyleBackColor = true
        '
        'chkKySorule
        '
        Me.chkKySorule.AutoSize = true
        Me.chkKySorule.Location = New System.Drawing.Point(220, 192)
        Me.chkKySorule.Name = "chkKySorule"
        Me.chkKySorule.Size = New System.Drawing.Size(135, 22)
        Me.chkKySorule.TabIndex = 13
        Me.chkKySorule.Text = "契約送金ルール情報"
        Me.chkKySorule.UseVisualStyleBackColor = true
        '
        'CheckBox16
        '
        Me.CheckBox16.AutoSize = true
        Me.CheckBox16.Location = New System.Drawing.Point(576, 187)
        Me.CheckBox16.Name = "CheckBox16"
        Me.CheckBox16.Size = New System.Drawing.Size(183, 22)
        Me.CheckBox16.TabIndex = 20
        Me.CheckBox16.Text = "契約敷金保証金随時処理情報"
        Me.CheckBox16.UseVisualStyleBackColor = true
        '
        'CheckBox15
        '
        Me.CheckBox15.AutoSize = true
        Me.CheckBox15.Location = New System.Drawing.Point(576, 159)
        Me.CheckBox15.Name = "CheckBox15"
        Me.CheckBox15.Size = New System.Drawing.Size(147, 22)
        Me.CheckBox15.TabIndex = 19
        Me.CheckBox15.Text = "契約関連ファイル情報"
        Me.CheckBox15.UseVisualStyleBackColor = true
        '
        'chkKyNyukyo
        '
        Me.chkKyNyukyo.AutoSize = true
        Me.chkKyNyukyo.Location = New System.Drawing.Point(45, 109)
        Me.chkKyNyukyo.Name = "chkKyNyukyo"
        Me.chkKyNyukyo.Size = New System.Drawing.Size(111, 22)
        Me.chkKyNyukyo.TabIndex = 3
        Me.chkKyNyukyo.Text = "契約入居者情報"
        Me.chkKyNyukyo.UseVisualStyleBackColor = true
        '
        'chkKyNkinkomkNx
        '
        Me.chkKyNkinkomkNx.AutoSize = true
        Me.chkKyNkinkomkNx.Location = New System.Drawing.Point(220, 109)
        Me.chkKyNkinkomkNx.Name = "chkKyNkinkomkNx"
        Me.chkKyNkinkomkNx.Size = New System.Drawing.Size(147, 22)
        Me.chkKyNkinkomkNx.TabIndex = 10
        Me.chkKyNkinkomkNx.Text = "契約次回入金項目情報"
        Me.chkKyNkinkomkNx.UseVisualStyleBackColor = true
        '
        'chkKyNkinkomk
        '
        Me.chkKyNkinkomk.AutoSize = true
        Me.chkKyNkinkomk.Location = New System.Drawing.Point(220, 81)
        Me.chkKyNkinkomk.Name = "chkKyNkinkomk"
        Me.chkKyNkinkomk.Size = New System.Drawing.Size(123, 22)
        Me.chkKyNkinkomk.TabIndex = 9
        Me.chkKyNkinkomk.Text = "契約入金項目情報"
        Me.chkKyNkinkomk.UseVisualStyleBackColor = true
        '
        'chkKyHoken
        '
        Me.chkKyHoken.AutoSize = true
        Me.chkKyHoken.Location = New System.Drawing.Point(45, 193)
        Me.chkKyHoken.Name = "chkKyHoken"
        Me.chkKyHoken.Size = New System.Drawing.Size(99, 22)
        Me.chkKyHoken.TabIndex = 6
        Me.chkKyHoken.Text = "契約保険情報"
        Me.chkKyHoken.UseVisualStyleBackColor = true
        '
        'chkKyKojoRule
        '
        Me.chkKyKojoRule.AutoSize = true
        Me.chkKyKojoRule.Location = New System.Drawing.Point(220, 164)
        Me.chkKyKojoRule.Name = "chkKyKojoRule"
        Me.chkKyKojoRule.Size = New System.Drawing.Size(135, 22)
        Me.chkKyKojoRule.TabIndex = 12
        Me.chkKyKojoRule.Text = "契約控除ルール情報"
        Me.chkKyKojoRule.UseVisualStyleBackColor = true
        '
        'chkKyKys
        '
        Me.chkKyKys.AutoSize = true
        Me.chkKyKys.Location = New System.Drawing.Point(45, 81)
        Me.chkKyKys.Name = "chkKyKys"
        Me.chkKyKys.Size = New System.Drawing.Size(111, 22)
        Me.chkKyKys.TabIndex = 2
        Me.chkKyKys.Text = "契約契約者情報"
        Me.chkKyKys.UseVisualStyleBackColor = true
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = true
        Me.CheckBox5.Location = New System.Drawing.Point(576, 132)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(123, 22)
        Me.CheckBox5.TabIndex = 18
        Me.CheckBox5.Text = "契約空室待ち情報"
        Me.CheckBox5.UseVisualStyleBackColor = true
        '
        'chkKyRireki
        '
        Me.chkKyRireki.AutoSize = true
        Me.chkKyRireki.Location = New System.Drawing.Point(45, 53)
        Me.chkKyRireki.Name = "chkKyRireki"
        Me.chkKyRireki.Size = New System.Drawing.Size(99, 22)
        Me.chkKyRireki.TabIndex = 1
        Me.chkKyRireki.Text = "契約履歴情報"
        Me.chkKyRireki.UseVisualStyleBackColor = true
        '
        'chkKyHendo
        '
        Me.chkKyHendo.AutoSize = true
        Me.chkKyHendo.Location = New System.Drawing.Point(220, 137)
        Me.chkKyHendo.Name = "chkKyHendo"
        Me.chkKyHendo.Size = New System.Drawing.Size(183, 22)
        Me.chkKyHendo.TabIndex = 11
        Me.chkKyHendo.Text = "契約変動費各戸メーター情報"
        Me.chkKyHendo.UseVisualStyleBackColor = true
        '
        'chkKyCar
        '
        Me.chkKyCar.AutoSize = true
        Me.chkKyCar.Location = New System.Drawing.Point(45, 165)
        Me.chkKyCar.Name = "chkKyCar"
        Me.chkKyCar.Size = New System.Drawing.Size(87, 22)
        Me.chkKyCar.TabIndex = 5
        Me.chkKyCar.Text = "契約車情報"
        Me.chkKyCar.UseVisualStyleBackColor = true
        '
        'chkKyBase
        '
        Me.chkKyBase.AutoSize = true
        Me.chkKyBase.Location = New System.Drawing.Point(25, 25)
        Me.chkKyBase.Name = "chkKyBase"
        Me.chkKyBase.Size = New System.Drawing.Size(99, 22)
        Me.chkKyBase.TabIndex = 0
        Me.chkKyBase.Text = "契約基本情報"
        Me.chkKyBase.UseVisualStyleBackColor = true
        '
        'tabPageBase180
        '
        Me.tabPageBase180.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase180.Controls.Add(Me.grpSq)
        Me.tabPageBase180.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase180.Name = "tabPageBase180"
        Me.tabPageBase180.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase180.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase180.TabIndex = 7
        Me.tabPageBase180.Text = "請求情報"
        '
        'grpSq
        '
        Me.grpSq.Controls.Add(Me.chkSqSqKojo)
        Me.grpSq.Controls.Add(Me.chkSqKoteiKojo)
        Me.grpSq.Controls.Add(Me.chkSqHendokensin)
        Me.grpSq.Controls.Add(Me.chkSqUnyotaino)
        Me.grpSq.Controls.Add(Me.chkSqSq)
        Me.grpSq.Controls.Add(Me.chkSqKajyo)
        Me.grpSq.Location = New System.Drawing.Point(10, 5)
        Me.grpSq.Name = "grpSq"
        Me.grpSq.Size = New System.Drawing.Size(282, 265)
        Me.grpSq.TabIndex = 0
        Me.grpSq.TabStop = false
        Me.grpSq.Text = "請求情報"
        '
        'chkSqSqKojo
        '
        Me.chkSqSqKojo.AutoSize = true
        Me.chkSqSqKojo.Location = New System.Drawing.Point(25, 165)
        Me.chkSqSqKojo.Name = "chkSqSqKojo"
        Me.chkSqSqKojo.Size = New System.Drawing.Size(123, 22)
        Me.chkSqSqKojo.TabIndex = 5
        Me.chkSqSqKojo.Text = "家主請求控除情報"
        Me.chkSqSqKojo.UseVisualStyleBackColor = true
        '
        'chkSqKoteiKojo
        '
        Me.chkSqKoteiKojo.AutoSize = true
        Me.chkSqKoteiKojo.Location = New System.Drawing.Point(25, 137)
        Me.chkSqKoteiKojo.Name = "chkSqKoteiKojo"
        Me.chkSqKoteiKojo.Size = New System.Drawing.Size(123, 22)
        Me.chkSqKoteiKojo.TabIndex = 4
        Me.chkSqKoteiKojo.Text = "家主固定控除情報"
        Me.chkSqKoteiKojo.UseVisualStyleBackColor = true
        '
        'chkSqHendokensin
        '
        Me.chkSqHendokensin.AutoSize = true
        Me.chkSqHendokensin.Location = New System.Drawing.Point(25, 109)
        Me.chkSqHendokensin.Name = "chkSqHendokensin"
        Me.chkSqHendokensin.Size = New System.Drawing.Size(111, 22)
        Me.chkSqHendokensin.TabIndex = 3
        Me.chkSqHendokensin.Text = "変動費検針情報"
        Me.chkSqHendokensin.UseVisualStyleBackColor = true
        '
        'chkSqUnyotaino
        '
        Me.chkSqUnyotaino.AutoSize = true
        Me.chkSqUnyotaino.Location = New System.Drawing.Point(25, 53)
        Me.chkSqUnyotaino.Name = "chkSqUnyotaino"
        Me.chkSqUnyotaino.Size = New System.Drawing.Size(111, 22)
        Me.chkSqUnyotaino.TabIndex = 1
        Me.chkSqUnyotaino.Text = "未収滞納金情報"
        Me.chkSqUnyotaino.UseVisualStyleBackColor = true
        '
        'chkSqSq
        '
        Me.chkSqSq.AutoSize = true
        Me.chkSqSq.Location = New System.Drawing.Point(25, 81)
        Me.chkSqSq.Name = "chkSqSq"
        Me.chkSqSq.Size = New System.Drawing.Size(111, 22)
        Me.chkSqSq.TabIndex = 2
        Me.chkSqSq.Text = "その他請求情報"
        Me.chkSqSq.UseVisualStyleBackColor = true
        '
        'chkSqKajyo
        '
        Me.chkSqKajyo.AutoSize = true
        Me.chkSqKajyo.Location = New System.Drawing.Point(25, 25)
        Me.chkSqKajyo.Name = "chkSqKajyo"
        Me.chkSqKajyo.Size = New System.Drawing.Size(87, 22)
        Me.chkSqKajyo.TabIndex = 0
        Me.chkSqKajyo.Text = "預り金情報"
        Me.chkSqKajyo.UseVisualStyleBackColor = true
        '
        'tabPageBase190
        '
        Me.tabPageBase190.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase190.Controls.Add(Me.grpSzen)
        Me.tabPageBase190.Controls.Add(Me.grpClaim)
        Me.tabPageBase190.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase190.Name = "tabPageBase190"
        Me.tabPageBase190.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase190.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase190.TabIndex = 18
        Me.tabPageBase190.Text = "クレーム修繕情報"
        '
        'grpSzen
        '
        Me.grpSzen.Controls.Add(Me.chkSzenRelfile)
        Me.grpSzen.Controls.Add(Me.chkSzenClaim)
        Me.grpSzen.Controls.Add(Me.chkSzenMemo)
        Me.grpSzen.Controls.Add(Me.chkSzenSzen)
        Me.grpSzen.Controls.Add(Me.chkSzenSzenmeisai)
        Me.grpSzen.Controls.Add(Me.chkSzenBase)
        Me.grpSzen.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpSzen.Location = New System.Drawing.Point(312, 8)
        Me.grpSzen.Name = "grpSzen"
        Me.grpSzen.Size = New System.Drawing.Size(282, 265)
        Me.grpSzen.TabIndex = 2
        Me.grpSzen.TabStop = false
        Me.grpSzen.Text = "修繕情報"
        '
        'chkSzenRelfile
        '
        Me.chkSzenRelfile.AutoSize = true
        Me.chkSzenRelfile.Location = New System.Drawing.Point(48, 144)
        Me.chkSzenRelfile.Name = "chkSzenRelfile"
        Me.chkSzenRelfile.Size = New System.Drawing.Size(147, 22)
        Me.chkSzenRelfile.TabIndex = 6
        Me.chkSzenRelfile.Text = "修繕関連ファイル情報"
        Me.chkSzenRelfile.UseVisualStyleBackColor = true
        '
        'chkSzenClaim
        '
        Me.chkSzenClaim.AutoSize = true
        Me.chkSzenClaim.Location = New System.Drawing.Point(48, 112)
        Me.chkSzenClaim.Name = "chkSzenClaim"
        Me.chkSzenClaim.Size = New System.Drawing.Size(171, 22)
        Me.chkSzenClaim.TabIndex = 4
        Me.chkSzenClaim.Text = "修繕クレーム関連付け情報"
        Me.chkSzenClaim.UseVisualStyleBackColor = true
        '
        'chkSzenMemo
        '
        Me.chkSzenMemo.AutoSize = true
        Me.chkSzenMemo.Location = New System.Drawing.Point(48, 176)
        Me.chkSzenMemo.Name = "chkSzenMemo"
        Me.chkSzenMemo.Size = New System.Drawing.Size(99, 22)
        Me.chkSzenMemo.TabIndex = 3
        Me.chkSzenMemo.Text = "修繕メモ情報"
        Me.chkSzenMemo.UseVisualStyleBackColor = true
        '
        'chkSzenSzen
        '
        Me.chkSzenSzen.AutoSize = true
        Me.chkSzenSzen.Location = New System.Drawing.Point(45, 53)
        Me.chkSzenSzen.Name = "chkSzenSzen"
        Me.chkSzenSzen.Size = New System.Drawing.Size(99, 22)
        Me.chkSzenSzen.TabIndex = 1
        Me.chkSzenSzen.Text = "修繕見積情報"
        Me.chkSzenSzen.UseVisualStyleBackColor = true
        '
        'chkSzenSzenmeisai
        '
        Me.chkSzenSzenmeisai.AutoSize = true
        Me.chkSzenSzenmeisai.Location = New System.Drawing.Point(45, 81)
        Me.chkSzenSzenmeisai.Name = "chkSzenSzenmeisai"
        Me.chkSzenSzenmeisai.Size = New System.Drawing.Size(123, 22)
        Me.chkSzenSzenmeisai.TabIndex = 2
        Me.chkSzenSzenmeisai.Text = "修繕見積詳細情報"
        Me.chkSzenSzenmeisai.UseVisualStyleBackColor = true
        '
        'chkSzenBase
        '
        Me.chkSzenBase.AutoSize = true
        Me.chkSzenBase.Location = New System.Drawing.Point(25, 25)
        Me.chkSzenBase.Name = "chkSzenBase"
        Me.chkSzenBase.Size = New System.Drawing.Size(99, 22)
        Me.chkSzenBase.TabIndex = 0
        Me.chkSzenBase.Text = "修繕基本情報"
        Me.chkSzenBase.UseVisualStyleBackColor = true
        '
        'grpClaim
        '
        Me.grpClaim.Controls.Add(Me.chkClaimTaiorireki)
        Me.grpClaim.Controls.Add(Me.chkClaimRelfile)
        Me.grpClaim.Controls.Add(Me.chkClaimBase)
        Me.grpClaim.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpClaim.Location = New System.Drawing.Point(16, 8)
        Me.grpClaim.Name = "grpClaim"
        Me.grpClaim.Size = New System.Drawing.Size(282, 265)
        Me.grpClaim.TabIndex = 1
        Me.grpClaim.TabStop = false
        Me.grpClaim.Text = "クレーム情報"
        '
        'chkClaimTaiorireki
        '
        Me.chkClaimTaiorireki.AutoSize = true
        Me.chkClaimTaiorireki.Location = New System.Drawing.Point(45, 53)
        Me.chkClaimTaiorireki.Name = "chkClaimTaiorireki"
        Me.chkClaimTaiorireki.Size = New System.Drawing.Size(147, 22)
        Me.chkClaimTaiorireki.TabIndex = 1
        Me.chkClaimTaiorireki.Text = "クレーム対応履歴情報"
        Me.chkClaimTaiorireki.UseVisualStyleBackColor = true
        '
        'chkClaimRelfile
        '
        Me.chkClaimRelfile.AutoSize = true
        Me.chkClaimRelfile.Location = New System.Drawing.Point(45, 81)
        Me.chkClaimRelfile.Name = "chkClaimRelfile"
        Me.chkClaimRelfile.Size = New System.Drawing.Size(171, 22)
        Me.chkClaimRelfile.TabIndex = 2
        Me.chkClaimRelfile.Text = "クレーム関連ファイル情報"
        Me.chkClaimRelfile.UseVisualStyleBackColor = true
        '
        'chkClaimBase
        '
        Me.chkClaimBase.AutoSize = true
        Me.chkClaimBase.Location = New System.Drawing.Point(25, 25)
        Me.chkClaimBase.Name = "chkClaimBase"
        Me.chkClaimBase.Size = New System.Drawing.Size(123, 22)
        Me.chkClaimBase.TabIndex = 0
        Me.chkClaimBase.Text = "クレーム基本情報"
        Me.chkClaimBase.UseVisualStyleBackColor = true
        '
        'tabPageBase200
        '
        Me.tabPageBase200.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase200.Controls.Add(Me.grpSyskanri)
        Me.tabPageBase200.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase200.Name = "tabPageBase200"
        Me.tabPageBase200.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase200.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase200.TabIndex = 19
        Me.tabPageBase200.Text = "初期設定"
        '
        'grpSyskanri
        '
        Me.grpSyskanri.Controls.Add(Me.chkSyskanriNkinkomkmerge)
        Me.grpSyskanri.Controls.Add(Me.chkSyskanriHenkanmoji)
        Me.grpSyskanri.Controls.Add(Me.chkSyskanriZei)
        Me.grpSyskanri.Controls.Add(Me.chkSyskanriBase)
        Me.grpSyskanri.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpSyskanri.Location = New System.Drawing.Point(24, 16)
        Me.grpSyskanri.Name = "grpSyskanri"
        Me.grpSyskanri.Size = New System.Drawing.Size(237, 240)
        Me.grpSyskanri.TabIndex = 4
        Me.grpSyskanri.TabStop = false
        Me.grpSyskanri.Text = "初期設定"
        '
        'chkSyskanriNkinkomkmerge
        '
        Me.chkSyskanriNkinkomkmerge.AutoSize = true
        Me.chkSyskanriNkinkomkmerge.Location = New System.Drawing.Point(32, 120)
        Me.chkSyskanriNkinkomkmerge.Name = "chkSyskanriNkinkomkmerge"
        Me.chkSyskanriNkinkomkmerge.Size = New System.Drawing.Size(123, 22)
        Me.chkSyskanriNkinkomkmerge.TabIndex = 3
        Me.chkSyskanriNkinkomkmerge.Text = "入金項目集約情報"
        Me.chkSyskanriNkinkomkmerge.UseVisualStyleBackColor = true
        '
        'chkSyskanriHenkanmoji
        '
        Me.chkSyskanriHenkanmoji.AutoSize = true
        Me.chkSyskanriHenkanmoji.Location = New System.Drawing.Point(32, 88)
        Me.chkSyskanriHenkanmoji.Name = "chkSyskanriHenkanmoji"
        Me.chkSyskanriHenkanmoji.Size = New System.Drawing.Size(99, 22)
        Me.chkSyskanriHenkanmoji.TabIndex = 2
        Me.chkSyskanriHenkanmoji.Text = "変換文字情報"
        Me.chkSyskanriHenkanmoji.UseVisualStyleBackColor = true
        '
        'chkSyskanriZei
        '
        Me.chkSyskanriZei.AutoSize = true
        Me.chkSyskanriZei.Location = New System.Drawing.Point(32, 56)
        Me.chkSyskanriZei.Name = "chkSyskanriZei"
        Me.chkSyskanriZei.Size = New System.Drawing.Size(87, 22)
        Me.chkSyskanriZei.TabIndex = 1
        Me.chkSyskanriZei.Text = "税編集情報"
        Me.chkSyskanriZei.UseVisualStyleBackColor = true
        '
        'chkSyskanriBase
        '
        Me.chkSyskanriBase.AutoSize = true
        Me.chkSyskanriBase.Location = New System.Drawing.Point(25, 25)
        Me.chkSyskanriBase.Name = "chkSyskanriBase"
        Me.chkSyskanriBase.Size = New System.Drawing.Size(123, 22)
        Me.chkSyskanriBase.TabIndex = 0
        Me.chkSyskanriBase.Text = "初期設定基本情報"
        Me.chkSyskanriBase.UseVisualStyleBackColor = true
        '
        'tabPageBase210
        '
        Me.tabPageBase210.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase210.Controls.Add(Me.grpRendo)
        Me.tabPageBase210.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase210.Name = "tabPageBase210"
        Me.tabPageBase210.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase210.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase210.TabIndex = 21
        Me.tabPageBase210.Text = "物件データ連動情報"
        '
        'grpRendo
        '
        Me.grpRendo.Controls.Add(Me.chkRendoMapdisp)
        Me.grpRendo.Controls.Add(Me.chkRendoBtoBgroup)
        Me.grpRendo.Controls.Add(Me.chkRendoHysosin)
        Me.grpRendo.Controls.Add(Me.chkRendoHyrui)
        Me.grpRendo.Controls.Add(Me.chkRendoKokokuSuumo)
        Me.grpRendo.Controls.Add(Me.chkRendoKokokuAthome)
        Me.grpRendo.Controls.Add(Me.chkRendoKokokuHomes)
        Me.grpRendo.Controls.Add(Me.chkRendoKokokuJisyaweb)
        Me.grpRendo.Controls.Add(Me.chkRendoSosinSuumo)
        Me.grpRendo.Controls.Add(Me.chkRendoSosinAthome)
        Me.grpRendo.Controls.Add(Me.chkRendoSosinHomes)
        Me.grpRendo.Controls.Add(Me.chkRendoSosinJisyaweb)
        Me.grpRendo.Controls.Add(Me.chkRendoSosinBase)
        Me.grpRendo.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpRendo.Location = New System.Drawing.Point(16, 16)
        Me.grpRendo.Name = "grpRendo"
        Me.grpRendo.Size = New System.Drawing.Size(567, 240)
        Me.grpRendo.TabIndex = 7
        Me.grpRendo.TabStop = false
        Me.grpRendo.Text = "物件データ連動情報"
        '
        'chkRendoMapdisp
        '
        Me.chkRendoMapdisp.AutoSize = true
        Me.chkRendoMapdisp.Location = New System.Drawing.Point(360, 137)
        Me.chkRendoMapdisp.Name = "chkRendoMapdisp"
        Me.chkRendoMapdisp.Size = New System.Drawing.Size(147, 22)
        Me.chkRendoMapdisp.TabIndex = 15
        Me.chkRendoMapdisp.Text = "地図表示詳細設定情報"
        Me.chkRendoMapdisp.UseVisualStyleBackColor = true
        '
        'chkRendoBtoBgroup
        '
        Me.chkRendoBtoBgroup.AutoSize = true
        Me.chkRendoBtoBgroup.Location = New System.Drawing.Point(360, 109)
        Me.chkRendoBtoBgroup.Name = "chkRendoBtoBgroup"
        Me.chkRendoBtoBgroup.Size = New System.Drawing.Size(151, 22)
        Me.chkRendoBtoBgroup.TabIndex = 14
        Me.chkRendoBtoBgroup.Text = "BtoBグループ設定情報"
        Me.chkRendoBtoBgroup.UseVisualStyleBackColor = true
        '
        'chkRendoHysosin
        '
        Me.chkRendoHysosin.AutoSize = true
        Me.chkRendoHysosin.Location = New System.Drawing.Point(360, 81)
        Me.chkRendoHysosin.Name = "chkRendoHysosin"
        Me.chkRendoHysosin.Size = New System.Drawing.Size(111, 22)
        Me.chkRendoHysosin.TabIndex = 13
        Me.chkRendoHysosin.Text = "部屋毎送信情報"
        Me.chkRendoHysosin.UseVisualStyleBackColor = true
        '
        'chkRendoHyrui
        '
        Me.chkRendoHyrui.AutoSize = true
        Me.chkRendoHyrui.Location = New System.Drawing.Point(360, 53)
        Me.chkRendoHyrui.Name = "chkRendoHyrui"
        Me.chkRendoHyrui.Size = New System.Drawing.Size(171, 22)
        Me.chkRendoHyrui.TabIndex = 12
        Me.chkRendoHyrui.Text = "ポータル連動部屋分類情報"
        Me.chkRendoHyrui.UseVisualStyleBackColor = true
        '
        'chkRendoKokokuSuumo
        '
        Me.chkRendoKokokuSuumo.AutoSize = true
        Me.chkRendoKokokuSuumo.Location = New System.Drawing.Point(188, 137)
        Me.chkRendoKokokuSuumo.Name = "chkRendoKokokuSuumo"
        Me.chkRendoKokokuSuumo.Size = New System.Drawing.Size(144, 22)
        Me.chkRendoKokokuSuumo.TabIndex = 11
        Me.chkRendoKokokuSuumo.Text = "広告補足SUUMO情報"
        Me.chkRendoKokokuSuumo.UseVisualStyleBackColor = true
        '
        'chkRendoKokokuAthome
        '
        Me.chkRendoKokokuAthome.AutoSize = true
        Me.chkRendoKokokuAthome.Location = New System.Drawing.Point(188, 109)
        Me.chkRendoKokokuAthome.Name = "chkRendoKokokuAthome"
        Me.chkRendoKokokuAthome.Size = New System.Drawing.Size(144, 22)
        Me.chkRendoKokokuAthome.TabIndex = 10
        Me.chkRendoKokokuAthome.Text = "広告補足athome情報"
        Me.chkRendoKokokuAthome.UseVisualStyleBackColor = true
        '
        'chkRendoKokokuHomes
        '
        Me.chkRendoKokokuHomes.AutoSize = true
        Me.chkRendoKokokuHomes.Location = New System.Drawing.Point(188, 81)
        Me.chkRendoKokokuHomes.Name = "chkRendoKokokuHomes"
        Me.chkRendoKokokuHomes.Size = New System.Drawing.Size(142, 22)
        Me.chkRendoKokokuHomes.TabIndex = 9
        Me.chkRendoKokokuHomes.Text = "広告補足HOMES情報"
        Me.chkRendoKokokuHomes.UseVisualStyleBackColor = true
        '
        'chkRendoKokokuJisyaweb
        '
        Me.chkRendoKokokuJisyaweb.AutoSize = true
        Me.chkRendoKokokuJisyaweb.Location = New System.Drawing.Point(188, 53)
        Me.chkRendoKokokuJisyaweb.Name = "chkRendoKokokuJisyaweb"
        Me.chkRendoKokokuJisyaweb.Size = New System.Drawing.Size(147, 22)
        Me.chkRendoKokokuJisyaweb.TabIndex = 8
        Me.chkRendoKokokuJisyaweb.Text = "広告補足自社web情報"
        Me.chkRendoKokokuJisyaweb.UseVisualStyleBackColor = true
        '
        'chkRendoSosinSuumo
        '
        Me.chkRendoSosinSuumo.AutoSize = true
        Me.chkRendoSosinSuumo.Location = New System.Drawing.Point(25, 137)
        Me.chkRendoSosinSuumo.Name = "chkRendoSosinSuumo"
        Me.chkRendoSosinSuumo.Size = New System.Drawing.Size(144, 22)
        Me.chkRendoSosinSuumo.TabIndex = 7
        Me.chkRendoSosinSuumo.Text = "送信設定SUUMO情報"
        Me.chkRendoSosinSuumo.UseVisualStyleBackColor = true
        '
        'chkRendoSosinAthome
        '
        Me.chkRendoSosinAthome.AutoSize = true
        Me.chkRendoSosinAthome.Location = New System.Drawing.Point(25, 109)
        Me.chkRendoSosinAthome.Name = "chkRendoSosinAthome"
        Me.chkRendoSosinAthome.Size = New System.Drawing.Size(144, 22)
        Me.chkRendoSosinAthome.TabIndex = 6
        Me.chkRendoSosinAthome.Text = "送信設定athome情報"
        Me.chkRendoSosinAthome.UseVisualStyleBackColor = true
        '
        'chkRendoSosinHomes
        '
        Me.chkRendoSosinHomes.AutoSize = true
        Me.chkRendoSosinHomes.Location = New System.Drawing.Point(25, 81)
        Me.chkRendoSosinHomes.Name = "chkRendoSosinHomes"
        Me.chkRendoSosinHomes.Size = New System.Drawing.Size(142, 22)
        Me.chkRendoSosinHomes.TabIndex = 5
        Me.chkRendoSosinHomes.Text = "送信設定HOMES情報"
        Me.chkRendoSosinHomes.UseVisualStyleBackColor = true
        '
        'chkRendoSosinJisyaweb
        '
        Me.chkRendoSosinJisyaweb.AutoSize = true
        Me.chkRendoSosinJisyaweb.Location = New System.Drawing.Point(25, 53)
        Me.chkRendoSosinJisyaweb.Name = "chkRendoSosinJisyaweb"
        Me.chkRendoSosinJisyaweb.Size = New System.Drawing.Size(147, 22)
        Me.chkRendoSosinJisyaweb.TabIndex = 1
        Me.chkRendoSosinJisyaweb.Text = "送信設定自社web情報"
        Me.chkRendoSosinJisyaweb.UseVisualStyleBackColor = true
        '
        'chkRendoSosinBase
        '
        Me.chkRendoSosinBase.AutoSize = true
        Me.chkRendoSosinBase.Location = New System.Drawing.Point(25, 25)
        Me.chkRendoSosinBase.Name = "chkRendoSosinBase"
        Me.chkRendoSosinBase.Size = New System.Drawing.Size(123, 22)
        Me.chkRendoSosinBase.TabIndex = 0
        Me.chkRendoSosinBase.Text = "送信設定基本情報"
        Me.chkRendoSosinBase.UseVisualStyleBackColor = true
        '
        'tabPageBase900
        '
        Me.tabPageBase900.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageBase900.Controls.Add(Me.Label22)
        Me.tabPageBase900.Location = New System.Drawing.Point(4, 27)
        Me.tabPageBase900.Name = "tabPageBase900"
        Me.tabPageBase900.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageBase900.Size = New System.Drawing.Size(847, 327)
        Me.tabPageBase900.TabIndex = 8
        Me.tabPageBase900.Text = " -"
        '
        'Label22
        '
        Me.Label22.AutoSize = true
        Me.Label22.Font = New System.Drawing.Font("メイリオ", 48!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Red
        Me.Label22.Location = New System.Drawing.Point(307, 92)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(232, 96)
        Me.Label22.TabIndex = 0
        Me.Label22.Text = "作成中"
        '
        'tabPageHanyo110
        '
        Me.tabPageHanyo110.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHanyo110.Controls.Add(Me.grpHMst)
        Me.tabPageHanyo110.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHanyo110.Name = "tabPageHanyo110"
        Me.tabPageHanyo110.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHanyo110.Size = New System.Drawing.Size(847, 327)
        Me.tabPageHanyo110.TabIndex = 12
        Me.tabPageHanyo110.Text = " 没選択1(汎用)"
        '
        'grpHMst
        '
        Me.grpHMst.Controls.Add(Me.Label138)
        Me.grpHMst.Controls.Add(Me.Label150)
        Me.grpHMst.Controls.Add(Me.Label151)
        Me.grpHMst.Controls.Add(Me.Label152)
        Me.grpHMst.Controls.Add(Me.Label153)
        Me.grpHMst.Controls.Add(Me.Label154)
        Me.grpHMst.Controls.Add(Me.Label155)
        Me.grpHMst.Controls.Add(Me.Label156)
        Me.grpHMst.Controls.Add(Me.Label157)
        Me.grpHMst.Controls.Add(Me.CheckBox2)
        Me.grpHMst.Controls.Add(Me.CheckBox3)
        Me.grpHMst.Controls.Add(Me.CheckBox6)
        Me.grpHMst.Controls.Add(Me.CheckBox9)
        Me.grpHMst.Controls.Add(Me.CheckBox10)
        Me.grpHMst.Controls.Add(Me.CheckBox11)
        Me.grpHMst.Controls.Add(Me.CheckBox12)
        Me.grpHMst.Controls.Add(Me.CheckBox13)
        Me.grpHMst.Controls.Add(Me.CheckBox17)
        Me.grpHMst.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpHMst.Location = New System.Drawing.Point(10, 5)
        Me.grpHMst.Name = "grpHMst"
        Me.grpHMst.Size = New System.Drawing.Size(814, 265)
        Me.grpHMst.TabIndex = 0
        Me.grpHMst.TabStop = false
        Me.grpHMst.Text = " 各マスタ情報 "
        '
        'Label138
        '
        Me.Label138.AutoSize = true
        Me.Label138.ForeColor = System.Drawing.Color.Black
        Me.Label138.Location = New System.Drawing.Point(222, 188)
        Me.Label138.Name = "Label138"
        Me.Label138.Size = New System.Drawing.Size(119, 18)
        Me.Label138.TabIndex = 17
        Me.Label138.Text = "   (例:オートロック)"
        '
        'Label150
        '
        Me.Label150.AutoSize = true
        Me.Label150.ForeColor = System.Drawing.Color.Black
        Me.Label150.Location = New System.Drawing.Point(222, 93)
        Me.Label150.Name = "Label150"
        Me.Label150.Size = New System.Drawing.Size(170, 18)
        Me.Label150.TabIndex = 13
        Me.Label150.Text = "   (例:建物(専用部)-クレーム)"
        '
        'Label151
        '
        Me.Label151.AutoSize = true
        Me.Label151.ForeColor = System.Drawing.Color.Black
        Me.Label151.Location = New System.Drawing.Point(222, 50)
        Me.Label151.Name = "Label151"
        Me.Label151.Size = New System.Drawing.Size(131, 18)
        Me.Label151.TabIndex = 11
        Me.Label151.Text = "   (例:第一種住宅地域)"
        '
        'Label152
        '
        Me.Label152.AutoSize = true
        Me.Label152.ForeColor = System.Drawing.Color.Black
        Me.Label152.Location = New System.Drawing.Point(222, 142)
        Me.Label152.Name = "Label152"
        Me.Label152.Size = New System.Drawing.Size(162, 18)
        Me.Label152.TabIndex = 15
        Me.Label152.Text = "   (口径や料金表などの設定)"
        '
        'Label153
        '
        Me.Label153.AutoSize = true
        Me.Label153.ForeColor = System.Drawing.Color.Black
        Me.Label153.Location = New System.Drawing.Point(22, 234)
        Me.Label153.Name = "Label153"
        Me.Label153.Size = New System.Drawing.Size(155, 18)
        Me.Label153.TabIndex = 9
        Me.Label153.Text = "   (例:定期借地借家権契約)"
        '
        'Label154
        '
        Me.Label154.AutoSize = true
        Me.Label154.ForeColor = System.Drawing.Color.Black
        Me.Label154.Location = New System.Drawing.Point(22, 188)
        Me.Label154.Name = "Label154"
        Me.Label154.Size = New System.Drawing.Size(119, 18)
        Me.Label154.TabIndex = 7
        Me.Label154.Text = "   (例:住宅総合保険)"
        '
        'Label155
        '
        Me.Label155.AutoSize = true
        Me.Label155.ForeColor = System.Drawing.Color.Black
        Me.Label155.Location = New System.Drawing.Point(22, 142)
        Me.Label155.Name = "Label155"
        Me.Label155.Size = New System.Drawing.Size(131, 18)
        Me.Label155.TabIndex = 5
        Me.Label155.Text = "   (例:東部、東エリア)"
        '
        'Label156
        '
        Me.Label156.AutoSize = true
        Me.Label156.ForeColor = System.Drawing.Color.Black
        Me.Label156.Location = New System.Drawing.Point(22, 96)
        Me.Label156.Name = "Label156"
        Me.Label156.Size = New System.Drawing.Size(155, 18)
        Me.Label156.TabIndex = 3
        Me.Label156.Text = "   (例:都城市立明道小学校)"
        '
        'Label157
        '
        Me.Label157.AutoSize = true
        Me.Label157.ForeColor = System.Drawing.Color.Black
        Me.Label157.Location = New System.Drawing.Point(22, 50)
        Me.Label157.Name = "Label157"
        Me.Label157.Size = New System.Drawing.Size(184, 18)
        Me.Label157.TabIndex = 1
        Me.Label157.Text = "   (例:宮崎交通-バスセンター前)"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = true
        Me.CheckBox2.Location = New System.Drawing.Point(225, 117)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(111, 22)
        Me.CheckBox2.TabIndex = 14
        Me.CheckBox2.Text = "変動費設定内容"
        Me.CheckBox2.UseVisualStyleBackColor = true
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = true
        Me.CheckBox3.Location = New System.Drawing.Point(225, 71)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(147, 22)
        Me.CheckBox3.TabIndex = 12
        Me.CheckBox3.Text = "クレーム分類設定内容"
        Me.CheckBox3.UseVisualStyleBackColor = true
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = true
        Me.CheckBox6.Location = New System.Drawing.Point(25, 25)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(111, 22)
        Me.CheckBox6.TabIndex = 0
        Me.CheckBox6.Text = "バス交通マスタ"
        Me.CheckBox6.UseVisualStyleBackColor = true
        '
        'CheckBox9
        '
        Me.CheckBox9.AutoSize = true
        Me.CheckBox9.Location = New System.Drawing.Point(25, 71)
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox9.TabIndex = 2
        Me.CheckBox9.Text = "学校区マスタ"
        Me.CheckBox9.UseVisualStyleBackColor = true
        '
        'CheckBox10
        '
        Me.CheckBox10.AutoSize = true
        Me.CheckBox10.Location = New System.Drawing.Point(225, 25)
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.Size = New System.Drawing.Size(87, 22)
        Me.CheckBox10.TabIndex = 10
        Me.CheckBox10.Text = "特約マスタ"
        Me.CheckBox10.UseVisualStyleBackColor = true
        '
        'CheckBox11
        '
        Me.CheckBox11.AutoSize = true
        Me.CheckBox11.Location = New System.Drawing.Point(25, 209)
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.Size = New System.Drawing.Size(111, 22)
        Me.CheckBox11.TabIndex = 8
        Me.CheckBox11.Text = "契約分類マスタ"
        Me.CheckBox11.UseVisualStyleBackColor = true
        '
        'CheckBox12
        '
        Me.CheckBox12.AutoSize = true
        Me.CheckBox12.Location = New System.Drawing.Point(25, 163)
        Me.CheckBox12.Name = "CheckBox12"
        Me.CheckBox12.Size = New System.Drawing.Size(111, 22)
        Me.CheckBox12.TabIndex = 6
        Me.CheckBox12.Text = "保険種類マスタ"
        Me.CheckBox12.UseVisualStyleBackColor = true
        '
        'CheckBox13
        '
        Me.CheckBox13.AutoSize = true
        Me.CheckBox13.Location = New System.Drawing.Point(225, 163)
        Me.CheckBox13.Name = "CheckBox13"
        Me.CheckBox13.Size = New System.Drawing.Size(123, 22)
        Me.CheckBox13.TabIndex = 16
        Me.CheckBox13.Text = "鍵タイトルマスタ"
        Me.CheckBox13.UseVisualStyleBackColor = true
        '
        'CheckBox17
        '
        Me.CheckBox17.AutoSize = true
        Me.CheckBox17.Location = New System.Drawing.Point(25, 117)
        Me.CheckBox17.Name = "CheckBox17"
        Me.CheckBox17.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox17.TabIndex = 4
        Me.CheckBox17.Text = "エリアマスタ"
        Me.CheckBox17.UseVisualStyleBackColor = true
        '
        'tabPageHanyo120
        '
        Me.tabPageHanyo120.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHanyo120.Controls.Add(Me.grpHGy)
        Me.tabPageHanyo120.Controls.Add(Me.grpHOw)
        Me.tabPageHanyo120.Controls.Add(Me.grpHJisya)
        Me.tabPageHanyo120.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHanyo120.Name = "tabPageHanyo120"
        Me.tabPageHanyo120.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHanyo120.Size = New System.Drawing.Size(847, 327)
        Me.tabPageHanyo120.TabIndex = 13
        Me.tabPageHanyo120.Text = " 没選択2(汎用)"
        '
        'grpHGy
        '
        Me.grpHGy.Controls.Add(Me.CheckBox4)
        Me.grpHGy.Controls.Add(Me.CheckBox30)
        Me.grpHGy.Controls.Add(Me.CheckBox31)
        Me.grpHGy.Controls.Add(Me.CheckBox32)
        Me.grpHGy.Controls.Add(Me.CheckBox33)
        Me.grpHGy.Controls.Add(Me.CheckBox34)
        Me.grpHGy.Controls.Add(Me.CheckBox35)
        Me.grpHGy.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpHGy.Location = New System.Drawing.Point(586, 6)
        Me.grpHGy.Name = "grpHGy"
        Me.grpHGy.Size = New System.Drawing.Size(245, 265)
        Me.grpHGy.TabIndex = 2
        Me.grpHGy.TabStop = false
        Me.grpHGy.Text = " 業者情報 "
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = true
        Me.CheckBox4.Location = New System.Drawing.Point(25, 192)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox4.TabIndex = 6
        Me.CheckBox4.Text = "施工業者情報"
        Me.CheckBox4.UseVisualStyleBackColor = true
        '
        'CheckBox30
        '
        Me.CheckBox30.AutoSize = true
        Me.CheckBox30.Location = New System.Drawing.Point(25, 164)
        Me.CheckBox30.Name = "CheckBox30"
        Me.CheckBox30.Size = New System.Drawing.Size(123, 22)
        Me.CheckBox30.TabIndex = 5
        Me.CheckBox30.Text = "施設保守業者情報"
        Me.CheckBox30.UseVisualStyleBackColor = true
        '
        'CheckBox31
        '
        Me.CheckBox31.AutoSize = true
        Me.CheckBox31.Location = New System.Drawing.Point(25, 52)
        Me.CheckBox31.Name = "CheckBox31"
        Me.CheckBox31.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox31.TabIndex = 1
        Me.CheckBox31.Text = "修繕業者情報"
        Me.CheckBox31.UseVisualStyleBackColor = true
        '
        'CheckBox32
        '
        Me.CheckBox32.AutoSize = true
        Me.CheckBox32.Location = New System.Drawing.Point(25, 136)
        Me.CheckBox32.Name = "CheckBox32"
        Me.CheckBox32.Size = New System.Drawing.Size(123, 22)
        Me.CheckBox32.TabIndex = 4
        Me.CheckBox32.Text = "家賃保証業者情報"
        Me.CheckBox32.UseVisualStyleBackColor = true
        '
        'CheckBox33
        '
        Me.CheckBox33.AutoSize = true
        Me.CheckBox33.Location = New System.Drawing.Point(25, 108)
        Me.CheckBox33.Name = "CheckBox33"
        Me.CheckBox33.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox33.TabIndex = 3
        Me.CheckBox33.Text = "保険業者情報"
        Me.CheckBox33.UseVisualStyleBackColor = true
        '
        'CheckBox34
        '
        Me.CheckBox34.AutoSize = true
        Me.CheckBox34.Location = New System.Drawing.Point(25, 80)
        Me.CheckBox34.Name = "CheckBox34"
        Me.CheckBox34.Size = New System.Drawing.Size(147, 22)
        Me.CheckBox34.TabIndex = 2
        Me.CheckBox34.Text = "ライフライン業者情報"
        Me.CheckBox34.UseVisualStyleBackColor = true
        '
        'CheckBox35
        '
        Me.CheckBox35.AutoSize = true
        Me.CheckBox35.Location = New System.Drawing.Point(25, 25)
        Me.CheckBox35.Name = "CheckBox35"
        Me.CheckBox35.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox35.TabIndex = 0
        Me.CheckBox35.Text = "仲介業者情報"
        Me.CheckBox35.UseVisualStyleBackColor = true
        '
        'grpHOw
        '
        Me.grpHOw.Controls.Add(Me.CheckBox1)
        Me.grpHOw.Controls.Add(Me.CheckBox26)
        Me.grpHOw.Controls.Add(Me.CheckBox27)
        Me.grpHOw.Controls.Add(Me.CheckBox28)
        Me.grpHOw.Controls.Add(Me.CheckBox29)
        Me.grpHOw.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpHOw.Location = New System.Drawing.Point(269, 6)
        Me.grpHOw.Name = "grpHOw"
        Me.grpHOw.Size = New System.Drawing.Size(311, 265)
        Me.grpHOw.TabIndex = 1
        Me.grpHOw.TabStop = false
        Me.grpHOw.Text = "家主情報"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = true
        Me.CheckBox1.Location = New System.Drawing.Point(45, 136)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(217, 22)
        Me.CheckBox1.TabIndex = 4
        Me.CheckBox1.Text = "(未使用の家主データは移行しない)"
        Me.CheckBox1.UseVisualStyleBackColor = true
        '
        'CheckBox26
        '
        Me.CheckBox26.AutoSize = true
        Me.CheckBox26.Location = New System.Drawing.Point(45, 80)
        Me.CheckBox26.Name = "CheckBox26"
        Me.CheckBox26.Size = New System.Drawing.Size(123, 22)
        Me.CheckBox26.TabIndex = 2
        Me.CheckBox26.Text = "家主イベント情報"
        Me.CheckBox26.UseVisualStyleBackColor = true
        '
        'CheckBox27
        '
        Me.CheckBox27.AutoSize = true
        Me.CheckBox27.Location = New System.Drawing.Point(45, 53)
        Me.CheckBox27.Name = "CheckBox27"
        Me.CheckBox27.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox27.TabIndex = 1
        Me.CheckBox27.Text = "家主口座情報"
        Me.CheckBox27.UseVisualStyleBackColor = true
        '
        'CheckBox28
        '
        Me.CheckBox28.AutoSize = true
        Me.CheckBox28.Location = New System.Drawing.Point(45, 108)
        Me.CheckBox28.Name = "CheckBox28"
        Me.CheckBox28.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox28.TabIndex = 3
        Me.CheckBox28.Text = "家主メモ情報"
        Me.CheckBox28.UseVisualStyleBackColor = true
        '
        'CheckBox29
        '
        Me.CheckBox29.AutoSize = true
        Me.CheckBox29.Location = New System.Drawing.Point(25, 25)
        Me.CheckBox29.Name = "CheckBox29"
        Me.CheckBox29.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox29.TabIndex = 0
        Me.CheckBox29.Text = "家主基本情報"
        Me.CheckBox29.UseVisualStyleBackColor = true
        '
        'grpHJisya
        '
        Me.grpHJisya.Controls.Add(Me.CheckBox19)
        Me.grpHJisya.Controls.Add(Me.CheckBox20)
        Me.grpHJisya.Controls.Add(Me.CheckBox21)
        Me.grpHJisya.Controls.Add(Me.CheckBox22)
        Me.grpHJisya.Controls.Add(Me.CheckBox25)
        Me.grpHJisya.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.grpHJisya.Location = New System.Drawing.Point(10, 5)
        Me.grpHJisya.Name = "grpHJisya"
        Me.grpHJisya.Size = New System.Drawing.Size(253, 265)
        Me.grpHJisya.TabIndex = 0
        Me.grpHJisya.TabStop = false
        Me.grpHJisya.Text = " 自社情報 "
        '
        'CheckBox19
        '
        Me.CheckBox19.AutoSize = true
        Me.CheckBox19.Location = New System.Drawing.Point(45, 109)
        Me.CheckBox19.Name = "CheckBox19"
        Me.CheckBox19.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox19.TabIndex = 3
        Me.CheckBox19.Text = "口座振替情報"
        Me.CheckBox19.UseVisualStyleBackColor = true
        '
        'CheckBox20
        '
        Me.CheckBox20.AutoSize = true
        Me.CheckBox20.Location = New System.Drawing.Point(44, 81)
        Me.CheckBox20.Name = "CheckBox20"
        Me.CheckBox20.Size = New System.Drawing.Size(111, 22)
        Me.CheckBox20.TabIndex = 2
        Me.CheckBox20.Text = "振込依頼人情報"
        Me.CheckBox20.UseVisualStyleBackColor = true
        '
        'CheckBox21
        '
        Me.CheckBox21.AutoSize = true
        Me.CheckBox21.Location = New System.Drawing.Point(45, 53)
        Me.CheckBox21.Name = "CheckBox21"
        Me.CheckBox21.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox21.TabIndex = 1
        Me.CheckBox21.Text = "自社口座情報"
        Me.CheckBox21.UseVisualStyleBackColor = true
        '
        'CheckBox22
        '
        Me.CheckBox22.AutoSize = true
        Me.CheckBox22.Location = New System.Drawing.Point(44, 137)
        Me.CheckBox22.Name = "CheckBox22"
        Me.CheckBox22.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox22.TabIndex = 4
        Me.CheckBox22.Text = "自社メモ情報"
        Me.CheckBox22.UseVisualStyleBackColor = true
        '
        'CheckBox25
        '
        Me.CheckBox25.AutoSize = true
        Me.CheckBox25.Location = New System.Drawing.Point(25, 25)
        Me.CheckBox25.Name = "CheckBox25"
        Me.CheckBox25.Size = New System.Drawing.Size(99, 22)
        Me.CheckBox25.TabIndex = 0
        Me.CheckBox25.Text = "自社基本情報"
        Me.CheckBox25.UseVisualStyleBackColor = true
        '
        'tabPageHanyo130
        '
        Me.tabPageHanyo130.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHanyo130.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHanyo130.Name = "tabPageHanyo130"
        Me.tabPageHanyo130.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHanyo130.Size = New System.Drawing.Size(847, 327)
        Me.tabPageHanyo130.TabIndex = 17
        Me.tabPageHanyo130.Text = " 没選択3(汎用)"
        '
        'tabPageHanyo140
        '
        Me.tabPageHanyo140.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHanyo140.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHanyo140.Name = "tabPageHanyo140"
        Me.tabPageHanyo140.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHanyo140.Size = New System.Drawing.Size(847, 327)
        Me.tabPageHanyo140.TabIndex = 16
        Me.tabPageHanyo140.Text = " 没選択4(汎用)"
        '
        'tabPageHanyo150
        '
        Me.tabPageHanyo150.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHanyo150.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHanyo150.Name = "tabPageHanyo150"
        Me.tabPageHanyo150.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHanyo150.Size = New System.Drawing.Size(847, 327)
        Me.tabPageHanyo150.TabIndex = 15
        Me.tabPageHanyo150.Text = " 没選択5(汎用)"
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Menu
        Me.TabPage1.Controls.Add(Me.grpKizon4)
        Me.TabPage1.Location = New System.Drawing.Point(4, 27)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(847, 327)
        Me.TabPage1.TabIndex = 20
        Me.TabPage1.Text = "TabPage1"
        '
        'grpKizon4
        '
        Me.grpKizon4.Controls.Add(Me.lblSelectPageCnt3)
        Me.grpKizon4.Controls.Add(Me.Label26)
        Me.grpKizon4.Controls.Add(Me.Label73)
        Me.grpKizon4.Controls.Add(Me.Label70)
        Me.grpKizon4.Controls.Add(Me.Label71)
        Me.grpKizon4.Controls.Add(Me.Label72)
        Me.grpKizon4.Controls.Add(Me.Label74)
        Me.grpKizon4.Controls.Add(Me.chkHSongai)
        Me.grpKizon4.Controls.Add(Me.chkHJisyaKoza)
        Me.grpKizon4.Controls.Add(Me.chkHSetubi)
        Me.grpKizon4.Controls.Add(Me.chkHDataFmt)
        Me.grpKizon4.Controls.Add(Me.Label69)
        Me.grpKizon4.Controls.Add(Me.Label68)
        Me.grpKizon4.Controls.Add(Me.Label67)
        Me.grpKizon4.Controls.Add(Me.Label66)
        Me.grpKizon4.Controls.Add(Me.Label65)
        Me.grpKizon4.Controls.Add(Me.Label64)
        Me.grpKizon4.Controls.Add(Me.Label63)
        Me.grpKizon4.Controls.Add(Me.Label62)
        Me.grpKizon4.Controls.Add(Me.Label61)
        Me.grpKizon4.Controls.Add(Me.Label60)
        Me.grpKizon4.Controls.Add(Me.Label59)
        Me.grpKizon4.Controls.Add(Me.chkHNkinKomk)
        Me.grpKizon4.Controls.Add(Me.chkHTaiyo)
        Me.grpKizon4.Controls.Add(Me.chkHKozo)
        Me.grpKizon4.Controls.Add(Me.chkHKozaSyu)
        Me.grpKizon4.Controls.Add(Me.chkHEki)
        Me.grpKizon4.Controls.Add(Me.chkHEnsen)
        Me.grpKizon4.Controls.Add(Me.chkHKinyuSiten)
        Me.grpKizon4.Controls.Add(Me.chkHYouto)
        Me.grpKizon4.Controls.Add(Me.chkHBkBunrui)
        Me.grpKizon4.Controls.Add(Me.chkHKinyu)
        Me.grpKizon4.Controls.Add(Me.chkHKyBunrui)
        Me.grpKizon4.Controls.Add(Me.chkHKagi)
        Me.grpKizon4.Controls.Add(Me.chkHHouKenri)
        Me.grpKizon4.Controls.Add(Me.chkHNkinKbn)
        Me.grpKizon4.Controls.Add(Me.chkHHyBunrui)
        Me.grpKizon4.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.grpKizon4.Location = New System.Drawing.Point(6, 7)
        Me.grpKizon4.Name = "grpKizon4"
        Me.grpKizon4.Size = New System.Drawing.Size(835, 265)
        Me.grpKizon4.TabIndex = 1
        Me.grpKizon4.TabStop = false
        Me.grpKizon4.Text = " 紐付設定項目 "
        '
        'lblSelectPageCnt3
        '
        Me.lblSelectPageCnt3.AutoSize = true
        Me.lblSelectPageCnt3.Location = New System.Drawing.Point(794, 242)
        Me.lblSelectPageCnt3.Name = "lblSelectPageCnt3"
        Me.lblSelectPageCnt3.Size = New System.Drawing.Size(31, 18)
        Me.lblSelectPageCnt3.TabIndex = 36
        Me.lblSelectPageCnt3.Text = "- / -"
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label26.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Red
        Me.Label26.Location = New System.Drawing.Point(82, 80)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(610, 159)
        Me.Label26.TabIndex = 35
        Me.Label26.Text = "※没"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"ここで用意している紐付設定については、前画面の関連する親情報(物件情報や部屋情報など)にチェックがある場合に自動呼出ししています。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(自動で呼び出すよう"& _ 
    "にしている為、ここのチェックは現時点で機能していません)"
        '
        'Label73
        '
        Me.Label73.BackColor = System.Drawing.SystemColors.Menu
        Me.Label73.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label73.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label73.Location = New System.Drawing.Point(11, 20)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(795, 40)
        Me.Label73.TabIndex = 0
        Me.Label73.Text = "紐付を行う必要がある項目に自動でチェックが入ります。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"(例えば、物件情報の項目にチェックを入れた場合、それに関連する項目(物件用途など)に自動でチェックが入りま"& _ 
    "す)"
        Me.Label73.UseCompatibleTextRendering = true
        '
        'Label70
        '
        Me.Label70.AutoSize = true
        Me.Label70.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label70.ForeColor = System.Drawing.Color.Black
        Me.Label70.Location = New System.Drawing.Point(347, 234)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(129, 17)
        Me.Label70.TabIndex = 24
        Me.Label70.Text = "   (該当入金項目の選定)"
        '
        'Label71
        '
        Me.Label71.AutoSize = true
        Me.Label71.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label71.ForeColor = System.Drawing.Color.Black
        Me.Label71.Location = New System.Drawing.Point(507, 188)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(140, 17)
        Me.Label71.TabIndex = 30
        Me.Label71.Text = "   (全銀フォーマット選択)"
        '
        'Label72
        '
        Me.Label72.AutoSize = true
        Me.Label72.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label72.ForeColor = System.Drawing.Color.Black
        Me.Label72.Location = New System.Drawing.Point(507, 142)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(107, 17)
        Me.Label72.TabIndex = 28
        Me.Label72.Text = "   (自社口座の選択)"
        '
        'Label74
        '
        Me.Label74.AutoSize = true
        Me.Label74.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label74.ForeColor = System.Drawing.Color.Black
        Me.Label74.Location = New System.Drawing.Point(182, 142)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(122, 17)
        Me.Label74.TabIndex = 12
        Me.Label74.Text = "   (例:収納-物置-有り)"
        '
        'chkHSongai
        '
        Me.chkHSongai.AutoSize = true
        Me.chkHSongai.ForeColor = System.Drawing.Color.Black
        Me.chkHSongai.Location = New System.Drawing.Point(350, 209)
        Me.chkHSongai.Name = "chkHSongai"
        Me.chkHSongai.Size = New System.Drawing.Size(87, 22)
        Me.chkHSongai.TabIndex = 23
        Me.chkHSongai.Text = "損害保険料"
        Me.chkHSongai.UseVisualStyleBackColor = true
        '
        'chkHJisyaKoza
        '
        Me.chkHJisyaKoza.AutoSize = true
        Me.chkHJisyaKoza.ForeColor = System.Drawing.Color.Black
        Me.chkHJisyaKoza.Location = New System.Drawing.Point(510, 117)
        Me.chkHJisyaKoza.Name = "chkHJisyaKoza"
        Me.chkHJisyaKoza.Size = New System.Drawing.Size(147, 22)
        Me.chkHJisyaKoza.TabIndex = 27
        Me.chkHJisyaKoza.Text = "振込・振替・家賃口座"
        Me.chkHJisyaKoza.UseVisualStyleBackColor = true
        '
        'chkHSetubi
        '
        Me.chkHSetubi.AutoSize = true
        Me.chkHSetubi.ForeColor = System.Drawing.Color.Black
        Me.chkHSetubi.Location = New System.Drawing.Point(185, 117)
        Me.chkHSetubi.Name = "chkHSetubi"
        Me.chkHSetubi.Size = New System.Drawing.Size(75, 22)
        Me.chkHSetubi.TabIndex = 11
        Me.chkHSetubi.Text = "部屋設備"
        Me.chkHSetubi.UseVisualStyleBackColor = true
        '
        'chkHDataFmt
        '
        Me.chkHDataFmt.AutoSize = true
        Me.chkHDataFmt.ForeColor = System.Drawing.Color.Black
        Me.chkHDataFmt.Location = New System.Drawing.Point(510, 163)
        Me.chkHDataFmt.Name = "chkHDataFmt"
        Me.chkHDataFmt.Size = New System.Drawing.Size(135, 22)
        Me.chkHDataFmt.TabIndex = 29
        Me.chkHDataFmt.Text = "データフォーマット"
        Me.chkHDataFmt.UseVisualStyleBackColor = true
        '
        'Label69
        '
        Me.Label69.AutoSize = true
        Me.Label69.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label69.ForeColor = System.Drawing.Color.Black
        Me.Label69.Location = New System.Drawing.Point(347, 188)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(111, 17)
        Me.Label69.TabIndex = 22
        Me.Label69.Text = "   (例:賃料(毎月時))"
        '
        'Label68
        '
        Me.Label68.AutoSize = true
        Me.Label68.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label68.ForeColor = System.Drawing.Color.Black
        Me.Label68.Location = New System.Drawing.Point(347, 142)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(68, 17)
        Me.Label68.TabIndex = 20
        Me.Label68.Text = "   (例:振替)"
        '
        'Label67
        '
        Me.Label67.AutoSize = true
        Me.Label67.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label67.ForeColor = System.Drawing.Color.Black
        Me.Label67.Location = New System.Drawing.Point(182, 234)
        Me.Label67.Name = "Label67"
        Me.Label67.Size = New System.Drawing.Size(68, 17)
        Me.Label67.TabIndex = 16
        Me.Label67.Text = "   (例:仲介)"
        '
        'Label66
        '
        Me.Label66.AutoSize = true
        Me.Label66.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label66.ForeColor = System.Drawing.Color.Black
        Me.Label66.Location = New System.Drawing.Point(347, 96)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(90, 17)
        Me.Label66.TabIndex = 18
        Me.Label66.Text = "   (例:普通預金)"
        '
        'Label65
        '
        Me.Label65.AutoSize = true
        Me.Label65.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label65.ForeColor = System.Drawing.Color.Black
        Me.Label65.Location = New System.Drawing.Point(182, 188)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(145, 17)
        Me.Label65.TabIndex = 14
        Me.Label65.Text = "   (例:普通建物賃貸借契約)"
        '
        'Label64
        '
        Me.Label64.AutoSize = true
        Me.Label64.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label64.ForeColor = System.Drawing.Color.Black
        Me.Label64.Location = New System.Drawing.Point(22, 142)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(145, 17)
        Me.Label64.TabIndex = 4
        Me.Label64.Text = "   (例:鉄筋コンクリート造)"
        '
        'Label63
        '
        Me.Label63.AutoSize = true
        Me.Label63.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label63.ForeColor = System.Drawing.Color.Black
        Me.Label63.Location = New System.Drawing.Point(507, 96)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(96, 17)
        Me.Label63.TabIndex = 26
        Me.Label63.Text = "   (共用鍵の設定)"
        '
        'Label62
        '
        Me.Label62.AutoSize = true
        Me.Label62.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label62.ForeColor = System.Drawing.Color.Black
        Me.Label62.Location = New System.Drawing.Point(182, 96)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(133, 17)
        Me.Label62.TabIndex = 10
        Me.Label62.Text = "   (例:アパート(事業用))"
        '
        'Label61
        '
        Me.Label61.AutoSize = true
        Me.Label61.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label61.ForeColor = System.Drawing.Color.Black
        Me.Label61.Location = New System.Drawing.Point(22, 234)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(90, 17)
        Me.Label61.TabIndex = 8
        Me.Label61.Text = "   (例:アパート)"
        '
        'Label60
        '
        Me.Label60.AutoSize = true
        Me.Label60.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label60.ForeColor = System.Drawing.Color.Black
        Me.Label60.Location = New System.Drawing.Point(22, 188)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(140, 17)
        Me.Label60.TabIndex = 6
        Me.Label60.Text = "   (例：土砂災害警戒区域)"
        '
        'Label59
        '
        Me.Label59.AutoSize = true
        Me.Label59.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label59.ForeColor = System.Drawing.Color.Black
        Me.Label59.Location = New System.Drawing.Point(22, 96)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(123, 17)
        Me.Label59.TabIndex = 2
        Me.Label59.Text = "   (例:第一種住宅地域)"
        '
        'chkHNkinKomk
        '
        Me.chkHNkinKomk.AutoSize = true
        Me.chkHNkinKomk.ForeColor = System.Drawing.Color.Black
        Me.chkHNkinKomk.Location = New System.Drawing.Point(350, 163)
        Me.chkHNkinKomk.Name = "chkHNkinKomk"
        Me.chkHNkinKomk.Size = New System.Drawing.Size(75, 22)
        Me.chkHNkinKomk.TabIndex = 21
        Me.chkHNkinKomk.Text = "入金項目"
        Me.chkHNkinKomk.UseVisualStyleBackColor = true
        '
        'chkHTaiyo
        '
        Me.chkHTaiyo.AutoSize = true
        Me.chkHTaiyo.ForeColor = System.Drawing.Color.Black
        Me.chkHTaiyo.Location = New System.Drawing.Point(185, 209)
        Me.chkHTaiyo.Name = "chkHTaiyo"
        Me.chkHTaiyo.Size = New System.Drawing.Size(75, 22)
        Me.chkHTaiyo.TabIndex = 15
        Me.chkHTaiyo.Text = "取引態様"
        Me.chkHTaiyo.UseVisualStyleBackColor = true
        '
        'chkHKozo
        '
        Me.chkHKozo.AutoSize = true
        Me.chkHKozo.ForeColor = System.Drawing.Color.Black
        Me.chkHKozo.Location = New System.Drawing.Point(25, 117)
        Me.chkHKozo.Name = "chkHKozo"
        Me.chkHKozo.Size = New System.Drawing.Size(75, 22)
        Me.chkHKozo.TabIndex = 3
        Me.chkHKozo.Text = "物件構造"
        Me.chkHKozo.UseVisualStyleBackColor = true
        '
        'chkHKozaSyu
        '
        Me.chkHKozaSyu.AutoSize = true
        Me.chkHKozaSyu.ForeColor = System.Drawing.Color.Black
        Me.chkHKozaSyu.Location = New System.Drawing.Point(350, 70)
        Me.chkHKozaSyu.Name = "chkHKozaSyu"
        Me.chkHKozaSyu.Size = New System.Drawing.Size(75, 22)
        Me.chkHKozaSyu.TabIndex = 17
        Me.chkHKozaSyu.Text = "口座種別"
        Me.chkHKozaSyu.UseVisualStyleBackColor = true
        '
        'chkHEki
        '
        Me.chkHEki.AutoSize = true
        Me.chkHEki.ForeColor = System.Drawing.Color.Black
        Me.chkHEki.Location = New System.Drawing.Point(670, 209)
        Me.chkHEki.Name = "chkHEki"
        Me.chkHEki.Size = New System.Drawing.Size(39, 22)
        Me.chkHEki.TabIndex = 34
        Me.chkHEki.Text = "駅"
        Me.chkHEki.UseVisualStyleBackColor = true
        '
        'chkHEnsen
        '
        Me.chkHEnsen.AutoSize = true
        Me.chkHEnsen.ForeColor = System.Drawing.Color.Black
        Me.chkHEnsen.Location = New System.Drawing.Point(670, 163)
        Me.chkHEnsen.Name = "chkHEnsen"
        Me.chkHEnsen.Size = New System.Drawing.Size(51, 22)
        Me.chkHEnsen.TabIndex = 33
        Me.chkHEnsen.Text = "沿線"
        Me.chkHEnsen.UseVisualStyleBackColor = true
        '
        'chkHKinyuSiten
        '
        Me.chkHKinyuSiten.AutoSize = true
        Me.chkHKinyuSiten.ForeColor = System.Drawing.Color.Black
        Me.chkHKinyuSiten.Location = New System.Drawing.Point(670, 117)
        Me.chkHKinyuSiten.Name = "chkHKinyuSiten"
        Me.chkHKinyuSiten.Size = New System.Drawing.Size(111, 22)
        Me.chkHKinyuSiten.TabIndex = 32
        Me.chkHKinyuSiten.Text = "金融機関本支店"
        Me.chkHKinyuSiten.UseVisualStyleBackColor = true
        '
        'chkHYouto
        '
        Me.chkHYouto.AutoSize = true
        Me.chkHYouto.ForeColor = System.Drawing.Color.Black
        Me.chkHYouto.Location = New System.Drawing.Point(25, 70)
        Me.chkHYouto.Name = "chkHYouto"
        Me.chkHYouto.Size = New System.Drawing.Size(75, 22)
        Me.chkHYouto.TabIndex = 1
        Me.chkHYouto.Text = "物件用途"
        Me.chkHYouto.UseVisualStyleBackColor = true
        '
        'chkHBkBunrui
        '
        Me.chkHBkBunrui.AutoSize = true
        Me.chkHBkBunrui.ForeColor = System.Drawing.Color.Black
        Me.chkHBkBunrui.Location = New System.Drawing.Point(25, 209)
        Me.chkHBkBunrui.Name = "chkHBkBunrui"
        Me.chkHBkBunrui.Size = New System.Drawing.Size(75, 22)
        Me.chkHBkBunrui.TabIndex = 7
        Me.chkHBkBunrui.Text = "物件分類"
        Me.chkHBkBunrui.UseVisualStyleBackColor = true
        '
        'chkHKinyu
        '
        Me.chkHKinyu.AutoSize = true
        Me.chkHKinyu.ForeColor = System.Drawing.Color.Black
        Me.chkHKinyu.Location = New System.Drawing.Point(665, 71)
        Me.chkHKinyu.Name = "chkHKinyu"
        Me.chkHKinyu.Size = New System.Drawing.Size(75, 22)
        Me.chkHKinyu.TabIndex = 31
        Me.chkHKinyu.Text = "金融機関"
        Me.chkHKinyu.UseVisualStyleBackColor = true
        '
        'chkHKyBunrui
        '
        Me.chkHKyBunrui.AutoSize = true
        Me.chkHKyBunrui.ForeColor = System.Drawing.Color.Black
        Me.chkHKyBunrui.Location = New System.Drawing.Point(185, 163)
        Me.chkHKyBunrui.Name = "chkHKyBunrui"
        Me.chkHKyBunrui.Size = New System.Drawing.Size(75, 22)
        Me.chkHKyBunrui.TabIndex = 13
        Me.chkHKyBunrui.Text = "契約分類"
        Me.chkHKyBunrui.UseVisualStyleBackColor = true
        '
        'chkHKagi
        '
        Me.chkHKagi.AutoSize = true
        Me.chkHKagi.ForeColor = System.Drawing.Color.Black
        Me.chkHKagi.Location = New System.Drawing.Point(510, 70)
        Me.chkHKagi.Name = "chkHKagi"
        Me.chkHKagi.Size = New System.Drawing.Size(63, 22)
        Me.chkHKagi.TabIndex = 25
        Me.chkHKagi.Text = "鍵情報"
        Me.chkHKagi.UseVisualStyleBackColor = true
        '
        'chkHHouKenri
        '
        Me.chkHHouKenri.AutoSize = true
        Me.chkHHouKenri.ForeColor = System.Drawing.Color.Black
        Me.chkHHouKenri.Location = New System.Drawing.Point(25, 163)
        Me.chkHHouKenri.Name = "chkHHouKenri"
        Me.chkHHouKenri.Size = New System.Drawing.Size(111, 22)
        Me.chkHHouKenri.TabIndex = 5
        Me.chkHHouKenri.Text = "物件法令・権利"
        Me.chkHHouKenri.UseVisualStyleBackColor = true
        '
        'chkHNkinKbn
        '
        Me.chkHNkinKbn.AutoSize = true
        Me.chkHNkinKbn.ForeColor = System.Drawing.Color.Black
        Me.chkHNkinKbn.Location = New System.Drawing.Point(350, 117)
        Me.chkHNkinKbn.Name = "chkHNkinKbn"
        Me.chkHNkinKbn.Size = New System.Drawing.Size(75, 22)
        Me.chkHNkinKbn.TabIndex = 19
        Me.chkHNkinKbn.Text = "入金区分"
        Me.chkHNkinKbn.UseVisualStyleBackColor = true
        '
        'chkHHyBunrui
        '
        Me.chkHHyBunrui.AutoSize = true
        Me.chkHHyBunrui.ForeColor = System.Drawing.Color.Black
        Me.chkHHyBunrui.Location = New System.Drawing.Point(185, 71)
        Me.chkHHyBunrui.Name = "chkHHyBunrui"
        Me.chkHHyBunrui.Size = New System.Drawing.Size(75, 22)
        Me.chkHHyBunrui.TabIndex = 9
        Me.chkHHyBunrui.Text = "部屋分類"
        Me.chkHHyBunrui.UseVisualStyleBackColor = true
        '
        'pnlRekiClear
        '
        Me.pnlRekiClear.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlRekiClear.Controls.Add(Me.btnRekiClear)
        Me.pnlRekiClear.Controls.Add(Me.Label79)
        Me.pnlRekiClear.Controls.Add(Me.Label298)
        Me.pnlRekiClear.Location = New System.Drawing.Point(563, 36)
        Me.pnlRekiClear.Name = "pnlRekiClear"
        Me.pnlRekiClear.Size = New System.Drawing.Size(326, 45)
        Me.pnlRekiClear.TabIndex = 45
        '
        'btnRekiClear
        '
        Me.btnRekiClear.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnRekiClear.Image = CType(resources.GetObject("btnRekiClear.Image"),System.Drawing.Image)
        Me.btnRekiClear.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRekiClear.Location = New System.Drawing.Point(229, 5)
        Me.btnRekiClear.Name = "btnRekiClear"
        Me.btnRekiClear.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnRekiClear.Size = New System.Drawing.Size(88, 30)
        Me.btnRekiClear.TabIndex = 2
        Me.btnRekiClear.Text = " クリア"
        Me.btnRekiClear.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnRekiClear.UseVisualStyleBackColor = true
        '
        'Label79
        '
        Me.Label79.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label79.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label79.Location = New System.Drawing.Point(18, 2)
        Me.Label79.Name = "Label79"
        Me.Label79.Size = New System.Drawing.Size(32, 18)
        Me.Label79.TabIndex = 1
        Me.Label79.Text = "緑色"
        '
        'Label298
        '
        Me.Label298.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label298.ForeColor = System.Drawing.Color.Black
        Me.Label298.Location = New System.Drawing.Point(3, 3)
        Me.Label298.Name = "Label298"
        Me.Label298.Size = New System.Drawing.Size(228, 38)
        Me.Label298.TabIndex = 0
        Me.Label298.Text = "※　　　に着色されているチェック"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　ボックスはコンバート実行済み項目です。"
        '
        'lblDatacvSelectDescription1
        '
        Me.lblDatacvSelectDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblDatacvSelectDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblDatacvSelectDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblDatacvSelectDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblDatacvSelectDescription1.Name = "lblDatacvSelectDescription1"
        Me.lblDatacvSelectDescription1.Size = New System.Drawing.Size(876, 80)
        Me.lblDatacvSelectDescription1.TabIndex = 0
        Me.lblDatacvSelectDescription1.Text = "コンバート対象項目の選択を行います。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"コンバートを行う項目にチェックを入れて、[次へ] ボタンを押して下さい。"
        Me.lblDatacvSelectDescription1.UseCompatibleTextRendering = true
        '
        'tabPageJikko
        '
        Me.tabPageJikko.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageJikko.Controls.Add(Me.lblDatacvJikkoCaution)
        Me.tabPageJikko.Controls.Add(Me.lblDatacvJikkoDescription1)
        Me.tabPageJikko.Controls.Add(Me.grpTotalProcess)
        Me.tabPageJikko.Location = New System.Drawing.Point(4, 27)
        Me.tabPageJikko.Name = "tabPageJikko"
        Me.tabPageJikko.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageJikko.Size = New System.Drawing.Size(912, 549)
        Me.tabPageJikko.TabIndex = 3
        Me.tabPageJikko.Text = "移行処理"
        '
        'lblDatacvJikkoCaution
        '
        Me.lblDatacvJikkoCaution.BackColor = System.Drawing.SystemColors.Menu
        Me.lblDatacvJikkoCaution.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblDatacvJikkoCaution.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblDatacvJikkoCaution.Location = New System.Drawing.Point(30, 45)
        Me.lblDatacvJikkoCaution.Name = "lblDatacvJikkoCaution"
        Me.lblDatacvJikkoCaution.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDatacvJikkoCaution.Size = New System.Drawing.Size(851, 60)
        Me.lblDatacvJikkoCaution.TabIndex = 119
        Me.lblDatacvJikkoCaution.Text = "※コンバート処理中は、賃貸革命を使用しないで下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※「キャンセル」ボタンで処理を中止します (途中再開(リジューム)はできません)。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※中間ファイル書込処"& _ 
    "理完了後に紐付設定画面が表示されます。項目の紐付設定を行って下さい。"
        Me.lblDatacvJikkoCaution.UseCompatibleTextRendering = true
        '
        'lblDatacvJikkoDescription1
        '
        Me.lblDatacvJikkoDescription1.BackColor = System.Drawing.SystemColors.Menu
        Me.lblDatacvJikkoDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblDatacvJikkoDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblDatacvJikkoDescription1.Location = New System.Drawing.Point(30, 20)
        Me.lblDatacvJikkoDescription1.Name = "lblDatacvJikkoDescription1"
        Me.lblDatacvJikkoDescription1.Size = New System.Drawing.Size(851, 80)
        Me.lblDatacvJikkoDescription1.TabIndex = 115
        Me.lblDatacvJikkoDescription1.Text = "データコンバート処理を行います。処理が完了したら、完了ダイアログが表示されます。表示後、「次へ」ボタンを押して下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)
        Me.lblDatacvJikkoDescription1.UseCompatibleTextRendering = true
        '
        'grpTotalProcess
        '
        Me.grpTotalProcess.Controls.Add(Me.lblCVItem)
        Me.grpTotalProcess.Controls.Add(Me.txtPartialSituation)
        Me.grpTotalProcess.Controls.Add(Me.Label31)
        Me.grpTotalProcess.Controls.Add(Me.lblPgbPartial)
        Me.grpTotalProcess.Controls.Add(Me.lblPgbTotal)
        Me.grpTotalProcess.Controls.Add(Me.txtTotalSituation)
        Me.grpTotalProcess.Controls.Add(Me.pgbpartial)
        Me.grpTotalProcess.Controls.Add(Me.pgbTotal)
        Me.grpTotalProcess.Controls.Add(Me.Label32)
        Me.grpTotalProcess.Controls.Add(Me.lblTotalSituation)
        Me.grpTotalProcess.Location = New System.Drawing.Point(22, 121)
        Me.grpTotalProcess.Name = "grpTotalProcess"
        Me.grpTotalProcess.Size = New System.Drawing.Size(859, 408)
        Me.grpTotalProcess.TabIndex = 113
        Me.grpTotalProcess.TabStop = false
        Me.grpTotalProcess.Text = "【進捗状況】"
        '
        'lblCVItem
        '
        Me.lblCVItem.Location = New System.Drawing.Point(16, 87)
        Me.lblCVItem.Name = "lblCVItem"
        Me.lblCVItem.Size = New System.Drawing.Size(336, 16)
        Me.lblCVItem.TabIndex = 114
        Me.lblCVItem.Text = "処理項目"
        Me.lblCVItem.UseCompatibleTextRendering = true
        '
        'txtPartialSituation
        '
        Me.txtPartialSituation.Location = New System.Drawing.Point(452, 202)
        Me.txtPartialSituation.Multiline = true
        Me.txtPartialSituation.Name = "txtPartialSituation"
        Me.txtPartialSituation.Size = New System.Drawing.Size(365, 190)
        Me.txtPartialSituation.TabIndex = 116
        '
        'Label31
        '
        Me.Label31.Location = New System.Drawing.Point(452, 169)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(80, 16)
        Me.Label31.TabIndex = 113
        Me.Label31.Text = "処理完了項目"
        Me.Label31.UseCompatibleTextRendering = true
        '
        'lblPgbPartial
        '
        Me.lblPgbPartial.Location = New System.Drawing.Point(796, 107)
        Me.lblPgbPartial.Name = "lblPgbPartial"
        Me.lblPgbPartial.Size = New System.Drawing.Size(52, 20)
        Me.lblPgbPartial.TabIndex = 117
        Me.lblPgbPartial.Text = "0 %"
        Me.lblPgbPartial.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblPgbPartial.UseCompatibleTextRendering = true
        '
        'lblPgbTotal
        '
        Me.lblPgbTotal.Location = New System.Drawing.Point(797, 50)
        Me.lblPgbTotal.Name = "lblPgbTotal"
        Me.lblPgbTotal.Size = New System.Drawing.Size(51, 20)
        Me.lblPgbTotal.TabIndex = 115
        Me.lblPgbTotal.Text = "0 %"
        Me.lblPgbTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblPgbTotal.UseCompatibleTextRendering = true
        '
        'txtTotalSituation
        '
        Me.txtTotalSituation.Location = New System.Drawing.Point(40, 202)
        Me.txtTotalSituation.Multiline = true
        Me.txtTotalSituation.Name = "txtTotalSituation"
        Me.txtTotalSituation.Size = New System.Drawing.Size(365, 190)
        Me.txtTotalSituation.TabIndex = 114
        '
        'pgbpartial
        '
        Me.pgbpartial.Location = New System.Drawing.Point(40, 107)
        Me.pgbpartial.Name = "pgbpartial"
        Me.pgbpartial.Size = New System.Drawing.Size(750, 20)
        Me.pgbpartial.TabIndex = 115
        '
        'pgbTotal
        '
        Me.pgbTotal.Location = New System.Drawing.Point(40, 50)
        Me.pgbTotal.Name = "pgbTotal"
        Me.pgbTotal.Size = New System.Drawing.Size(751, 20)
        Me.pgbTotal.TabIndex = 113
        '
        'Label32
        '
        Me.Label32.Location = New System.Drawing.Point(40, 169)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(80, 16)
        Me.Label32.TabIndex = 112
        Me.Label32.Text = "処理経過状況"
        Me.Label32.UseCompatibleTextRendering = true
        '
        'lblTotalSituation
        '
        Me.lblTotalSituation.Location = New System.Drawing.Point(16, 29)
        Me.lblTotalSituation.Name = "lblTotalSituation"
        Me.lblTotalSituation.Size = New System.Drawing.Size(336, 16)
        Me.lblTotalSituation.TabIndex = 104
        Me.lblTotalSituation.Text = "コンバート処理中…"
        Me.lblTotalSituation.UseCompatibleTextRendering = true
        '
        'tabPageEndOK
        '
        Me.tabPageEndOK.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageEndOK.Controls.Add(Me.lblDatacvEndDescription1)
        Me.tabPageEndOK.Controls.Add(Me.lblKDatacvEndLabel)
        Me.tabPageEndOK.Controls.Add(Me.lblHDatacvEndLabel)
        Me.tabPageEndOK.Controls.Add(Me.PictureBox15)
        Me.tabPageEndOK.Location = New System.Drawing.Point(4, 27)
        Me.tabPageEndOK.Name = "tabPageEndOK"
        Me.tabPageEndOK.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageEndOK.Size = New System.Drawing.Size(912, 549)
        Me.tabPageEndOK.TabIndex = 4
        Me.tabPageEndOK.Text = " 終了"
        '
        'lblDatacvEndDescription1
        '
        Me.lblDatacvEndDescription1.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblDatacvEndDescription1.ForeColor = System.Drawing.Color.Navy
        Me.lblDatacvEndDescription1.Location = New System.Drawing.Point(34, 49)
        Me.lblDatacvEndDescription1.Name = "lblDatacvEndDescription1"
        Me.lblDatacvEndDescription1.Size = New System.Drawing.Size(820, 51)
        Me.lblDatacvEndDescription1.TabIndex = 129
        Me.lblDatacvEndDescription1.Text = "[ログ確認]ボタンよりコンバート結果を確認して下さい。"
        Me.lblDatacvEndDescription1.UseCompatibleTextRendering = true
        '
        'lblKDatacvEndLabel
        '
        Me.lblKDatacvEndLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblKDatacvEndLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.lblKDatacvEndLabel.Location = New System.Drawing.Point(30, 20)
        Me.lblKDatacvEndLabel.Name = "lblKDatacvEndLabel"
        Me.lblKDatacvEndLabel.Size = New System.Drawing.Size(820, 80)
        Me.lblKDatacvEndLabel.TabIndex = 91
        Me.lblKDatacvEndLabel.Text = "「賃貸革命V7」から「賃貸革命10」のコンバートが完了しました。"
        Me.lblKDatacvEndLabel.UseCompatibleTextRendering = true
        '
        'lblHDatacvEndLabel
        '
        Me.lblHDatacvEndLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblHDatacvEndLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(192,Byte),Integer))
        Me.lblHDatacvEndLabel.Location = New System.Drawing.Point(30, 20)
        Me.lblHDatacvEndLabel.Name = "lblHDatacvEndLabel"
        Me.lblHDatacvEndLabel.Size = New System.Drawing.Size(820, 80)
        Me.lblHDatacvEndLabel.TabIndex = 131
        Me.lblHDatacvEndLabel.Text = "「賃貸革命10」のコンバートが完了しました。"
        Me.lblHDatacvEndLabel.UseCompatibleTextRendering = true
        '
        'PictureBox15
        '
        Me.PictureBox15.Image = CType(resources.GetObject("PictureBox15.Image"),System.Drawing.Image)
        Me.PictureBox15.Location = New System.Drawing.Point(466, 360)
        Me.PictureBox15.Name = "PictureBox15"
        Me.PictureBox15.Size = New System.Drawing.Size(420, 170)
        Me.PictureBox15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox15.TabIndex = 130
        Me.PictureBox15.TabStop = false
        '
        'tabPageEndError
        '
        Me.tabPageEndError.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageEndError.Controls.Add(Me.GroupBox23)
        Me.tabPageEndError.Controls.Add(Me.lblDatacvErrorDescription1)
        Me.tabPageEndError.Controls.Add(Me.GroupBox19)
        Me.tabPageEndError.Controls.Add(Me.GroupBox20)
        Me.tabPageEndError.Controls.Add(Me.lblDatacvErrorLabel)
        Me.tabPageEndError.Controls.Add(Me.PictureBox23)
        Me.tabPageEndError.Location = New System.Drawing.Point(4, 27)
        Me.tabPageEndError.Name = "tabPageEndError"
        Me.tabPageEndError.Size = New System.Drawing.Size(912, 549)
        Me.tabPageEndError.TabIndex = 10
        Me.tabPageEndError.Text = "異常終了"
        '
        'GroupBox23
        '
        Me.GroupBox23.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox23.Controls.Add(Me.Label163)
        Me.GroupBox23.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox23.Location = New System.Drawing.Point(31, 329)
        Me.GroupBox23.Name = "GroupBox23"
        Me.GroupBox23.Size = New System.Drawing.Size(420, 200)
        Me.GroupBox23.TabIndex = 134
        Me.GroupBox23.TabStop = false
        Me.GroupBox23.Text = "【対処方法例】"
        '
        'Label163
        '
        Me.Label163.Location = New System.Drawing.Point(10, 25)
        Me.Label163.Name = "Label163"
        Me.Label163.Size = New System.Drawing.Size(400, 150)
        Me.Label163.TabIndex = 117
        Me.Label163.Text = "(例) 接続設定情報画面の賃貸革命V7と10の接続情報が正しく登録されているか確認して下さい。"
        Me.Label163.UseCompatibleTextRendering = true
        '
        'lblDatacvErrorDescription1
        '
        Me.lblDatacvErrorDescription1.ForeColor = System.Drawing.Color.Red
        Me.lblDatacvErrorDescription1.Location = New System.Drawing.Point(31, 47)
        Me.lblDatacvErrorDescription1.Name = "lblDatacvErrorDescription1"
        Me.lblDatacvErrorDescription1.Size = New System.Drawing.Size(820, 51)
        Me.lblDatacvErrorDescription1.TabIndex = 134
        Me.lblDatacvErrorDescription1.Text = "コンバート処理を正常に終了することができませんでした。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"[ログ確認]ボタンよりコンバート結果を確認して下さい。"
        Me.lblDatacvErrorDescription1.UseCompatibleTextRendering = true
        '
        'GroupBox19
        '
        Me.GroupBox19.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox19.Controls.Add(Me.Label170)
        Me.GroupBox19.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox19.Location = New System.Drawing.Point(466, 123)
        Me.GroupBox19.Name = "GroupBox19"
        Me.GroupBox19.Size = New System.Drawing.Size(420, 200)
        Me.GroupBox19.TabIndex = 133
        Me.GroupBox19.TabStop = false
        Me.GroupBox19.Text = "【エラー内容】"
        '
        'Label170
        '
        Me.Label170.Location = New System.Drawing.Point(10, 25)
        Me.Label170.Name = "Label170"
        Me.Label170.Size = New System.Drawing.Size(400, 150)
        Me.Label170.TabIndex = 117
        Me.Label170.Text = "(例) 賃貸革命V7と10の設定値が逆になっている可能性があります。"
        Me.Label170.UseCompatibleTextRendering = true
        '
        'GroupBox20
        '
        Me.GroupBox20.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox20.Controls.Add(Me.Label176)
        Me.GroupBox20.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox20.Location = New System.Drawing.Point(31, 123)
        Me.GroupBox20.Name = "GroupBox20"
        Me.GroupBox20.Size = New System.Drawing.Size(420, 200)
        Me.GroupBox20.TabIndex = 132
        Me.GroupBox20.TabStop = false
        Me.GroupBox20.Text = "【エラー個所】"
        '
        'Label176
        '
        Me.Label176.Location = New System.Drawing.Point(10, 25)
        Me.Label176.Name = "Label176"
        Me.Label176.Size = New System.Drawing.Size(400, 150)
        Me.Label176.TabIndex = 117
        Me.Label176.Text = "(例) 接続設定画面"
        Me.Label176.UseCompatibleTextRendering = true
        '
        'lblDatacvErrorLabel
        '
        Me.lblDatacvErrorLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblDatacvErrorLabel.ForeColor = System.Drawing.Color.Red
        Me.lblDatacvErrorLabel.Location = New System.Drawing.Point(27, 18)
        Me.lblDatacvErrorLabel.Name = "lblDatacvErrorLabel"
        Me.lblDatacvErrorLabel.Size = New System.Drawing.Size(820, 80)
        Me.lblDatacvErrorLabel.TabIndex = 131
        Me.lblDatacvErrorLabel.Text = "コンバートを異常終了しました。"
        Me.lblDatacvErrorLabel.UseCompatibleTextRendering = true
        '
        'PictureBox23
        '
        Me.PictureBox23.Image = CType(resources.GetObject("PictureBox23.Image"),System.Drawing.Image)
        Me.PictureBox23.Location = New System.Drawing.Point(466, 360)
        Me.PictureBox23.Name = "PictureBox23"
        Me.PictureBox23.Size = New System.Drawing.Size(420, 170)
        Me.PictureBox23.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox23.TabIndex = 135
        Me.PictureBox23.TabStop = false
        '
        'tabPageEndCancel
        '
        Me.tabPageEndCancel.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageEndCancel.Controls.Add(Me.lblDatacvCancelDescription1)
        Me.tabPageEndCancel.Controls.Add(Me.GroupBox21)
        Me.tabPageEndCancel.Controls.Add(Me.GroupBox22)
        Me.tabPageEndCancel.Controls.Add(Me.lblDatacvCancelLabel)
        Me.tabPageEndCancel.Controls.Add(Me.PictureBox25)
        Me.tabPageEndCancel.Location = New System.Drawing.Point(4, 27)
        Me.tabPageEndCancel.Name = "tabPageEndCancel"
        Me.tabPageEndCancel.Size = New System.Drawing.Size(912, 549)
        Me.tabPageEndCancel.TabIndex = 11
        Me.tabPageEndCancel.Text = "キャンセル終了"
        '
        'lblDatacvCancelDescription1
        '
        Me.lblDatacvCancelDescription1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblDatacvCancelDescription1.Location = New System.Drawing.Point(31, 47)
        Me.lblDatacvCancelDescription1.Name = "lblDatacvCancelDescription1"
        Me.lblDatacvCancelDescription1.Size = New System.Drawing.Size(820, 51)
        Me.lblDatacvCancelDescription1.TabIndex = 134
        Me.lblDatacvCancelDescription1.Text = "途中キャンセルした場合は不完全なデータの移行となる為、データの整合性が取れない場合があります。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"完全なデータの移行を行う為、本プログラムを再起動し、コンバート作"& _ 
    "業を再度行って下さい。"
        Me.lblDatacvCancelDescription1.UseCompatibleTextRendering = true
        '
        'GroupBox21
        '
        Me.GroupBox21.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox21.Controls.Add(Me.Label186)
        Me.GroupBox21.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox21.Location = New System.Drawing.Point(466, 123)
        Me.GroupBox21.Name = "GroupBox21"
        Me.GroupBox21.Size = New System.Drawing.Size(420, 200)
        Me.GroupBox21.TabIndex = 133
        Me.GroupBox21.TabStop = false
        Me.GroupBox21.Text = "【キャンセル内容】"
        '
        'Label186
        '
        Me.Label186.Location = New System.Drawing.Point(10, 25)
        Me.Label186.Name = "Label186"
        Me.Label186.Size = New System.Drawing.Size(400, 150)
        Me.Label186.TabIndex = 117
        Me.Label186.Text = "(例) 紐付設定処理中"
        Me.Label186.UseCompatibleTextRendering = true
        '
        'GroupBox22
        '
        Me.GroupBox22.BackColor = System.Drawing.SystemColors.Menu
        Me.GroupBox22.Controls.Add(Me.Label192)
        Me.GroupBox22.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.GroupBox22.Location = New System.Drawing.Point(31, 123)
        Me.GroupBox22.Name = "GroupBox22"
        Me.GroupBox22.Size = New System.Drawing.Size(400, 200)
        Me.GroupBox22.TabIndex = 132
        Me.GroupBox22.TabStop = false
        Me.GroupBox22.Text = "【キャンセル箇所】"
        '
        'Label192
        '
        Me.Label192.Location = New System.Drawing.Point(10, 25)
        Me.Label192.Name = "Label192"
        Me.Label192.Size = New System.Drawing.Size(384, 150)
        Me.Label192.TabIndex = 117
        Me.Label192.Text = "(例) 紐付設定画面"
        Me.Label192.UseCompatibleTextRendering = true
        '
        'lblDatacvCancelLabel
        '
        Me.lblDatacvCancelLabel.Font = New System.Drawing.Font("メイリオ", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblDatacvCancelLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.lblDatacvCancelLabel.Location = New System.Drawing.Point(27, 18)
        Me.lblDatacvCancelLabel.Name = "lblDatacvCancelLabel"
        Me.lblDatacvCancelLabel.Size = New System.Drawing.Size(820, 80)
        Me.lblDatacvCancelLabel.TabIndex = 131
        Me.lblDatacvCancelLabel.Text = "コンバート処理がキャンセルされました。"
        Me.lblDatacvCancelLabel.UseCompatibleTextRendering = true
        '
        'PictureBox25
        '
        Me.PictureBox25.Image = CType(resources.GetObject("PictureBox25.Image"),System.Drawing.Image)
        Me.PictureBox25.Location = New System.Drawing.Point(466, 360)
        Me.PictureBox25.Name = "PictureBox25"
        Me.PictureBox25.Size = New System.Drawing.Size(420, 170)
        Me.PictureBox25.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox25.TabIndex = 135
        Me.PictureBox25.TabStop = false
        '
        'tabPageIkkatu
        '
        Me.tabPageIkkatu.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageIkkatu.Controls.Add(Me.Label114)
        Me.tabPageIkkatu.Controls.Add(Me.Label144)
        Me.tabPageIkkatu.Location = New System.Drawing.Point(4, 27)
        Me.tabPageIkkatu.Name = "tabPageIkkatu"
        Me.tabPageIkkatu.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageIkkatu.Size = New System.Drawing.Size(912, 549)
        Me.tabPageIkkatu.TabIndex = 9
        Me.tabPageIkkatu.Text = " 必須項目一括設定(未定)"
        '
        'Label114
        '
        Me.Label114.AutoSize = true
        Me.Label114.Font = New System.Drawing.Font("メイリオ", 72!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label114.Location = New System.Drawing.Point(235, 215)
        Me.Label114.Name = "Label114"
        Me.Label114.Size = New System.Drawing.Size(348, 144)
        Me.Label114.TabIndex = 37
        Me.Label114.Text = "※未定"
        '
        'Label144
        '
        Me.Label144.BackColor = System.Drawing.SystemColors.Menu
        Me.Label144.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label144.ForeColor = System.Drawing.Color.Navy
        Me.Label144.Location = New System.Drawing.Point(30, 20)
        Me.Label144.Name = "Label144"
        Me.Label144.Size = New System.Drawing.Size(820, 80)
        Me.Label144.TabIndex = 36
        Me.Label144.Text = "以下の必須項目について設定下さい。設定したら「次へ」を押して下さい。"
        Me.Label144.UseCompatibleTextRendering = true
        '
        'tabPageHanyoJizen
        '
        Me.tabPageHanyoJizen.BackColor = System.Drawing.SystemColors.Menu
        Me.tabPageHanyoJizen.Controls.Add(Me.Label139)
        Me.tabPageHanyoJizen.Controls.Add(Me.Label145)
        Me.tabPageHanyoJizen.Controls.Add(Me.GroupBox11)
        Me.tabPageHanyoJizen.Location = New System.Drawing.Point(4, 27)
        Me.tabPageHanyoJizen.Name = "tabPageHanyoJizen"
        Me.tabPageHanyoJizen.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPageHanyoJizen.Size = New System.Drawing.Size(912, 549)
        Me.tabPageHanyoJizen.TabIndex = 8
        Me.tabPageHanyoJizen.Text = " 事前調整作業(汎用)"
        '
        'Label139
        '
        Me.Label139.BackColor = System.Drawing.SystemColors.Menu
        Me.Label139.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label139.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192,Byte),Integer), CType(CType(0,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.Label139.Location = New System.Drawing.Point(30, 90)
        Me.Label139.Name = "Label139"
        Me.Label139.Size = New System.Drawing.Size(820, 20)
        Me.Label139.TabIndex = 117
        Me.Label139.Text = "※コンバート処理を実行する前に必ず対応下さい。対応しなかった場合、異なる期待結果となる場合があります。"
        Me.Label139.UseCompatibleTextRendering = true
        '
        'Label145
        '
        Me.Label145.BackColor = System.Drawing.SystemColors.Menu
        Me.Label145.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label145.ForeColor = System.Drawing.Color.Navy
        Me.Label145.Location = New System.Drawing.Point(30, 20)
        Me.Label145.Name = "Label145"
        Me.Label145.Size = New System.Drawing.Size(820, 80)
        Me.Label145.TabIndex = 116
        Me.Label145.Text = "事前調整作業を行う必要があります。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"以下の内容を確認し、必要な作業を完了させて下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"作業が完了したら「次へ」ボタンを押して下さい。"
        Me.Label145.UseCompatibleTextRendering = true
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.Label140)
        Me.GroupBox11.Controls.Add(Me.Label141)
        Me.GroupBox11.Controls.Add(Me.Label128)
        Me.GroupBox11.Controls.Add(Me.Label134)
        Me.GroupBox11.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.GroupBox11.Location = New System.Drawing.Point(30, 150)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Size = New System.Drawing.Size(420, 213)
        Me.GroupBox11.TabIndex = 114
        Me.GroupBox11.TabStop = false
        Me.GroupBox11.Text = "【コンバーター仕様による作業】"
        '
        'Label140
        '
        Me.Label140.Location = New System.Drawing.Point(34, 145)
        Me.Label140.Name = "Label140"
        Me.Label140.Size = New System.Drawing.Size(340, 65)
        Me.Label140.TabIndex = 133
        Me.Label140.Text = "賃貸革命V7のファイル形式に合わせる必要があります。形式が異なる場合は移行できません。ファイル形式については仕様書をご確認下さい。"
        Me.Label140.UseCompatibleTextRendering = true
        '
        'Label141
        '
        Me.Label141.Location = New System.Drawing.Point(20, 121)
        Me.Label141.Name = "Label141"
        Me.Label141.Size = New System.Drawing.Size(390, 20)
        Me.Label141.TabIndex = 132
        Me.Label141.Text = "● 画像ファイル形式をV7形式に調整"
        Me.Label141.UseCompatibleTextRendering = true
        '
        'Label128
        '
        Me.Label128.Location = New System.Drawing.Point(34, 50)
        Me.Label128.Name = "Label128"
        Me.Label128.Size = New System.Drawing.Size(340, 65)
        Me.Label128.TabIndex = 131
        Me.Label128.Text = "中間ファイルへ事前にデータを登録しておく必要があります。登録を完了させてからコンバート処理を実行して下さい。"
        Me.Label128.UseCompatibleTextRendering = true
        '
        'Label134
        '
        Me.Label134.Location = New System.Drawing.Point(20, 26)
        Me.Label134.Name = "Label134"
        Me.Label134.Size = New System.Drawing.Size(390, 20)
        Me.Label134.TabIndex = 130
        Me.Label134.Text = "● 中間ファイル登録作業"
        Me.Label134.UseCompatibleTextRendering = true
        '
        'pnlHJizenGazoKeisiki
        '
        Me.pnlHJizenGazoKeisiki.Controls.Add(Me.Label343)
        Me.pnlHJizenGazoKeisiki.Controls.Add(Me.chkHJizenGazoKeisiki)
        Me.pnlHJizenGazoKeisiki.Controls.Add(Me.Label344)
        Me.pnlHJizenGazoKeisiki.Location = New System.Drawing.Point(1155, 6)
        Me.pnlHJizenGazoKeisiki.Name = "pnlHJizenGazoKeisiki"
        Me.pnlHJizenGazoKeisiki.Size = New System.Drawing.Size(81, 21)
        Me.pnlHJizenGazoKeisiki.TabIndex = 1
        Me.pnlHJizenGazoKeisiki.Visible = false
        '
        'Label343
        '
        Me.Label343.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label343.ForeColor = System.Drawing.Color.Gray
        Me.Label343.Location = New System.Drawing.Point(51, 133)
        Me.Label343.Name = "Label343"
        Me.Label343.Size = New System.Drawing.Size(340, 151)
        Me.Label343.TabIndex = 2
        Me.Label343.Text = resources.GetString("Label343.Text")
        Me.Label343.UseCompatibleTextRendering = true
        '
        'chkHJizenGazoKeisiki
        '
        Me.chkHJizenGazoKeisiki.AutoSize = true
        Me.chkHJizenGazoKeisiki.Font = New System.Drawing.Font("メイリオ", 9!, CType((System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Strikeout),System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkHJizenGazoKeisiki.ForeColor = System.Drawing.Color.Gray
        Me.chkHJizenGazoKeisiki.Location = New System.Drawing.Point(16, 4)
        Me.chkHJizenGazoKeisiki.Name = "chkHJizenGazoKeisiki"
        Me.chkHJizenGazoKeisiki.Size = New System.Drawing.Size(159, 22)
        Me.chkHJizenGazoKeisiki.TabIndex = 0
        Me.chkHJizenGazoKeisiki.Text = "画像ファイル名調整作業"
        Me.chkHJizenGazoKeisiki.UseVisualStyleBackColor = true
        '
        'Label344
        '
        Me.Label344.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label344.ForeColor = System.Drawing.Color.Gray
        Me.Label344.Location = New System.Drawing.Point(36, 29)
        Me.Label344.Name = "Label344"
        Me.Label344.Size = New System.Drawing.Size(340, 100)
        Me.Label344.TabIndex = 1
        Me.Label344.Text = "ユーザで用意された画像ファイルについて、本プログラムの仕様(ファイル名)に合わせる必要があります。以下の内容を参考に、画像ファイルを調整して下さい。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※画像ファ"& _ 
    "イルだけのコンバートはできません。(例：物件画像をコンバートする場合、物件情報(親データ)が必要です)"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"※詳細は手順書を参照下さい。"
        Me.Label344.UseCompatibleTextRendering = true
        '
        'pnlJizenSo
        '
        Me.pnlJizenSo.Controls.Add(Me.chkJizenSo)
        Me.pnlJizenSo.Controls.Add(Me.Label90)
        Me.pnlJizenSo.Location = New System.Drawing.Point(1086, 6)
        Me.pnlJizenSo.Name = "pnlJizenSo"
        Me.pnlJizenSo.Size = New System.Drawing.Size(66, 21)
        Me.pnlJizenSo.TabIndex = 8
        Me.pnlJizenSo.Visible = false
        '
        'chkJizenSo
        '
        Me.chkJizenSo.AutoSize = true
        Me.chkJizenSo.Font = New System.Drawing.Font("メイリオ", 9!, CType((System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Strikeout),System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.chkJizenSo.Location = New System.Drawing.Point(16, 4)
        Me.chkJizenSo.Name = "chkJizenSo"
        Me.chkJizenSo.Size = New System.Drawing.Size(75, 22)
        Me.chkJizenSo.TabIndex = 8
        Me.chkJizenSo.TabStop = false
        Me.chkJizenSo.Text = "送金処理"
        Me.chkJizenSo.UseVisualStyleBackColor = true
        '
        'Label90
        '
        Me.Label90.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label90.Location = New System.Drawing.Point(36, 29)
        Me.Label90.Name = "Label90"
        Me.Label90.Size = New System.Drawing.Size(340, 33)
        Me.Label90.TabIndex = 9
        Me.Label90.Text = "送金データは移行対象外です。送金確定していないデータについて、賃貸革命V7で送金確定を行って下さい。"
        Me.Label90.UseCompatibleTextRendering = true
        '
        'btnDevTabChange
        '
        Me.btnDevTabChange.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnDevTabChange.Location = New System.Drawing.Point(66, 638)
        Me.btnDevTabChange.Name = "btnDevTabChange"
        Me.btnDevTabChange.Size = New System.Drawing.Size(180, 30)
        Me.btnDevTabChange.TabIndex = 4
        Me.btnDevTabChange.Text = "開発用画面表示ON/OFF"
        Me.btnDevTabChange.UseVisualStyleBackColor = true
        Me.btnDevTabChange.Visible = false
        '
        'Label393
        '
        Me.Label393.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label393.Location = New System.Drawing.Point(40, 186)
        Me.Label393.Name = "Label393"
        Me.Label393.Size = New System.Drawing.Size(340, 65)
        Me.Label393.TabIndex = 5
        Me.Label393.UseCompatibleTextRendering = true
        '
        'Label392
        '
        Me.Label392.AutoSize = true
        Me.Label392.Location = New System.Drawing.Point(794, 310)
        Me.Label392.Name = "Label392"
        Me.Label392.Size = New System.Drawing.Size(35, 18)
        Me.Label392.TabIndex = 0
        '
        'CheckBox51
        '
        Me.CheckBox51.AutoSize = true
        Me.CheckBox51.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.CheckBox51.Location = New System.Drawing.Point(20, 161)
        Me.CheckBox51.Name = "CheckBox51"
        Me.CheckBox51.Size = New System.Drawing.Size(147, 22)
        Me.CheckBox51.TabIndex = 4
        Me.CheckBox51.Text = "控除支払データの確認"
        Me.CheckBox51.UseVisualStyleBackColor = true
        '
        'Label391
        '
        Me.Label391.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label391.Location = New System.Drawing.Point(40, 50)
        Me.Label391.Name = "Label391"
        Me.Label391.Size = New System.Drawing.Size(340, 33)
        Me.Label391.TabIndex = 1
        Me.Label391.UseCompatibleTextRendering = true
        '
        'CheckBox50
        '
        Me.CheckBox50.AutoSize = true
        Me.CheckBox50.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.CheckBox50.Location = New System.Drawing.Point(20, 25)
        Me.CheckBox50.Name = "CheckBox50"
        Me.CheckBox50.Size = New System.Drawing.Size(195, 22)
        Me.CheckBox50.TabIndex = 0
        Me.CheckBox50.Text = "分割入金の未収分データの確認"
        Me.CheckBox50.UseVisualStyleBackColor = true
        '
        'Panel24
        '
        Me.Panel24.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel24.Location = New System.Drawing.Point(40, 254)
        Me.Panel24.Name = "Panel24"
        Me.Panel24.Size = New System.Drawing.Size(340, 60)
        Me.Panel24.TabIndex = 6
        '
        'Label390
        '
        Me.Label390.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label390.Location = New System.Drawing.Point(218, 31)
        Me.Label390.Name = "Label390"
        Me.Label390.Size = New System.Drawing.Size(83, 16)
        Me.Label390.TabIndex = 3
        Me.Label390.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label389
        '
        Me.Label389.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label389.Location = New System.Drawing.Point(131, 31)
        Me.Label389.Name = "Label389"
        Me.Label389.Size = New System.Drawing.Size(92, 16)
        Me.Label389.TabIndex = 2
        Me.Label389.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label388
        '
        Me.Label388.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label388.Location = New System.Drawing.Point(307, 31)
        Me.Label388.Name = "Label388"
        Me.Label388.Size = New System.Drawing.Size(22, 16)
        Me.Label388.TabIndex = 4
        Me.Label388.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label387
        '
        Me.Label387.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label387.Location = New System.Drawing.Point(8, 8)
        Me.Label387.Name = "Label387"
        Me.Label387.Size = New System.Drawing.Size(321, 16)
        Me.Label387.TabIndex = 0
        Me.Label387.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label387.UseCompatibleTextRendering = true
        '
        'Panel23
        '
        Me.Panel23.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel23.Location = New System.Drawing.Point(40, 86)
        Me.Panel23.Name = "Panel23"
        Me.Panel23.Size = New System.Drawing.Size(340, 60)
        Me.Panel23.TabIndex = 3
        '
        'Label386
        '
        Me.Label386.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label386.Location = New System.Drawing.Point(218, 31)
        Me.Label386.Name = "Label386"
        Me.Label386.Size = New System.Drawing.Size(83, 16)
        Me.Label386.TabIndex = 3
        Me.Label386.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label385
        '
        Me.Label385.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label385.Location = New System.Drawing.Point(131, 31)
        Me.Label385.Name = "Label385"
        Me.Label385.Size = New System.Drawing.Size(92, 16)
        Me.Label385.TabIndex = 2
        Me.Label385.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label384
        '
        Me.Label384.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label384.Location = New System.Drawing.Point(307, 31)
        Me.Label384.Name = "Label384"
        Me.Label384.Size = New System.Drawing.Size(22, 16)
        Me.Label384.TabIndex = 4
        Me.Label384.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label383
        '
        Me.Label383.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label383.Location = New System.Drawing.Point(8, 8)
        Me.Label383.Name = "Label383"
        Me.Label383.Size = New System.Drawing.Size(321, 16)
        Me.Label383.TabIndex = 0
        Me.Label383.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label383.UseCompatibleTextRendering = true
        '
        'Label382
        '
        Me.Label382.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label382.Location = New System.Drawing.Point(461, 50)
        Me.Label382.Name = "Label382"
        Me.Label382.Size = New System.Drawing.Size(340, 85)
        Me.Label382.TabIndex = 8
        Me.Label382.UseCompatibleTextRendering = true
        '
        'CheckBox49
        '
        Me.CheckBox49.AutoSize = true
        Me.CheckBox49.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.CheckBox49.Location = New System.Drawing.Point(441, 25)
        Me.CheckBox49.Name = "CheckBox49"
        Me.CheckBox49.Size = New System.Drawing.Size(231, 22)
        Me.CheckBox49.TabIndex = 7
        Me.CheckBox49.Text = "修繕項目毎契約者送金率データの確認"
        Me.CheckBox49.UseVisualStyleBackColor = true
        '
        'Panel22
        '
        Me.Panel22.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel22.Location = New System.Drawing.Point(461, 138)
        Me.Panel22.Name = "Panel22"
        Me.Panel22.Size = New System.Drawing.Size(340, 60)
        Me.Panel22.TabIndex = 9
        '
        'Label381
        '
        Me.Label381.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label381.Location = New System.Drawing.Point(218, 31)
        Me.Label381.Name = "Label381"
        Me.Label381.Size = New System.Drawing.Size(83, 16)
        Me.Label381.TabIndex = 3
        Me.Label381.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label380
        '
        Me.Label380.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label380.Location = New System.Drawing.Point(131, 31)
        Me.Label380.Name = "Label380"
        Me.Label380.Size = New System.Drawing.Size(92, 16)
        Me.Label380.TabIndex = 2
        Me.Label380.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label379
        '
        Me.Label379.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label379.Location = New System.Drawing.Point(307, 31)
        Me.Label379.Name = "Label379"
        Me.Label379.Size = New System.Drawing.Size(22, 16)
        Me.Label379.TabIndex = 4
        Me.Label379.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label378
        '
        Me.Label378.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Label378.Location = New System.Drawing.Point(8, 8)
        Me.Label378.Name = "Label378"
        Me.Label378.Size = New System.Drawing.Size(321, 16)
        Me.Label378.TabIndex = 0
        Me.Label378.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label378.UseCompatibleTextRendering = true
        '
        'pnlMenuMain
        '
        Me.pnlMenuMain.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlMenuMain.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlMenuMain.Controls.Add(Me.pnlPrgMStart)
        Me.pnlMenuMain.Controls.Add(Me.picArwMSession2)
        Me.pnlMenuMain.Controls.Add(Me.picArwMSession1)
        Me.pnlMenuMain.Controls.Add(Me.pnlPrgMEnd)
        Me.pnlMenuMain.Controls.Add(Me.pnlPrgMMenu)
        Me.pnlMenuMain.Controls.Add(Me.pnlPrgMSyoki)
        Me.pnlMenuMain.Controls.Add(Me.pnlPrgMSession)
        Me.pnlMenuMain.Controls.Add(Me.picArwMEnd2)
        Me.pnlMenuMain.Controls.Add(Me.picArwMMenu2)
        Me.pnlMenuMain.Controls.Add(Me.picArwMSyoki2)
        Me.pnlMenuMain.Controls.Add(Me.picArwMEnd1)
        Me.pnlMenuMain.Controls.Add(Me.picArwMMenu1)
        Me.pnlMenuMain.Controls.Add(Me.picArwMSyoki1)
        Me.pnlMenuMain.Location = New System.Drawing.Point(33, 60)
        Me.pnlMenuMain.Name = "pnlMenuMain"
        Me.pnlMenuMain.Size = New System.Drawing.Size(250, 561)
        Me.pnlMenuMain.TabIndex = 138
        '
        'pnlPrgMStart
        '
        Me.pnlPrgMStart.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgMStart.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgMStart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgMStart.CausesValidation = false
        Me.pnlPrgMStart.Controls.Add(Me.picIcoMStart2)
        Me.pnlPrgMStart.Controls.Add(Me.picIcoMStart1)
        Me.pnlPrgMStart.Controls.Add(Me.PictureBox3)
        Me.pnlPrgMStart.Controls.Add(Me.lblPrgMStart)
        Me.pnlPrgMStart.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.pnlPrgMStart.Location = New System.Drawing.Point(12, 8)
        Me.pnlPrgMStart.Name = "pnlPrgMStart"
        Me.pnlPrgMStart.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgMStart.TabIndex = 143
        '
        'picIcoMStart2
        '
        Me.picIcoMStart2.Image = CType(resources.GetObject("picIcoMStart2.Image"),System.Drawing.Image)
        Me.picIcoMStart2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMStart2.Name = "picIcoMStart2"
        Me.picIcoMStart2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMStart2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMStart2.TabIndex = 143
        Me.picIcoMStart2.TabStop = false
        '
        'picIcoMStart1
        '
        Me.picIcoMStart1.Image = CType(resources.GetObject("picIcoMStart1.Image"),System.Drawing.Image)
        Me.picIcoMStart1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMStart1.Name = "picIcoMStart1"
        Me.picIcoMStart1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMStart1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMStart1.TabIndex = 3
        Me.picIcoMStart1.TabStop = false
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"),System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox3.TabIndex = 2
        Me.PictureBox3.TabStop = false
        '
        'lblPrgMStart
        '
        Me.lblPrgMStart.AutoSize = true
        Me.lblPrgMStart.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgMStart.ForeColor = System.Drawing.Color.Black
        Me.lblPrgMStart.Location = New System.Drawing.Point(56, 9)
        Me.lblPrgMStart.Name = "lblPrgMStart"
        Me.lblPrgMStart.Size = New System.Drawing.Size(126, 31)
        Me.lblPrgMStart.TabIndex = 0
        Me.lblPrgMStart.Text = " 作 業 開 始"
        '
        'picArwMSession2
        '
        Me.picArwMSession2.Image = CType(resources.GetObject("picArwMSession2.Image"),System.Drawing.Image)
        Me.picArwMSession2.Location = New System.Drawing.Point(103, 62)
        Me.picArwMSession2.Name = "picArwMSession2"
        Me.picArwMSession2.Size = New System.Drawing.Size(40, 35)
        Me.picArwMSession2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMSession2.TabIndex = 145
        Me.picArwMSession2.TabStop = false
        '
        'picArwMSession1
        '
        Me.picArwMSession1.Image = CType(resources.GetObject("picArwMSession1.Image"),System.Drawing.Image)
        Me.picArwMSession1.Location = New System.Drawing.Point(103, 62)
        Me.picArwMSession1.Name = "picArwMSession1"
        Me.picArwMSession1.Size = New System.Drawing.Size(40, 35)
        Me.picArwMSession1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMSession1.TabIndex = 144
        Me.picArwMSession1.TabStop = false
        '
        'pnlPrgMEnd
        '
        Me.pnlPrgMEnd.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgMEnd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgMEnd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgMEnd.CausesValidation = false
        Me.pnlPrgMEnd.Controls.Add(Me.picIcoMEnd2)
        Me.pnlPrgMEnd.Controls.Add(Me.picIcoMEnd1)
        Me.pnlPrgMEnd.Controls.Add(Me.picIcoJizen1)
        Me.pnlPrgMEnd.Controls.Add(Me.lblPrgMEnd)
        Me.pnlPrgMEnd.Location = New System.Drawing.Point(12, 394)
        Me.pnlPrgMEnd.Name = "pnlPrgMEnd"
        Me.pnlPrgMEnd.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgMEnd.TabIndex = 3
        '
        'picIcoMEnd2
        '
        Me.picIcoMEnd2.Image = CType(resources.GetObject("picIcoMEnd2.Image"),System.Drawing.Image)
        Me.picIcoMEnd2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMEnd2.Name = "picIcoMEnd2"
        Me.picIcoMEnd2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMEnd2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMEnd2.TabIndex = 146
        Me.picIcoMEnd2.TabStop = false
        '
        'picIcoMEnd1
        '
        Me.picIcoMEnd1.Image = CType(resources.GetObject("picIcoMEnd1.Image"),System.Drawing.Image)
        Me.picIcoMEnd1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMEnd1.Name = "picIcoMEnd1"
        Me.picIcoMEnd1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMEnd1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMEnd1.TabIndex = 4
        Me.picIcoMEnd1.TabStop = false
        '
        'picIcoJizen1
        '
        Me.picIcoJizen1.Image = CType(resources.GetObject("picIcoJizen1.Image"),System.Drawing.Image)
        Me.picIcoJizen1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoJizen1.Name = "picIcoJizen1"
        Me.picIcoJizen1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoJizen1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoJizen1.TabIndex = 3
        Me.picIcoJizen1.TabStop = false
        '
        'lblPrgMEnd
        '
        Me.lblPrgMEnd.AutoSize = true
        Me.lblPrgMEnd.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgMEnd.ForeColor = System.Drawing.Color.Black
        Me.lblPrgMEnd.Location = New System.Drawing.Point(53, 9)
        Me.lblPrgMEnd.Name = "lblPrgMEnd"
        Me.lblPrgMEnd.Size = New System.Drawing.Size(126, 31)
        Me.lblPrgMEnd.TabIndex = 0
        Me.lblPrgMEnd.Text = " 作 業 終 了"
        '
        'pnlPrgMMenu
        '
        Me.pnlPrgMMenu.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgMMenu.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgMMenu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgMMenu.CausesValidation = false
        Me.pnlPrgMMenu.Controls.Add(Me.picIcoMMenu2)
        Me.pnlPrgMMenu.Controls.Add(Me.picIcoMMenu1)
        Me.pnlPrgMMenu.Controls.Add(Me.picIcoSession1)
        Me.pnlPrgMMenu.Controls.Add(Me.lblPrgMMenu)
        Me.pnlPrgMMenu.Location = New System.Drawing.Point(12, 297)
        Me.pnlPrgMMenu.Name = "pnlPrgMMenu"
        Me.pnlPrgMMenu.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgMMenu.TabIndex = 2
        '
        'picIcoMMenu2
        '
        Me.picIcoMMenu2.Image = CType(resources.GetObject("picIcoMMenu2.Image"),System.Drawing.Image)
        Me.picIcoMMenu2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMMenu2.Name = "picIcoMMenu2"
        Me.picIcoMMenu2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMMenu2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMMenu2.TabIndex = 145
        Me.picIcoMMenu2.TabStop = false
        '
        'picIcoMMenu1
        '
        Me.picIcoMMenu1.Image = CType(resources.GetObject("picIcoMMenu1.Image"),System.Drawing.Image)
        Me.picIcoMMenu1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMMenu1.Name = "picIcoMMenu1"
        Me.picIcoMMenu1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMMenu1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMMenu1.TabIndex = 4
        Me.picIcoMMenu1.TabStop = false
        '
        'picIcoSession1
        '
        Me.picIcoSession1.Image = CType(resources.GetObject("picIcoSession1.Image"),System.Drawing.Image)
        Me.picIcoSession1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoSession1.Name = "picIcoSession1"
        Me.picIcoSession1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoSession1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoSession1.TabIndex = 3
        Me.picIcoSession1.TabStop = false
        '
        'lblPrgMMenu
        '
        Me.lblPrgMMenu.AutoSize = true
        Me.lblPrgMMenu.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgMMenu.ForeColor = System.Drawing.Color.Black
        Me.lblPrgMMenu.Location = New System.Drawing.Point(53, 9)
        Me.lblPrgMMenu.Name = "lblPrgMMenu"
        Me.lblPrgMMenu.Size = New System.Drawing.Size(126, 31)
        Me.lblPrgMMenu.TabIndex = 0
        Me.lblPrgMMenu.Text = " 作 業 選 択"
        '
        'pnlPrgMSyoki
        '
        Me.pnlPrgMSyoki.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgMSyoki.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgMSyoki.CausesValidation = false
        Me.pnlPrgMSyoki.Controls.Add(Me.picIcoMSyoki2)
        Me.pnlPrgMSyoki.Controls.Add(Me.picIcoMSyoki1)
        Me.pnlPrgMSyoki.Controls.Add(Me.picIcoSyoki1)
        Me.pnlPrgMSyoki.Controls.Add(Me.lblPrgMSyoki)
        Me.pnlPrgMSyoki.Font = New System.Drawing.Font("メイリオ", 9!)
        Me.pnlPrgMSyoki.Location = New System.Drawing.Point(12, 200)
        Me.pnlPrgMSyoki.Name = "pnlPrgMSyoki"
        Me.pnlPrgMSyoki.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgMSyoki.TabIndex = 1
        '
        'picIcoMSyoki2
        '
        Me.picIcoMSyoki2.Image = CType(resources.GetObject("picIcoMSyoki2.Image"),System.Drawing.Image)
        Me.picIcoMSyoki2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMSyoki2.Name = "picIcoMSyoki2"
        Me.picIcoMSyoki2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMSyoki2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMSyoki2.TabIndex = 144
        Me.picIcoMSyoki2.TabStop = false
        '
        'picIcoMSyoki1
        '
        Me.picIcoMSyoki1.Image = CType(resources.GetObject("picIcoMSyoki1.Image"),System.Drawing.Image)
        Me.picIcoMSyoki1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMSyoki1.Name = "picIcoMSyoki1"
        Me.picIcoMSyoki1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMSyoki1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMSyoki1.TabIndex = 4
        Me.picIcoMSyoki1.TabStop = false
        '
        'picIcoSyoki1
        '
        Me.picIcoSyoki1.Image = CType(resources.GetObject("picIcoSyoki1.Image"),System.Drawing.Image)
        Me.picIcoSyoki1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoSyoki1.Name = "picIcoSyoki1"
        Me.picIcoSyoki1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoSyoki1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoSyoki1.TabIndex = 3
        Me.picIcoSyoki1.TabStop = false
        '
        'lblPrgMSyoki
        '
        Me.lblPrgMSyoki.AutoSize = true
        Me.lblPrgMSyoki.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgMSyoki.ForeColor = System.Drawing.Color.Black
        Me.lblPrgMSyoki.Location = New System.Drawing.Point(53, 9)
        Me.lblPrgMSyoki.Name = "lblPrgMSyoki"
        Me.lblPrgMSyoki.Size = New System.Drawing.Size(119, 31)
        Me.lblPrgMSyoki.TabIndex = 0
        Me.lblPrgMSyoki.Text = "初 期 設 定"
        '
        'pnlPrgMSession
        '
        Me.pnlPrgMSession.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlPrgMSession.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pnlPrgMSession.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgMSession.CausesValidation = false
        Me.pnlPrgMSession.Controls.Add(Me.picIcoMSession2)
        Me.pnlPrgMSession.Controls.Add(Me.picIcoMSession1)
        Me.pnlPrgMSession.Controls.Add(Me.PictureBox4)
        Me.pnlPrgMSession.Controls.Add(Me.lblPrgMSession)
        Me.pnlPrgMSession.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.pnlPrgMSession.Location = New System.Drawing.Point(12, 103)
        Me.pnlPrgMSession.Name = "pnlPrgMSession"
        Me.pnlPrgMSession.Size = New System.Drawing.Size(221, 47)
        Me.pnlPrgMSession.TabIndex = 0
        '
        'picIcoMSession2
        '
        Me.picIcoMSession2.Image = CType(resources.GetObject("picIcoMSession2.Image"),System.Drawing.Image)
        Me.picIcoMSession2.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMSession2.Name = "picIcoMSession2"
        Me.picIcoMSession2.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMSession2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMSession2.TabIndex = 143
        Me.picIcoMSession2.TabStop = false
        '
        'picIcoMSession1
        '
        Me.picIcoMSession1.Image = CType(resources.GetObject("picIcoMSession1.Image"),System.Drawing.Image)
        Me.picIcoMSession1.Location = New System.Drawing.Point(3, 3)
        Me.picIcoMSession1.Name = "picIcoMSession1"
        Me.picIcoMSession1.Size = New System.Drawing.Size(40, 40)
        Me.picIcoMSession1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcoMSession1.TabIndex = 3
        Me.picIcoMSession1.TabStop = false
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"),System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox4.TabIndex = 2
        Me.PictureBox4.TabStop = false
        '
        'lblPrgMSession
        '
        Me.lblPrgMSession.AutoSize = true
        Me.lblPrgMSession.Font = New System.Drawing.Font("メイリオ", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPrgMSession.ForeColor = System.Drawing.Color.Black
        Me.lblPrgMSession.Location = New System.Drawing.Point(56, 8)
        Me.lblPrgMSession.Name = "lblPrgMSession"
        Me.lblPrgMSession.Size = New System.Drawing.Size(126, 31)
        Me.lblPrgMSession.TabIndex = 0
        Me.lblPrgMSession.Text = "接 続 設 定 "
        '
        'picArwMEnd2
        '
        Me.picArwMEnd2.Image = CType(resources.GetObject("picArwMEnd2.Image"),System.Drawing.Image)
        Me.picArwMEnd2.Location = New System.Drawing.Point(103, 351)
        Me.picArwMEnd2.Name = "picArwMEnd2"
        Me.picArwMEnd2.Size = New System.Drawing.Size(40, 35)
        Me.picArwMEnd2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMEnd2.TabIndex = 141
        Me.picArwMEnd2.TabStop = false
        '
        'picArwMMenu2
        '
        Me.picArwMMenu2.Image = CType(resources.GetObject("picArwMMenu2.Image"),System.Drawing.Image)
        Me.picArwMMenu2.Location = New System.Drawing.Point(103, 254)
        Me.picArwMMenu2.Name = "picArwMMenu2"
        Me.picArwMMenu2.Size = New System.Drawing.Size(40, 35)
        Me.picArwMMenu2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMMenu2.TabIndex = 140
        Me.picArwMMenu2.TabStop = false
        '
        'picArwMSyoki2
        '
        Me.picArwMSyoki2.Image = CType(resources.GetObject("picArwMSyoki2.Image"),System.Drawing.Image)
        Me.picArwMSyoki2.Location = New System.Drawing.Point(103, 157)
        Me.picArwMSyoki2.Name = "picArwMSyoki2"
        Me.picArwMSyoki2.Size = New System.Drawing.Size(40, 35)
        Me.picArwMSyoki2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMSyoki2.TabIndex = 19
        Me.picArwMSyoki2.TabStop = false
        '
        'picArwMEnd1
        '
        Me.picArwMEnd1.Image = CType(resources.GetObject("picArwMEnd1.Image"),System.Drawing.Image)
        Me.picArwMEnd1.Location = New System.Drawing.Point(103, 351)
        Me.picArwMEnd1.Name = "picArwMEnd1"
        Me.picArwMEnd1.Size = New System.Drawing.Size(40, 35)
        Me.picArwMEnd1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMEnd1.TabIndex = 16
        Me.picArwMEnd1.TabStop = false
        '
        'picArwMMenu1
        '
        Me.picArwMMenu1.Image = CType(resources.GetObject("picArwMMenu1.Image"),System.Drawing.Image)
        Me.picArwMMenu1.Location = New System.Drawing.Point(103, 254)
        Me.picArwMMenu1.Name = "picArwMMenu1"
        Me.picArwMMenu1.Size = New System.Drawing.Size(40, 35)
        Me.picArwMMenu1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMMenu1.TabIndex = 5
        Me.picArwMMenu1.TabStop = false
        '
        'picArwMSyoki1
        '
        Me.picArwMSyoki1.Image = CType(resources.GetObject("picArwMSyoki1.Image"),System.Drawing.Image)
        Me.picArwMSyoki1.Location = New System.Drawing.Point(103, 157)
        Me.picArwMSyoki1.Name = "picArwMSyoki1"
        Me.picArwMSyoki1.Size = New System.Drawing.Size(40, 35)
        Me.picArwMSyoki1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picArwMSyoki1.TabIndex = 4
        Me.picArwMSyoki1.TabStop = false
        '
        'lblLine0
        '
        Me.lblLine0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLine0.Location = New System.Drawing.Point(319, 65)
        Me.lblLine0.Name = "lblLine0"
        Me.lblLine0.Size = New System.Drawing.Size(916, 2)
        Me.lblLine0.TabIndex = 139
        '
        'pnlRefresh
        '
        Me.pnlRefresh.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlRefresh.Controls.Add(Me.Label251)
        Me.pnlRefresh.Controls.Add(Me.btnRefresh)
        Me.pnlRefresh.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.pnlRefresh.Location = New System.Drawing.Point(316, 626)
        Me.pnlRefresh.Name = "pnlRefresh"
        Me.pnlRefresh.Size = New System.Drawing.Size(389, 50)
        Me.pnlRefresh.TabIndex = 46
        Me.pnlRefresh.Visible = false
        '
        'Label251
        '
        Me.Label251.ForeColor = System.Drawing.Color.Black
        Me.Label251.Location = New System.Drawing.Point(4, 5)
        Me.Label251.Name = "Label251"
        Me.Label251.Size = New System.Drawing.Size(278, 38)
        Me.Label251.TabIndex = 0
        Me.Label251.Text = "※表示されている抽出件数を再読み込みします。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"　データの調整を行った後に押して下さい。"
        '
        'btnRefresh
        '
        Me.btnRefresh.Font = New System.Drawing.Font("メイリオ", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"),System.Drawing.Image)
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefresh.Location = New System.Drawing.Point(288, 8)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnRefresh.Size = New System.Drawing.Size(88, 30)
        Me.btnRefresh.TabIndex = 2
        Me.btnRefresh.Text = " 再読込"
        Me.btnRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnRefresh.UseVisualStyleBackColor = true
        '
        'PictureBox24
        '
        Me.PictureBox24.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox24.Image = CType(resources.GetObject("PictureBox24.Image"),System.Drawing.Image)
        Me.PictureBox24.Location = New System.Drawing.Point(969, 1)
        Me.PictureBox24.Name = "PictureBox24"
        Me.PictureBox24.Size = New System.Drawing.Size(267, 31)
        Me.PictureBox24.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox24.TabIndex = 137
        Me.PictureBox24.TabStop = false
        Me.PictureBox24.Visible = false
        '
        'btnBack
        '
        Me.btnBack.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnBack.Font = New System.Drawing.Font("メイリオ", 11.25!)
        Me.btnBack.Image = CType(resources.GetObject("btnBack.Image"),System.Drawing.Image)
        Me.btnBack.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBack.Location = New System.Drawing.Point(817, 640)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnBack.Size = New System.Drawing.Size(120, 30)
        Me.btnBack.TabIndex = 5
        Me.btnBack.Text = "   戻  る"
        Me.btnBack.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnBack.UseVisualStyleBackColor = true
        '
        'btnNext
        '
        Me.btnNext.Font = New System.Drawing.Font("メイリオ", 11.25!)
        Me.btnNext.Image = CType(resources.GetObject("btnNext.Image"),System.Drawing.Image)
        Me.btnNext.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnNext.Location = New System.Drawing.Point(952, 640)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnNext.Size = New System.Drawing.Size(120, 30)
        Me.btnNext.TabIndex = 6
        Me.btnNext.Text = "   次  へ"
        Me.btnNext.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnNext.UseVisualStyleBackColor = true
        '
        'btnEnd
        '
        Me.btnEnd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnEnd.Font = New System.Drawing.Font("メイリオ", 11.25!)
        Me.btnEnd.Image = CType(resources.GetObject("btnEnd.Image"),System.Drawing.Image)
        Me.btnEnd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEnd.Location = New System.Drawing.Point(1089, 640)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.btnEnd.Size = New System.Drawing.Size(120, 30)
        Me.btnEnd.TabIndex = 7
        Me.btnEnd.Text = "   終  了"
        Me.btnEnd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnEnd.UseVisualStyleBackColor = true
        '
        'Button21
        '
        Me.Button21.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Button21.Image = CType(resources.GetObject("Button21.Image"),System.Drawing.Image)
        Me.Button21.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button21.Location = New System.Drawing.Point(8, 27)
        Me.Button21.Name = "Button21"
        Me.Button21.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.Button21.Size = New System.Drawing.Size(100, 25)
        Me.Button21.TabIndex = 1
        Me.Button21.Text = "リスト出力"
        Me.Button21.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button21.UseVisualStyleBackColor = true
        '
        'Button20
        '
        Me.Button20.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Button20.Image = CType(resources.GetObject("Button20.Image"),System.Drawing.Image)
        Me.Button20.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button20.Location = New System.Drawing.Point(8, 27)
        Me.Button20.Name = "Button20"
        Me.Button20.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.Button20.Size = New System.Drawing.Size(100, 25)
        Me.Button20.TabIndex = 1
        Me.Button20.Text = "リスト出力"
        Me.Button20.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button20.UseVisualStyleBackColor = true
        '
        'Button19
        '
        Me.Button19.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.Button19.Image = CType(resources.GetObject("Button19.Image"),System.Drawing.Image)
        Me.Button19.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button19.Location = New System.Drawing.Point(8, 27)
        Me.Button19.Name = "Button19"
        Me.Button19.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.Button19.Size = New System.Drawing.Size(100, 25)
        Me.Button19.TabIndex = 1
        Me.Button19.Text = "リスト出力"
        Me.Button19.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button19.UseVisualStyleBackColor = true
        '
        'lblTitleH
        '
        Me.lblTitleH.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTitleH.Font = New System.Drawing.Font("メイリオ", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic),System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblTitleH.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.lblTitleH.Location = New System.Drawing.Point(33, 18)
        Me.lblTitleH.Name = "lblTitleH"
        Me.lblTitleH.Size = New System.Drawing.Size(250, 26)
        Me.lblTitleH.TabIndex = 140
        Me.lblTitleH.Text = "[中間ファイル → 賃貸革命10]"
        Me.lblTitleH.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pgbCheck
        '
        Me.pgbCheck.Location = New System.Drawing.Point(176, 3)
        Me.pgbCheck.Name = "pgbCheck"
        Me.pgbCheck.Size = New System.Drawing.Size(190, 20)
        Me.pgbCheck.TabIndex = 141
        '
        'lblPgbCheck
        '
        Me.lblPgbCheck.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128,Byte))
        Me.lblPgbCheck.Location = New System.Drawing.Point(372, 4)
        Me.lblPgbCheck.Name = "lblPgbCheck"
        Me.lblPgbCheck.Size = New System.Drawing.Size(51, 20)
        Me.lblPgbCheck.TabIndex = 142
        Me.lblPgbCheck.Text = "0 %"
        Me.lblPgbCheck.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblPgbCheck.UseCompatibleTextRendering = true
        '
        'lblCheckSituation
        '
        Me.lblCheckSituation.Font = New System.Drawing.Font("メイリオ", 8.25!)
        Me.lblCheckSituation.Location = New System.Drawing.Point(4, 4)
        Me.lblCheckSituation.Name = "lblCheckSituation"
        Me.lblCheckSituation.Size = New System.Drawing.Size(166, 20)
        Me.lblCheckSituation.TabIndex = 143
        Me.lblCheckSituation.Text = "..."
        Me.lblCheckSituation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblCheckSituation.UseCompatibleTextRendering = true
        '
        'pnlPrgChk
        '
        Me.pnlPrgChk.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrgChk.Controls.Add(Me.pgbCheck)
        Me.pnlPrgChk.Controls.Add(Me.lblPgbCheck)
        Me.pnlPrgChk.Controls.Add(Me.lblCheckSituation)
        Me.pnlPrgChk.Location = New System.Drawing.Point(320, 18)
        Me.pnlPrgChk.Name = "pnlPrgChk"
        Me.pnlPrgChk.Size = New System.Drawing.Size(441, 29)
        Me.pnlPrgChk.TabIndex = 144
        '
        'MainFrm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 12!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 682)
        Me.Controls.Add(Me.pnlPrgChk)
        Me.Controls.Add(Me.pnlHJizenGazoKeisiki)
        Me.Controls.Add(Me.pnlJizenSo)
        Me.Controls.Add(Me.pnlMenuDatacv)
        Me.Controls.Add(Me.pnlMenuMain)
        Me.Controls.Add(Me.lblLine0)
        Me.Controls.Add(Me.pnlRefresh)
        Me.Controls.Add(Me.btnDevTabChange)
        Me.Controls.Add(Me.lblHidden1)
        Me.Controls.Add(Me.PictureBox24)
        Me.Controls.Add(Me.tabCtrlMain)
        Me.Controls.Add(Me.lblTitleKi)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.lblTitleH)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"),System.Drawing.Icon)
        Me.MaximizeBox = false
        Me.Name = "MainFrm"
        Me.Text = "データコンバート"
        Me.pnlMenuDatacv.ResumeLayout(false)
        Me.pnlPrgDCConv.ResumeLayout(false)
        CType(Me.picIcoDCConv2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoDCConv1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgDCRelation.ResumeLayout(false)
        CType(Me.picIcoDCRelation2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoDCRelation1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgDCFileWrite.ResumeLayout(false)
        CType(Me.picIcoDCFileWrite2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoDCFileWrite1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwDCEnd2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwDCJikko2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwDCSelect2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwDCEnd1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwDCJikko1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwDCSelect1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgDCEnd.ResumeLayout(false)
        CType(Me.picIcoDCEnd2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoDCEnd1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgDCJikko.ResumeLayout(false)
        CType(Me.picIcoDCJikko2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoDCJikko1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgDCSelect.ResumeLayout(false)
        CType(Me.picIcoDCSelect2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoDCSelect1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgDCHajimeni.ResumeLayout(false)
        CType(Me.picIcoDCHajimeni2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoDCHajimeni1,System.ComponentModel.ISupportInitialize).EndInit
        Me.tabCtrlMain.ResumeLayout(false)
        Me.tabPageDev.ResumeLayout(false)
        Me.GroupBox6.ResumeLayout(false)
        Me.GroupBox6.PerformLayout
        Me.grpDevSettingX.ResumeLayout(false)
        Me.grpDevSettingX.PerformLayout
        CType(Me.PictureBox1,System.ComponentModel.ISupportInitialize).EndInit
        Me.grpDevSetting1.ResumeLayout(false)
        Me.grpDevSetting1.PerformLayout
        Me.grpDevSetting2.ResumeLayout(false)
        Me.grpDevSetting2.PerformLayout
        Me.grpDevSetting3.ResumeLayout(false)
        Me.grpDevSetting3.PerformLayout
        Me.GroupBox3.ResumeLayout(false)
        Me.GroupBox2.ResumeLayout(false)
        Me.GroupBox2.PerformLayout
        Me.GroupBox1.ResumeLayout(false)
        Me.GroupBox1.PerformLayout
        Me.grptaihi.ResumeLayout(false)
        Me.grptaihi.PerformLayout
        Me.tabPageStart.ResumeLayout(false)
        Me.grpKFirstNaiyo.ResumeLayout(false)
        Me.grpHFirstNaiyo.ResumeLayout(false)
        CType(Me.PictureBox2,System.ComponentModel.ISupportInitialize).EndInit
        Me.tabPageSession.ResumeLayout(false)
        Me.grpTimeOut.ResumeLayout(false)
        Me.grpTimeOut.PerformLayout
        Me.grp10ConnectInfo.ResumeLayout(false)
        Me.grp10ConnectInfo.PerformLayout
        Me.grpV10Authent.ResumeLayout(false)
        Me.grpV10Authent.PerformLayout
        Me.grpV7ConnectInfo.ResumeLayout(false)
        Me.grpV7ConnectInfo.PerformLayout
        Me.grpV7Authent.ResumeLayout(false)
        Me.grpV7Authent.PerformLayout
        Me.tabPageSyoki.ResumeLayout(false)
        Me.pnlConvertType.ResumeLayout(false)
        Me.pnlCVType.ResumeLayout(false)
        Me.pnlCVType.PerformLayout
        Me.pnlUnyoKaisi.ResumeLayout(false)
        Me.pnlUnyoKaisi.PerformLayout
        Me.pnlOptSelect.ResumeLayout(false)
        Me.pnlOptionSelect.ResumeLayout(false)
        Me.pnlUserName.ResumeLayout(false)
        Me.pnlUserName.PerformLayout
        Me.pnlLogPath.ResumeLayout(false)
        Me.pnlLogPath.PerformLayout
        Me.tabPageMenu.ResumeLayout(false)
        Me.grpMenuHojyo.ResumeLayout(false)
        Me.grpMenuJizen.ResumeLayout(false)
        Me.grpMenuDatacv.ResumeLayout(false)
        Me.grpMenuGazocv.ResumeLayout(false)
        Me.grpMenuJigo.ResumeLayout(false)
        Me.tabPageJizen.ResumeLayout(false)
        Me.pnlJizenListPath.ResumeLayout(false)
        Me.pnlJizenListPath.PerformLayout
        Me.tabCtrlJizen.ResumeLayout(false)
        Me.tabPageJizen1.ResumeLayout(false)
        Me.grpJizen1.ResumeLayout(false)
        Me.grpJizen1.PerformLayout
        Me.pnlJizenHurikae.ResumeLayout(false)
        Me.pnlJizenHurikae.PerformLayout
        Me.pnlJizenAzu.ResumeLayout(false)
        Me.pnlJizenAzu.PerformLayout
        Me.Panel7.ResumeLayout(false)
        Me.pnlJizenKai.ResumeLayout(false)
        Me.pnlJizenKai.PerformLayout
        Me.Panel6.ResumeLayout(false)
        Me.tabPageJizen2.ResumeLayout(false)
        Me.grpJizen2.ResumeLayout(false)
        Me.grpJizen2.PerformLayout
        Me.Panel12.ResumeLayout(false)
        Me.Panel12.PerformLayout
        Me.Panel5.ResumeLayout(false)
        Me.pnlJizenHasseiOw.ResumeLayout(false)
        Me.pnlJizenHasseiOw.PerformLayout
        Me.Panel3.ResumeLayout(false)
        Me.pnlJizenYanuso.ResumeLayout(false)
        Me.pnlJizenYanuso.PerformLayout
        Me.Panel8.ResumeLayout(false)
        Me.pnlJizenBkhourei.ResumeLayout(false)
        Me.pnlJizenBkhourei.PerformLayout
        Me.Panel4.ResumeLayout(false)
        Me.tabPageJizen3.ResumeLayout(false)
        Me.grpJizen3.ResumeLayout(false)
        Me.grpJizen3.PerformLayout
        Me.pnlJizenKagi.ResumeLayout(false)
        Me.pnlJizenKagi.PerformLayout
        Me.Panel19.ResumeLayout(false)
        Me.pnlJizenSzenKyshutan.ResumeLayout(false)
        Me.pnlJizenSzenKyshutan.PerformLayout
        Me.Panel11.ResumeLayout(false)
        Me.tabPageJizen4.ResumeLayout(false)
        Me.grpJizen4.ResumeLayout(false)
        Me.grpJizen4.PerformLayout
        Me.Panel15.ResumeLayout(false)
        Me.Panel15.PerformLayout
        Me.Panel20.ResumeLayout(false)
        Me.Panel21.ResumeLayout(false)
        Me.Panel21.PerformLayout
        Me.Panel27.ResumeLayout(false)
        Me.tabPageJizen5.ResumeLayout(false)
        Me.GroupBox4.ResumeLayout(false)
        Me.GroupBox4.PerformLayout
        Me.Panel13.ResumeLayout(false)
        Me.Panel13.PerformLayout
        Me.Panel14.ResumeLayout(false)
        Me.tabPageHJizen1.ResumeLayout(false)
        Me.GroupBox5.ResumeLayout(false)
        Me.GroupBox5.PerformLayout
        Me.pnlHJizenTyukan.ResumeLayout(false)
        Me.pnlHJizenTyukan.PerformLayout
        Me.Panel26.ResumeLayout(false)
        Me.Panel26.PerformLayout
        Me.tabPageJigo.ResumeLayout(false)
        Me.pnlJigoListPath.ResumeLayout(false)
        Me.pnlJigoListPath.PerformLayout
        Me.tabCtrlJigo.ResumeLayout(false)
        Me.tabPageJigo1.ResumeLayout(false)
        Me.grpJigoUserSagyo.ResumeLayout(false)
        Me.grpJigoUserSagyo.PerformLayout
        Me.pnlKJigoCmtSyudo.ResumeLayout(false)
        Me.pnlHJigoCmtSyudo.ResumeLayout(false)
        Me.pnlJigoCmtSyusi.ResumeLayout(false)
        Me.pnlJizenMinus.ResumeLayout(false)
        Me.pnlJizenMinus.PerformLayout
        Me.Panel2.ResumeLayout(false)
        Me.pnlJigoCmtCsvSyuturyoku.ResumeLayout(false)
        Me.pnlJigoCmtSyosiki.ResumeLayout(false)
        Me.tabPageJigo2.ResumeLayout(false)
        Me.grpJigoCmtTyuui.ResumeLayout(false)
        Me.grpJigoCmtTyuui.PerformLayout
        Me.pnlJigoCmtHeiko.ResumeLayout(false)
        Me.grpJigoDonyuji.ResumeLayout(false)
        Me.pnlJigoCmtSoKotiku.ResumeLayout(false)
        Me.pnlJigoCmtNkNyuryoku.ResumeLayout(false)
        Me.pnlJigoCmtSqKotiku.ResumeLayout(false)
        Me.tabPageHojyo.ResumeLayout(false)
        Me.pnlKensyoListPath.ResumeLayout(false)
        Me.pnlKensyoListPath.PerformLayout
        Me.tabCtrlHojyo.ResumeLayout(false)
        Me.tabPageHojyo1.ResumeLayout(false)
        Me.GroupBox27.ResumeLayout(false)
        Me.GroupBox27.PerformLayout
        Me.pnlKiOpKys.ResumeLayout(false)
        Me.pnlKiOpKys.PerformLayout
        Me.Panel18.ResumeLayout(false)
        Me.pnlKiOpOw.ResumeLayout(false)
        Me.pnlKiOpOw.PerformLayout
        Me.Panel16.ResumeLayout(false)
        Me.pnlRelRename.ResumeLayout(false)
        Me.pnlRelRename.PerformLayout
        Me.Panel25.ResumeLayout(false)
        Me.Panel25.PerformLayout
        Me.tabPageHojyo2.ResumeLayout(false)
        Me.GroupBox25.ResumeLayout(false)
        Me.GroupBox25.PerformLayout
        Me.pnlRelationSet.ResumeLayout(false)
        Me.pnlRelationSet.PerformLayout
        Me.Panel17.ResumeLayout(false)
        Me.pnlSzenKysSorit.ResumeLayout(false)
        Me.pnlSzenKysSorit.PerformLayout
        Me.Panel1.ResumeLayout(false)
        Me.pnlKojyosh.ResumeLayout(false)
        Me.pnlKojyosh.PerformLayout
        Me.Panel10.ResumeLayout(false)
        Me.pnlBunkatumisyu.ResumeLayout(false)
        Me.pnlBunkatumisyu.PerformLayout
        Me.Panel9.ResumeLayout(false)
        Me.tabPageHajimeni.ResumeLayout(false)
        Me.GroupBox10.ResumeLayout(false)
        Me.pnlDcFstCmtH99.ResumeLayout(false)
        Me.pnlDcFstCmtK99.ResumeLayout(false)
        Me.pnlDcFstCmtH01.ResumeLayout(false)
        Me.pnlDcFstCmtH02.ResumeLayout(false)
        Me.pnlDcFstCmt11.ResumeLayout(false)
        Me.pnlDcFstCmtK01.ResumeLayout(false)
        Me.pnlDcFstCmt06.ResumeLayout(false)
        Me.pnlDcFstCmt05.ResumeLayout(false)
        Me.pnlDcFstCmt07.ResumeLayout(false)
        Me.pnlDcFstCmt10.ResumeLayout(false)
        Me.pnlDcFstCmt02.ResumeLayout(false)
        Me.pnlDcFstCmt08.ResumeLayout(false)
        Me.pnlDcFstCmt04.ResumeLayout(false)
        Me.pnlDcFstCmt03.ResumeLayout(false)
        Me.pnlDcFstCmt01.ResumeLayout(false)
        Me.pnlDcFstCmtK02.ResumeLayout(false)
        Me.pnlDcFstCmt09.ResumeLayout(false)
        CType(Me.PictureBox14,System.ComponentModel.ISupportInitialize).EndInit
        Me.tabPageSelect.ResumeLayout(false)
        Me.grpMiddleFile.ResumeLayout(false)
        Me.grpMiddleFile.PerformLayout
        Me.tabCtrlCVItem.ResumeLayout(false)
        Me.tabPageKizon110.ResumeLayout(false)
        Me.tabPageKizon110.PerformLayout
        Me.grpKizon1.ResumeLayout(false)
        Me.pnlKiMstHendo.ResumeLayout(false)
        Me.pnlKiMstHendo.PerformLayout
        Me.pnlKiMstTokuyaku.ResumeLayout(false)
        Me.pnlKiMstTokuyaku.PerformLayout
        Me.pnlKiMstKasyoClaimrui.ResumeLayout(false)
        Me.pnlKiMstKasyoClaimrui.PerformLayout
        Me.pnlKiMstTitle.ResumeLayout(false)
        Me.pnlKiMstTitle.PerformLayout
        Me.pnlKiMstArea.ResumeLayout(false)
        Me.pnlKiMstArea.PerformLayout
        Me.pnlKiMstSchool.ResumeLayout(false)
        Me.pnlKiMstSchool.PerformLayout
        Me.pnlKiMstHokenrui.ResumeLayout(false)
        Me.pnlKiMstHokenrui.PerformLayout
        Me.pnlKiMstBus.ResumeLayout(false)
        Me.pnlKiMstBus.PerformLayout
        Me.tabPageKizon120.ResumeLayout(false)
        Me.tabPageKizon120.PerformLayout
        Me.grpKizon3.ResumeLayout(false)
        Me.pnlKiGySyuzenBase.ResumeLayout(false)
        Me.pnlKiGySyuzenBase.PerformLayout
        Me.pnlKiGyYatinhosyoBase.ResumeLayout(false)
        Me.pnlKiGyYatinhosyoBase.PerformLayout
        Me.pnlKiGyLifelineBase.ResumeLayout(false)
        Me.pnlKiGyLifelineBase.PerformLayout
        Me.pnlKiGyHokenBase.ResumeLayout(false)
        Me.pnlKiGyHokenBase.PerformLayout
        Me.pnlKiGySisetuBase.ResumeLayout(false)
        Me.pnlKiGySisetuBase.PerformLayout
        Me.pnlKiGyCyukaiBase.ResumeLayout(false)
        Me.pnlKiGyCyukaiBase.PerformLayout
        Me.pnlKiGySekoBase.ResumeLayout(false)
        Me.pnlKiGySekoBase.PerformLayout
        Me.grpKizon5.ResumeLayout(false)
        Me.pnlKiOw.ResumeLayout(false)
        Me.pnlKiOw.PerformLayout
        Me.pnlKiJisya.ResumeLayout(false)
        Me.pnlKiJisya.PerformLayout
        Me.pnlKiSyskanriBase.ResumeLayout(false)
        Me.pnlKiSyskanriBase.PerformLayout
        Me.tabPageKizon130.ResumeLayout(false)
        Me.tabPageKizon130.PerformLayout
        Me.grpKizon2.ResumeLayout(false)
        Me.pnlRendo.ResumeLayout(false)
        Me.pnlRendo.PerformLayout
        Me.pnlKiSq.ResumeLayout(false)
        Me.pnlKiSq.PerformLayout
        Me.pnlSzen.ResumeLayout(false)
        Me.pnlSzen.PerformLayout
        Me.pnlKiClaim.ResumeLayout(false)
        Me.pnlKiClaim.PerformLayout
        Me.pnlKiKy.ResumeLayout(false)
        Me.pnlKiKy.PerformLayout
        Me.pnlKiKys.ResumeLayout(false)
        Me.pnlKiKys.PerformLayout
        Me.pnlKiHy.ResumeLayout(false)
        Me.pnlKiHy.PerformLayout
        Me.pnlKiBk.ResumeLayout(false)
        Me.pnlKiBk.PerformLayout
        Me.grpKizonKagi.ResumeLayout(false)
        Me.grpKizonKagi.PerformLayout
        Me.tabPageBase110.ResumeLayout(false)
        Me.grpMst.ResumeLayout(false)
        Me.grpMst.PerformLayout
        Me.tabPageBase120.ResumeLayout(false)
        Me.grpGy.ResumeLayout(false)
        Me.grpGy.PerformLayout
        Me.tabPageBase130.ResumeLayout(false)
        Me.grpKys.ResumeLayout(false)
        Me.grpKys.PerformLayout
        Me.grpOw.ResumeLayout(false)
        Me.grpOw.PerformLayout
        Me.grpJisya.ResumeLayout(false)
        Me.grpJisya.PerformLayout
        Me.tabPageBase140.ResumeLayout(false)
        Me.grpBk.ResumeLayout(false)
        Me.grpBk.PerformLayout
        Me.grpKagiSelect.ResumeLayout(false)
        Me.grpKagiSelect.PerformLayout
        Me.tabPageBase150.ResumeLayout(false)
        Me.grpHy.ResumeLayout(false)
        Me.grpHy.PerformLayout
        Me.tabPageBase160.ResumeLayout(false)
        Me.grpSorule.ResumeLayout(false)
        Me.grpSorule.PerformLayout
        Me.tabPageBase170.ResumeLayout(false)
        Me.grpKy.ResumeLayout(false)
        Me.grpKy.PerformLayout
        Me.tabPageBase180.ResumeLayout(false)
        Me.grpSq.ResumeLayout(false)
        Me.grpSq.PerformLayout
        Me.tabPageBase190.ResumeLayout(false)
        Me.grpSzen.ResumeLayout(false)
        Me.grpSzen.PerformLayout
        Me.grpClaim.ResumeLayout(false)
        Me.grpClaim.PerformLayout
        Me.tabPageBase200.ResumeLayout(false)
        Me.grpSyskanri.ResumeLayout(false)
        Me.grpSyskanri.PerformLayout
        Me.tabPageBase210.ResumeLayout(false)
        Me.grpRendo.ResumeLayout(false)
        Me.grpRendo.PerformLayout
        Me.tabPageBase900.ResumeLayout(false)
        Me.tabPageBase900.PerformLayout
        Me.tabPageHanyo110.ResumeLayout(false)
        Me.grpHMst.ResumeLayout(false)
        Me.grpHMst.PerformLayout
        Me.tabPageHanyo120.ResumeLayout(false)
        Me.grpHGy.ResumeLayout(false)
        Me.grpHGy.PerformLayout
        Me.grpHOw.ResumeLayout(false)
        Me.grpHOw.PerformLayout
        Me.grpHJisya.ResumeLayout(false)
        Me.grpHJisya.PerformLayout
        Me.TabPage1.ResumeLayout(false)
        Me.grpKizon4.ResumeLayout(false)
        Me.grpKizon4.PerformLayout
        Me.pnlRekiClear.ResumeLayout(false)
        Me.tabPageJikko.ResumeLayout(false)
        Me.grpTotalProcess.ResumeLayout(false)
        Me.grpTotalProcess.PerformLayout
        Me.tabPageEndOK.ResumeLayout(false)
        CType(Me.PictureBox15,System.ComponentModel.ISupportInitialize).EndInit
        Me.tabPageEndError.ResumeLayout(false)
        Me.GroupBox23.ResumeLayout(false)
        Me.GroupBox19.ResumeLayout(false)
        Me.GroupBox20.ResumeLayout(false)
        CType(Me.PictureBox23,System.ComponentModel.ISupportInitialize).EndInit
        Me.tabPageEndCancel.ResumeLayout(false)
        Me.GroupBox21.ResumeLayout(false)
        Me.GroupBox22.ResumeLayout(false)
        CType(Me.PictureBox25,System.ComponentModel.ISupportInitialize).EndInit
        Me.tabPageIkkatu.ResumeLayout(false)
        Me.tabPageIkkatu.PerformLayout
        Me.tabPageHanyoJizen.ResumeLayout(false)
        Me.GroupBox11.ResumeLayout(false)
        Me.pnlHJizenGazoKeisiki.ResumeLayout(false)
        Me.pnlHJizenGazoKeisiki.PerformLayout
        Me.pnlJizenSo.ResumeLayout(false)
        Me.pnlJizenSo.PerformLayout
        Me.pnlMenuMain.ResumeLayout(false)
        Me.pnlPrgMStart.ResumeLayout(false)
        Me.pnlPrgMStart.PerformLayout
        CType(Me.picIcoMStart2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoMStart1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.PictureBox3,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMSession2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMSession1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgMEnd.ResumeLayout(false)
        Me.pnlPrgMEnd.PerformLayout
        CType(Me.picIcoMEnd2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoMEnd1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoJizen1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgMMenu.ResumeLayout(false)
        Me.pnlPrgMMenu.PerformLayout
        CType(Me.picIcoMMenu2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoMMenu1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoSession1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgMSyoki.ResumeLayout(false)
        Me.pnlPrgMSyoki.PerformLayout
        CType(Me.picIcoMSyoki2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoMSyoki1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoSyoki1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgMSession.ResumeLayout(false)
        Me.pnlPrgMSession.PerformLayout
        CType(Me.picIcoMSession2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picIcoMSession1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.PictureBox4,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMEnd2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMMenu2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMSyoki2,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMEnd1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMMenu1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.picArwMSyoki1,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlRefresh.ResumeLayout(false)
        CType(Me.PictureBox24,System.ComponentModel.ISupportInitialize).EndInit
        Me.pnlPrgChk.ResumeLayout(false)
        Me.ResumeLayout(false)

End Sub
        Friend WithEvents btnConnectTest As System.Windows.Forms.Button
        Friend WithEvents btnAllChk As System.Windows.Forms.Button
        Friend WithEvents chkRelationFile As System.Windows.Forms.CheckBox
        Friend WithEvents Label9 As System.Windows.Forms.Label
        Friend WithEvents btnMidFileCheck As System.Windows.Forms.Button
        Friend WithEvents btnDefConInfoRead As System.Windows.Forms.Button
        Friend WithEvents lblTitleKi As System.Windows.Forms.Label
        Friend WithEvents pnlMenuDatacv As System.Windows.Forms.Panel
        Friend WithEvents picArwDCEnd2 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwDCJikko2 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwDCSelect2 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwDCEnd1 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwDCJikko1 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwDCSelect1 As System.Windows.Forms.PictureBox
        Friend WithEvents pnlPrgDCEnd As System.Windows.Forms.Panel
        Friend WithEvents picIcoDCEnd2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoDCEnd1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgDCEnd As System.Windows.Forms.Label
        Friend WithEvents pnlPrgDCJikko As System.Windows.Forms.Panel
        Friend WithEvents picIcoDCJikko2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoDCJikko1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgDCJikko As System.Windows.Forms.Label
        Friend WithEvents pnlPrgDCSelect As System.Windows.Forms.Panel
        Friend WithEvents picIcoDCSelect2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoDCSelect1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgDCSelect As System.Windows.Forms.Label
        Friend WithEvents pnlPrgDCHajimeni As System.Windows.Forms.Panel
        Friend WithEvents picIcoDCHajimeni2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoDCHajimeni1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgDCHajimeni As System.Windows.Forms.Label
        Friend WithEvents btnBack As System.Windows.Forms.Button
        Friend WithEvents btnEnd As System.Windows.Forms.Button
        Private WithEvents btnNext As System.Windows.Forms.Button
        Friend WithEvents lblHidden1 As System.Windows.Forms.Label
        Friend WithEvents PictureBox24 As System.Windows.Forms.PictureBox
        Friend WithEvents tabCtrlMain As System.Windows.Forms.TabControl
        Friend WithEvents tabPageHajimeni As System.Windows.Forms.TabPage
        Friend WithEvents PictureBox14 As System.Windows.Forms.PictureBox
        Friend WithEvents lblDatacvHajimeniDescription1 As System.Windows.Forms.Label
        Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
        Friend WithEvents Label126 As System.Windows.Forms.Label
        Friend WithEvents Label132 As System.Windows.Forms.Label
        Friend WithEvents Label122 As System.Windows.Forms.Label
        Friend WithEvents Label104 As System.Windows.Forms.Label
        Friend WithEvents Label123 As System.Windows.Forms.Label
        Friend WithEvents Label105 As System.Windows.Forms.Label
        Friend WithEvents Label127 As System.Windows.Forms.Label
        Friend WithEvents Label129 As System.Windows.Forms.Label
        Friend WithEvents Label133 As System.Windows.Forms.Label
        Friend WithEvents Label135 As System.Windows.Forms.Label
        Friend WithEvents lblKDatacvHajimeniLabel As System.Windows.Forms.Label
        Friend WithEvents tabPageSyoki As System.Windows.Forms.TabPage
        Friend WithEvents txtRecUser As System.Windows.Forms.TextBox
        Friend WithEvents lblExecutor As System.Windows.Forms.Label
        Friend WithEvents btnLogDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtLogDirPath As System.Windows.Forms.TextBox
        Friend WithEvents lblLogDir As System.Windows.Forms.Label
        Friend WithEvents lblSyokiDescription1 As System.Windows.Forms.Label
        Friend WithEvents tabPageSession As System.Windows.Forms.TabPage
        Friend WithEvents lblSessionCaution As System.Windows.Forms.Label
        Friend WithEvents lblTimeOutSec As System.Windows.Forms.Label
        Friend WithEvents txtTimeOut As System.Windows.Forms.TextBox
        Friend WithEvents Label36 As System.Windows.Forms.Label
        Friend WithEvents Label35 As System.Windows.Forms.Label
        Friend WithEvents lblNetworklib As System.Windows.Forms.Label
        Friend WithEvents Label30 As System.Windows.Forms.Label
        Friend WithEvents Label13 As System.Windows.Forms.Label
        Friend WithEvents grp10ConnectInfo As System.Windows.Forms.GroupBox
        Friend WithEvents grpV10Authent As System.Windows.Forms.GroupBox
        Friend WithEvents optV10Authent2 As System.Windows.Forms.RadioButton
        Friend WithEvents optV10Authent1 As System.Windows.Forms.RadioButton
        Friend WithEvents txtV10Pass As System.Windows.Forms.TextBox
        Friend WithEvents txtV10User As System.Windows.Forms.TextBox
        Friend WithEvents txtV10Catalog As System.Windows.Forms.TextBox
        Friend WithEvents txtV10Server As System.Windows.Forms.TextBox
        Friend WithEvents lblV10 As System.Windows.Forms.Label
        Friend WithEvents grpV7ConnectInfo As System.Windows.Forms.GroupBox
        Friend WithEvents grpV7Authent As System.Windows.Forms.GroupBox
        Friend WithEvents optV7Authent2 As System.Windows.Forms.RadioButton
        Friend WithEvents optV7Authent1 As System.Windows.Forms.RadioButton
        Friend WithEvents txtV7Pass As System.Windows.Forms.TextBox
        Friend WithEvents txtV7User As System.Windows.Forms.TextBox
        Friend WithEvents txtV7Catalog As System.Windows.Forms.TextBox
        Friend WithEvents txtV7Server As System.Windows.Forms.TextBox
        Friend WithEvents lblV7 As System.Windows.Forms.Label
        Friend WithEvents tabPageJizen As System.Windows.Forms.TabPage
        Friend WithEvents tabCtrlJizen As System.Windows.Forms.TabControl
        Friend WithEvents tabPageJizen1 As System.Windows.Forms.TabPage
        Friend WithEvents grpJizen1 As System.Windows.Forms.GroupBox
        Friend WithEvents Label90 As System.Windows.Forms.Label
        Friend WithEvents Label88 As System.Windows.Forms.Label
        Friend WithEvents Label85 As System.Windows.Forms.Label
        Friend WithEvents Label84 As System.Windows.Forms.Label
        Friend WithEvents tabPageJizen2 As System.Windows.Forms.TabPage
        Friend WithEvents grpJizen2 As System.Windows.Forms.GroupBox
        Friend WithEvents Label94 As System.Windows.Forms.Label
        Friend WithEvents Label100 As System.Windows.Forms.Label
        Friend WithEvents Label102 As System.Windows.Forms.Label
        Friend WithEvents lblCautionDescription As System.Windows.Forms.Label
        Friend WithEvents lblJizenDescription1 As System.Windows.Forms.Label
        Friend WithEvents tabPageSelect As System.Windows.Forms.TabPage
        Friend WithEvents btnRekiClear As System.Windows.Forms.Button
        Friend WithEvents grpMiddleFile As System.Windows.Forms.GroupBox
        Friend WithEvents btnMidDirLogSeach As System.Windows.Forms.Button
        Friend WithEvents txtMidDirLogPath As System.Windows.Forms.TextBox
        Friend WithEvents lblMiddleFileLog As System.Windows.Forms.Label
        Friend WithEvents btnMidDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtMidDirPath As System.Windows.Forms.TextBox
        Friend WithEvents lblMiddleFile As System.Windows.Forms.Label
        Friend WithEvents lblDatacvSelectCaution As System.Windows.Forms.Label
        Friend WithEvents lblDatacvSelectDescription1 As System.Windows.Forms.Label
        Friend WithEvents tabCtrlCVItem As System.Windows.Forms.TabControl
        Friend WithEvents tabPageKizon110 As System.Windows.Forms.TabPage
        Friend WithEvents grpKizon1 As System.Windows.Forms.GroupBox
        Friend WithEvents chkKiMstHendo As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiMstKasyoClaimrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiMstBus As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiMstSchool As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiMstTokuyaku As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiMstHokenrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiMstTitle As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiMstArea As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageKizon120 As System.Windows.Forms.TabPage
        Friend WithEvents tabPageKizon130 As System.Windows.Forms.TabPage
        Friend WithEvents tabPageBase110 As System.Windows.Forms.TabPage
        Friend WithEvents lblRelItemInfo As System.Windows.Forms.Label
        Friend WithEvents grpMst As System.Windows.Forms.GroupBox
        Friend WithEvents chkMstHendoitiran As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstHendo As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstKasyoClaimrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstBus As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstSchool As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstTokuyaku As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstHokenrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstBusKotu As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstKagititle As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstArea As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase120 As System.Windows.Forms.TabPage
        Friend WithEvents grpGy As System.Windows.Forms.GroupBox
        Friend WithEvents chkGySyuzenMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyHokenMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkGySyuzenKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkGySekoBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkGySisetuBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkGySyuzenBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyYatinhosyoMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyYatinhosyoBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyCyukaiKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyCyukaiMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyHokenBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyLifelineBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyHokenKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkGyCyukaiBase As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase130 As System.Windows.Forms.TabPage
        Friend WithEvents grpKys As System.Windows.Forms.GroupBox
        Friend WithEvents chkKysHosyonin As System.Windows.Forms.CheckBox
        Friend WithEvents chkKysSyogoKana As System.Windows.Forms.CheckBox
        Friend WithEvents chkKysKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkKysMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkKysBase As System.Windows.Forms.CheckBox
        Friend WithEvents grpOw As System.Windows.Forms.GroupBox
        Friend WithEvents chkOwEvent As System.Windows.Forms.CheckBox
        Friend WithEvents chkOwKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkOwMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkOwBase As System.Windows.Forms.CheckBox
        Friend WithEvents grpJisya As System.Windows.Forms.GroupBox
        Friend WithEvents chkFBKozafurikae As System.Windows.Forms.CheckBox
        Friend WithEvents chkFBFuriirai As System.Windows.Forms.CheckBox
        Friend WithEvents chkJisyaKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkJisyaMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkJisyaBase As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase140 As System.Windows.Forms.TabPage
        Friend WithEvents grpBk As System.Windows.Forms.GroupBox
        Friend WithEvents chkBkSyo As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkHendo As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkKinrincyusyajo As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkSansyofile As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkSzeniji As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkSyuhen As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkSetudo As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkKotu As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkKenri As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkGomi As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkSyosai As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkKagi As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkBkBase As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase150 As System.Windows.Forms.TabPage
        Friend WithEvents grpHy As System.Windows.Forms.GroupBox
        Friend WithEvents chkHySyo As System.Windows.Forms.CheckBox
        Friend WithEvents chkHySansyofile As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyGenjotanka As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyKenri As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyConfirm As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyHendo As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyCommonsalespoint As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyMenseki As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyNkinkomk As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyMadoriutiwake As System.Windows.Forms.CheckBox
        Friend WithEvents chkHySzeniji As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyTokuyaku As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyParking As System.Windows.Forms.CheckBox
        Friend WithEvents chkHySyosai As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyKagi As System.Windows.Forms.CheckBox
        Friend WithEvents chkHySetubi As System.Windows.Forms.CheckBox
        Friend WithEvents chkHyBase As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase160 As System.Windows.Forms.TabPage
        Friend WithEvents grpSorule As System.Windows.Forms.GroupBox
        Friend WithEvents chkSoruleSosaki As System.Windows.Forms.CheckBox
        Friend WithEvents chkSoruleKojo As System.Windows.Forms.CheckBox
        Friend WithEvents chkSoruleNkin As System.Windows.Forms.CheckBox
        Friend WithEvents chkSoruleBase As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase170 As System.Windows.Forms.TabPage
        Friend WithEvents grpKy As System.Windows.Forms.GroupBox
        Friend WithEvents chkKyKai As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox14 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox24 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox23 As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyHosyonin As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyTokuyaku As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox18 As System.Windows.Forms.CheckBox
        Friend WithEvents chkKySorule As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox16 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox15 As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyNyukyo As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyNkinkomkNx As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyNkinkomk As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyHoken As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyKojoRule As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyKys As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyRireki As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyHendo As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyCar As System.Windows.Forms.CheckBox
        Friend WithEvents chkKyBase As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase180 As System.Windows.Forms.TabPage
        Friend WithEvents grpSq As System.Windows.Forms.GroupBox
        Friend WithEvents chkSqSqKojo As System.Windows.Forms.CheckBox
        Friend WithEvents chkSqKoteiKojo As System.Windows.Forms.CheckBox
        Friend WithEvents chkSqHendokensin As System.Windows.Forms.CheckBox
        Friend WithEvents chkSqUnyotaino As System.Windows.Forms.CheckBox
        Friend WithEvents chkSqSq As System.Windows.Forms.CheckBox
        Friend WithEvents chkSqKajyo As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase900 As System.Windows.Forms.TabPage
        Friend WithEvents tabPageHanyo110 As System.Windows.Forms.TabPage
        Friend WithEvents grpHMst As System.Windows.Forms.GroupBox
        Friend WithEvents Label138 As System.Windows.Forms.Label
        Friend WithEvents Label150 As System.Windows.Forms.Label
        Friend WithEvents Label151 As System.Windows.Forms.Label
        Friend WithEvents Label152 As System.Windows.Forms.Label
        Friend WithEvents Label153 As System.Windows.Forms.Label
        Friend WithEvents Label154 As System.Windows.Forms.Label
        Friend WithEvents Label155 As System.Windows.Forms.Label
        Friend WithEvents Label156 As System.Windows.Forms.Label
        Friend WithEvents Label157 As System.Windows.Forms.Label
        Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox9 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox10 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox11 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox12 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox13 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox17 As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageHanyo120 As System.Windows.Forms.TabPage
        Friend WithEvents grpHGy As System.Windows.Forms.GroupBox
        Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox30 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox31 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox32 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox33 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox34 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox35 As System.Windows.Forms.CheckBox
        Friend WithEvents grpHOw As System.Windows.Forms.GroupBox
        Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox26 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox27 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox28 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox29 As System.Windows.Forms.CheckBox
        Friend WithEvents grpHJisya As System.Windows.Forms.GroupBox
        Friend WithEvents CheckBox19 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox20 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox21 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox22 As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox25 As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageHanyo130 As System.Windows.Forms.TabPage
        Friend WithEvents tabPageHanyo140 As System.Windows.Forms.TabPage
        Friend WithEvents tabPageHanyo150 As System.Windows.Forms.TabPage
        Friend WithEvents tabPageJikko As System.Windows.Forms.TabPage
        Friend WithEvents lblDatacvJikkoCaution As System.Windows.Forms.Label
        Friend WithEvents lblDatacvJikkoDescription1 As System.Windows.Forms.Label
        Friend WithEvents lblPgbPartial As System.Windows.Forms.Label
        Friend WithEvents txtPartialSituation As System.Windows.Forms.TextBox
        Friend WithEvents pgbpartial As System.Windows.Forms.ProgressBar
        Friend WithEvents lblCVItem As System.Windows.Forms.Label
        Friend WithEvents Label31 As System.Windows.Forms.Label
        Friend WithEvents grpTotalProcess As System.Windows.Forms.GroupBox
        Friend WithEvents lblPgbTotal As System.Windows.Forms.Label
        Friend WithEvents txtTotalSituation As System.Windows.Forms.TextBox
        Friend WithEvents pgbTotal As System.Windows.Forms.ProgressBar
        Friend WithEvents Label32 As System.Windows.Forms.Label
        Friend WithEvents lblTotalSituation As System.Windows.Forms.Label
        Friend WithEvents tabPageEndOK As System.Windows.Forms.TabPage
        Friend WithEvents PictureBox15 As System.Windows.Forms.PictureBox
        Friend WithEvents lblDatacvEndDescription1 As System.Windows.Forms.Label
        Friend WithEvents lblKDatacvEndLabel As System.Windows.Forms.Label
        Friend WithEvents tabPageDev As System.Windows.Forms.TabPage
        Friend WithEvents grpDevSetting1 As System.Windows.Forms.GroupBox
        Friend WithEvents Label51 As System.Windows.Forms.Label
        Friend WithEvents chkKyNewest As System.Windows.Forms.CheckBox
        Friend WithEvents Label48 As System.Windows.Forms.Label
        Friend WithEvents chkChildItemControl As System.Windows.Forms.CheckBox
        Friend WithEvents Label46 As System.Windows.Forms.Label
        Friend WithEvents Label45 As System.Windows.Forms.Label
        Friend WithEvents Label44 As System.Windows.Forms.Label
        Friend WithEvents chkRelTblDrop As System.Windows.Forms.CheckBox
        Friend WithEvents Label43 As System.Windows.Forms.Label
        Friend WithEvents Label42 As System.Windows.Forms.Label
        Friend WithEvents chkV7ViewDrop As System.Windows.Forms.CheckBox
        Friend WithEvents chkLogTblDrop As System.Windows.Forms.CheckBox
        Friend WithEvents chkOverWrite As System.Windows.Forms.CheckBox
        Friend WithEvents Label52 As System.Windows.Forms.Label
        Friend WithEvents grpDevSetting2 As System.Windows.Forms.GroupBox
        Friend WithEvents btnRelDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtRelationDirPath As System.Windows.Forms.TextBox
        Friend WithEvents lblRelationDir As System.Windows.Forms.Label
        Friend WithEvents Label11 As System.Windows.Forms.Label
        Friend WithEvents chkRelation As System.Windows.Forms.CheckBox
        Friend WithEvents Label47 As System.Windows.Forms.Label
        Friend WithEvents chkMidNotStop As System.Windows.Forms.CheckBox
        Friend WithEvents chkCVStart As System.Windows.Forms.CheckBox
        Friend WithEvents Label8 As System.Windows.Forms.Label
        Friend WithEvents Label10 As System.Windows.Forms.Label
        Friend WithEvents chkMiddleFile As System.Windows.Forms.CheckBox
        Friend WithEvents grpDevSetting3 As System.Windows.Forms.GroupBox
        Friend WithEvents Label55 As System.Windows.Forms.Label
        Friend WithEvents Label54 As System.Windows.Forms.Label
        Friend WithEvents chkRommDuplicate As System.Windows.Forms.CheckBox
        Friend WithEvents chkEmptyRoomNo As System.Windows.Forms.CheckBox
        Friend WithEvents Label53 As System.Windows.Forms.Label
        Friend WithEvents Label50 As System.Windows.Forms.Label
        Friend WithEvents Label49 As System.Windows.Forms.Label
        Friend WithEvents Label41 As System.Windows.Forms.Label
        Friend WithEvents txtLogOutputCnt As System.Windows.Forms.TextBox
        Friend WithEvents grpRelation As System.Windows.Forms.GroupBox
        Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
        Friend WithEvents Label39 As System.Windows.Forms.Label
        Friend WithEvents Label33 As System.Windows.Forms.Label
        Friend WithEvents Label6 As System.Windows.Forms.Label
        Friend WithEvents Label56 As System.Windows.Forms.Label
        Friend WithEvents Label16 As System.Windows.Forms.Label
        Friend WithEvents Label15 As System.Windows.Forms.Label
        Friend WithEvents Label14 As System.Windows.Forms.Label
        Friend WithEvents Label12 As System.Windows.Forms.Label
        Friend WithEvents Label18 As System.Windows.Forms.Label
        Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
        Friend WithEvents chkMstKasyorui As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstGenjotokuyaku As System.Windows.Forms.CheckBox
        Friend WithEvents lblTokuyakuInfo As System.Windows.Forms.Label
        Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
        Friend WithEvents Label34 As System.Windows.Forms.Label
        Friend WithEvents Label29 As System.Windows.Forms.Label
        Friend WithEvents chkMstKozasyubetu As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstSetubi As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstKozo As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstTorihikitaiyo As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstNkinkbn As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstHyrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstNkinkomok As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstBkrui As System.Windows.Forms.CheckBox
        Friend WithEvents grptaihi As System.Windows.Forms.GroupBox
        Friend WithEvents Label7 As System.Windows.Forms.Label
        Friend WithEvents btnFileSeach As System.Windows.Forms.Button
        Friend WithEvents Label5 As System.Windows.Forms.Label
        Friend WithEvents chkGyYatinhosyoKoza As System.Windows.Forms.CheckBox
        Friend WithEvents CheckBox66 As System.Windows.Forms.CheckBox
        Friend WithEvents lblGy As System.Windows.Forms.Label
        Friend WithEvents lblMst As System.Windows.Forms.Label
        Friend WithEvents lblKys As System.Windows.Forms.Label
        Friend WithEvents lblOwner As System.Windows.Forms.Label
        Friend WithEvents lblJisya As System.Windows.Forms.Label
        Friend WithEvents Label17 As System.Windows.Forms.Label
        Friend WithEvents Label19 As System.Windows.Forms.Label
        Friend WithEvents Label20 As System.Windows.Forms.Label
        Friend WithEvents chkMs25 As System.Windows.Forms.CheckBox
        Friend WithEvents chkMsKeiyakusyubetu As System.Windows.Forms.CheckBox
        Friend WithEvents chkMs23 As System.Windows.Forms.CheckBox
        Friend WithEvents chkMs24 As System.Windows.Forms.CheckBox
        Friend WithEvents chkMsOwevent As System.Windows.Forms.CheckBox
        Friend WithEvents chkMsYane As System.Windows.Forms.CheckBox
        Friend WithEvents chkMsKeikaikakunin As System.Windows.Forms.CheckBox
        Friend WithEvents chkMsKinyu As System.Windows.Forms.CheckBox
        Friend WithEvents chkMsKinyuten As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageIkkatu As System.Windows.Forms.TabPage
        Friend WithEvents Label114 As System.Windows.Forms.Label
        Friend WithEvents Label144 As System.Windows.Forms.Label
        Friend WithEvents tabPageHanyoJizen As System.Windows.Forms.TabPage
        Friend WithEvents Label139 As System.Windows.Forms.Label
        Friend WithEvents Label145 As System.Windows.Forms.Label
        Friend WithEvents GroupBox11 As System.Windows.Forms.GroupBox
        Friend WithEvents Label140 As System.Windows.Forms.Label
        Friend WithEvents Label141 As System.Windows.Forms.Label
        Friend WithEvents Label128 As System.Windows.Forms.Label
        Friend WithEvents Label134 As System.Windows.Forms.Label
        Friend WithEvents tabPageEndError As System.Windows.Forms.TabPage
        Friend WithEvents GroupBox23 As System.Windows.Forms.GroupBox
        Friend WithEvents Label163 As System.Windows.Forms.Label
        Friend WithEvents PictureBox23 As System.Windows.Forms.PictureBox
        Friend WithEvents lblDatacvErrorDescription1 As System.Windows.Forms.Label
        Friend WithEvents GroupBox19 As System.Windows.Forms.GroupBox
        Friend WithEvents Label170 As System.Windows.Forms.Label
        Friend WithEvents GroupBox20 As System.Windows.Forms.GroupBox
        Friend WithEvents Label176 As System.Windows.Forms.Label
        Friend WithEvents lblDatacvErrorLabel As System.Windows.Forms.Label
        Friend WithEvents tabPageEndCancel As System.Windows.Forms.TabPage
        Friend WithEvents PictureBox25 As System.Windows.Forms.PictureBox
        Friend WithEvents lblDatacvCancelDescription1 As System.Windows.Forms.Label
        Friend WithEvents GroupBox21 As System.Windows.Forms.GroupBox
        Friend WithEvents Label186 As System.Windows.Forms.Label
        Friend WithEvents GroupBox22 As System.Windows.Forms.GroupBox
        Friend WithEvents Label192 As System.Windows.Forms.Label
        Friend WithEvents lblDatacvCancelLabel As System.Windows.Forms.Label
        Friend WithEvents grpKagiSelect As System.Windows.Forms.GroupBox
        Friend WithEvents optKyKagi As System.Windows.Forms.RadioButton
        Friend WithEvents optHyKagi As System.Windows.Forms.RadioButton
        Friend WithEvents btnDevTabChange As System.Windows.Forms.Button
        Friend WithEvents Label22 As System.Windows.Forms.Label
        Friend WithEvents lblHidden2 As System.Windows.Forms.Label
        Friend WithEvents lblLine2 As System.Windows.Forms.Label
        Friend WithEvents lblLine1 As System.Windows.Forms.Label
        Friend WithEvents lblHidden3 As System.Windows.Forms.Label
        Friend WithEvents grpTimeOut As System.Windows.Forms.GroupBox
        Friend WithEvents Label119 As System.Windows.Forms.Label
        Friend WithEvents Label116 As System.Windows.Forms.Label
        Friend WithEvents btnDoui As System.Windows.Forms.Button
        Friend WithEvents chkJizenAzu As System.Windows.Forms.CheckBox
        Friend WithEvents chkJizenHurikae As System.Windows.Forms.CheckBox
        Friend WithEvents chkJizenSo As System.Windows.Forms.CheckBox
        Friend WithEvents chkJizenKai As System.Windows.Forms.CheckBox
        Friend WithEvents Label40 As System.Windows.Forms.Label
        Friend WithEvents Label37 As System.Windows.Forms.Label
        Friend WithEvents chkJizenBkhourei As System.Windows.Forms.CheckBox
        Friend WithEvents chkJizenHasseiOw As System.Windows.Forms.CheckBox
        Friend WithEvents chkJizenMinus As System.Windows.Forms.CheckBox
        Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
        Friend WithEvents chkOpOw As System.Windows.Forms.CheckBox
        Friend WithEvents chkOpKys As System.Windows.Forms.CheckBox
        Friend WithEvents lblJizenPageCnt1 As System.Windows.Forms.Label
        Friend WithEvents lblJizenPageCnt2 As System.Windows.Forms.Label
        Friend WithEvents chkMstYatinKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkFBNsSyutoku As System.Windows.Forms.CheckBox
        Friend WithEvents chkFBFuritesuryo As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase190 As System.Windows.Forms.TabPage
        Friend WithEvents grpClaim As System.Windows.Forms.GroupBox
        Friend WithEvents chkClaimTaiorireki As System.Windows.Forms.CheckBox
        Friend WithEvents chkClaimRelfile As System.Windows.Forms.CheckBox
        Friend WithEvents chkClaimBase As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageBase200 As System.Windows.Forms.TabPage
        Friend WithEvents grpSyskanri As System.Windows.Forms.GroupBox
        Friend WithEvents chkSyskanriBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkJisyaTanto As System.Windows.Forms.CheckBox
        Friend WithEvents chkSyskanriNkinkomkmerge As System.Windows.Forms.CheckBox
        Friend WithEvents chkSyskanriHenkanmoji As System.Windows.Forms.CheckBox
        Friend WithEvents chkSyskanriZei As System.Windows.Forms.CheckBox
        Friend WithEvents chkKySzenmeisai As System.Windows.Forms.CheckBox
        Friend WithEvents chkKySzen As System.Windows.Forms.CheckBox
        Friend WithEvents grpSzen As System.Windows.Forms.GroupBox
        Friend WithEvents chkSzenMemo As System.Windows.Forms.CheckBox
        Friend WithEvents chkSzenSzen As System.Windows.Forms.CheckBox
        Friend WithEvents chkSzenSzenmeisai As System.Windows.Forms.CheckBox
        Friend WithEvents chkSzenBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkSzenClaim As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstBikotitle As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstBikolst As System.Windows.Forms.CheckBox
        Friend WithEvents chkJizenHasseiHen As System.Windows.Forms.CheckBox
        Friend WithEvents tabPageJizen3 As System.Windows.Forms.TabPage
        Friend WithEvents Label101 As System.Windows.Forms.Label
        Friend WithEvents Label97 As System.Windows.Forms.Label
        Friend WithEvents txtUnyoYYYYMM As System.Windows.Forms.MaskedTextBox
        Friend WithEvents Label165 As System.Windows.Forms.Label
        Friend WithEvents Panel5 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenHasseiHen As System.Windows.Forms.Button
        Friend WithEvents Label125 As System.Windows.Forms.Label
        Friend WithEvents Label130 As System.Windows.Forms.Label
        Friend WithEvents Label131 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenHasseiHen As System.Windows.Forms.Label
        Friend WithEvents Panel4 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenBkhourei As System.Windows.Forms.Button
        Friend WithEvents Label103 As System.Windows.Forms.Label
        Friend WithEvents Label117 As System.Windows.Forms.Label
        Friend WithEvents Label171 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenBkhoureiNotDef As System.Windows.Forms.Label
        Friend WithEvents Panel3 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenHasseiOw As System.Windows.Forms.Button
        Friend WithEvents Label89 As System.Windows.Forms.Label
        Friend WithEvents Label118 As System.Windows.Forms.Label
        Friend WithEvents Label120 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenHasseiOw As System.Windows.Forms.Label
        Friend WithEvents Panel2 As System.Windows.Forms.Panel
        Friend WithEvents Label93 As System.Windows.Forms.Label
        Friend WithEvents Label182 As System.Windows.Forms.Label
        Friend WithEvents Label183 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenMinus As System.Windows.Forms.Label
        Friend WithEvents btnListJizenMinus As System.Windows.Forms.Button
        Friend WithEvents Panel7 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenAzu As System.Windows.Forms.Button
        Friend WithEvents Label166 As System.Windows.Forms.Label
        Friend WithEvents Label167 As System.Windows.Forms.Label
        Friend WithEvents Label168 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenAzu As System.Windows.Forms.Label
        Friend WithEvents Panel6 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenKai As System.Windows.Forms.Button
        Friend WithEvents Label161 As System.Windows.Forms.Label
        Friend WithEvents Label164 As System.Windows.Forms.Label
        Friend WithEvents Label185 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenKai As System.Windows.Forms.Label
        Friend WithEvents grpJizen3 As System.Windows.Forms.GroupBox
        Friend WithEvents Label224 As System.Windows.Forms.Label
        Friend WithEvents Panel11 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenSzenKyshutan As System.Windows.Forms.Button
        Friend WithEvents Label202 As System.Windows.Forms.Label
        Friend WithEvents Label203 As System.Windows.Forms.Label
        Friend WithEvents Label204 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenSzenKyshutan As System.Windows.Forms.Label
        Friend WithEvents chkJizenSzenKyshutan As System.Windows.Forms.CheckBox
        Friend WithEvents Label173 As System.Windows.Forms.Label
        Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
        Friend WithEvents grpKizon4 As System.Windows.Forms.GroupBox
        Friend WithEvents lblSelectPageCnt3 As System.Windows.Forms.Label
        Friend WithEvents Label26 As System.Windows.Forms.Label
        Friend WithEvents Label73 As System.Windows.Forms.Label
        Friend WithEvents Label70 As System.Windows.Forms.Label
        Friend WithEvents Label71 As System.Windows.Forms.Label
        Friend WithEvents Label72 As System.Windows.Forms.Label
        Friend WithEvents Label74 As System.Windows.Forms.Label
        Friend WithEvents chkHSongai As System.Windows.Forms.CheckBox
        Friend WithEvents chkHJisyaKoza As System.Windows.Forms.CheckBox
        Friend WithEvents chkHSetubi As System.Windows.Forms.CheckBox
        Friend WithEvents chkHDataFmt As System.Windows.Forms.CheckBox
        Friend WithEvents Label69 As System.Windows.Forms.Label
        Friend WithEvents Label68 As System.Windows.Forms.Label
        Friend WithEvents Label67 As System.Windows.Forms.Label
        Friend WithEvents Label66 As System.Windows.Forms.Label
        Friend WithEvents Label65 As System.Windows.Forms.Label
        Friend WithEvents Label64 As System.Windows.Forms.Label
        Friend WithEvents Label63 As System.Windows.Forms.Label
        Friend WithEvents Label62 As System.Windows.Forms.Label
        Friend WithEvents Label61 As System.Windows.Forms.Label
        Friend WithEvents Label60 As System.Windows.Forms.Label
        Friend WithEvents Label59 As System.Windows.Forms.Label
        Friend WithEvents chkHNkinKomk As System.Windows.Forms.CheckBox
        Friend WithEvents chkHTaiyo As System.Windows.Forms.CheckBox
        Friend WithEvents chkHKozo As System.Windows.Forms.CheckBox
        Friend WithEvents chkHKozaSyu As System.Windows.Forms.CheckBox
        Friend WithEvents chkHEki As System.Windows.Forms.CheckBox
        Friend WithEvents chkHEnsen As System.Windows.Forms.CheckBox
        Friend WithEvents chkHKinyuSiten As System.Windows.Forms.CheckBox
        Friend WithEvents chkHYouto As System.Windows.Forms.CheckBox
        Friend WithEvents chkHBkBunrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkHKinyu As System.Windows.Forms.CheckBox
        Friend WithEvents chkHKyBunrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkHKagi As System.Windows.Forms.CheckBox
        Friend WithEvents chkHHouKenri As System.Windows.Forms.CheckBox
        Friend WithEvents chkHNkinKbn As System.Windows.Forms.CheckBox
        Friend WithEvents chkHHyBunrui As System.Windows.Forms.CheckBox
        Friend WithEvents grpKizon2 As System.Windows.Forms.GroupBox
        Friend WithEvents chkKiSzenBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiClaimBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiSqBase As System.Windows.Forms.CheckBox
        Friend WithEvents grpKizonKagi As System.Windows.Forms.GroupBox
        Friend WithEvents optKiKyKagi As System.Windows.Forms.RadioButton
        Friend WithEvents optKiHyKagi As System.Windows.Forms.RadioButton
        Friend WithEvents Label86 As System.Windows.Forms.Label
        Friend WithEvents chkKiKysBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiKyBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiHySetubi As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiHyBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiBkBase As System.Windows.Forms.CheckBox
        Friend WithEvents Label78 As System.Windows.Forms.Label
        Friend WithEvents Label222 As System.Windows.Forms.Label
        Friend WithEvents Label223 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstTokuyakuCnt As System.Windows.Forms.Label
        Friend WithEvents Label218 As System.Windows.Forms.Label
        Friend WithEvents Label219 As System.Windows.Forms.Label
        Friend WithEvents Label220 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstHokenruiCnt As System.Windows.Forms.Label
        Friend WithEvents Label76 As System.Windows.Forms.Label
        Friend WithEvents Label77 As System.Windows.Forms.Label
        Friend WithEvents Label216 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstAreaCnt As System.Windows.Forms.Label
        Friend WithEvents Label75 As System.Windows.Forms.Label
        Friend WithEvents Label213 As System.Windows.Forms.Label
        Friend WithEvents Label214 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstSchoolCnt As System.Windows.Forms.Label
        Friend WithEvents Label210 As System.Windows.Forms.Label
        Friend WithEvents Label181 As System.Windows.Forms.Label
        Friend WithEvents Label211 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstBusCnt As System.Windows.Forms.Label
        Friend WithEvents grpKizon3 As System.Windows.Forms.GroupBox
        Friend WithEvents chkKiGySekoBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiGySisetuBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiGySyuzenBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiGyYatinhosyoBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiGyHokenBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiGyLifelineBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiGyCyukaiBase As System.Windows.Forms.CheckBox
        Friend WithEvents Label226 As System.Windows.Forms.Label
        Friend WithEvents Label227 As System.Windows.Forms.Label
        Friend WithEvents Label228 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstKasyoClaimruiCnt As System.Windows.Forms.Label
        Friend WithEvents Label80 As System.Windows.Forms.Label
        Friend WithEvents Label232 As System.Windows.Forms.Label
        Friend WithEvents Label233 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstKagititleCnt As System.Windows.Forms.Label
        Friend WithEvents Label81 As System.Windows.Forms.Label
        Friend WithEvents Label82 As System.Windows.Forms.Label
        Friend WithEvents Label230 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstHendoCnt As System.Windows.Forms.Label
        Friend WithEvents Label255 As System.Windows.Forms.Label
        Friend WithEvents Label256 As System.Windows.Forms.Label
        Friend WithEvents lblKiGySekoBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label252 As System.Windows.Forms.Label
        Friend WithEvents Label253 As System.Windows.Forms.Label
        Friend WithEvents lblKiGySisetuBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label249 As System.Windows.Forms.Label
        Friend WithEvents Label250 As System.Windows.Forms.Label
        Friend WithEvents lblKiGyYatinhosyoBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label246 As System.Windows.Forms.Label
        Friend WithEvents Label247 As System.Windows.Forms.Label
        Friend WithEvents lblKiGyHokenBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label243 As System.Windows.Forms.Label
        Friend WithEvents Label244 As System.Windows.Forms.Label
        Friend WithEvents lblKiGyLifelineBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label240 As System.Windows.Forms.Label
        Friend WithEvents Label241 As System.Windows.Forms.Label
        Friend WithEvents lblKiGySyuzenBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label87 As System.Windows.Forms.Label
        Friend WithEvents Label238 As System.Windows.Forms.Label
        Friend WithEvents lblKiGyCyukaiBaseCnt As System.Windows.Forms.Label
        Friend WithEvents grpKizon5 As System.Windows.Forms.GroupBox
        Friend WithEvents chkKiSyskanriBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiJisyaBase As System.Windows.Forms.CheckBox
        Friend WithEvents chkKiOwBase As System.Windows.Forms.CheckBox
        Friend WithEvents Label267 As System.Windows.Forms.Label
        Friend WithEvents Label268 As System.Windows.Forms.Label
        Friend WithEvents lblKiOwBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label263 As System.Windows.Forms.Label
        Friend WithEvents Label264 As System.Windows.Forms.Label
        Friend WithEvents Label265 As System.Windows.Forms.Label
        Friend WithEvents lblKiJisyaBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label258 As System.Windows.Forms.Label
        Friend WithEvents Label259 As System.Windows.Forms.Label
        Friend WithEvents Label260 As System.Windows.Forms.Label
        Friend WithEvents lblKiSyskanriBaseCnt As System.Windows.Forms.Label
        Friend WithEvents lblSelectPageCnt1 As System.Windows.Forms.Label
        Friend WithEvents lblSelectPageCnt2 As System.Windows.Forms.Label
        Friend WithEvents Label272 As System.Windows.Forms.Label
        Friend WithEvents Label261 As System.Windows.Forms.Label
        Friend WithEvents Label174 As System.Windows.Forms.Label
        Friend WithEvents Label294 As System.Windows.Forms.Label
        Friend WithEvents Label295 As System.Windows.Forms.Label
        Friend WithEvents lblKiSzenBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label180 As System.Windows.Forms.Label
        Friend WithEvents Label292 As System.Windows.Forms.Label
        Friend WithEvents lblKiClaimBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label288 As System.Windows.Forms.Label
        Friend WithEvents Label289 As System.Windows.Forms.Label
        Friend WithEvents Label290 As System.Windows.Forms.Label
        Friend WithEvents lblKiSqMiBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label284 As System.Windows.Forms.Label
        Friend WithEvents Label285 As System.Windows.Forms.Label
        Friend WithEvents Label286 As System.Windows.Forms.Label
        Friend WithEvents lblKiKyBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label280 As System.Windows.Forms.Label
        Friend WithEvents Label281 As System.Windows.Forms.Label
        Friend WithEvents Label282 As System.Windows.Forms.Label
        Friend WithEvents lblKiKysBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label277 As System.Windows.Forms.Label
        Friend WithEvents Label278 As System.Windows.Forms.Label
        Friend WithEvents lblKiHySetubiCnt As System.Windows.Forms.Label
        Friend WithEvents Label270 As System.Windows.Forms.Label
        Friend WithEvents Label275 As System.Windows.Forms.Label
        Friend WithEvents lblKiHyBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label271 As System.Windows.Forms.Label
        Friend WithEvents Label273 As System.Windows.Forms.Label
        Friend WithEvents lblKiBkBaseCnt As System.Windows.Forms.Label
        Friend WithEvents pnlRekiClear As System.Windows.Forms.Panel
        Friend WithEvents Label79 As System.Windows.Forms.Label
        Friend WithEvents Label298 As System.Windows.Forms.Label
        Friend WithEvents tabPageMenu As System.Windows.Forms.TabPage
        Friend WithEvents tabPageStart As System.Windows.Forms.TabPage
        Friend WithEvents lblFirstDescription1 As System.Windows.Forms.Label
        Friend WithEvents lblKFirstLabel As System.Windows.Forms.Label
        Friend WithEvents btnJizenListDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtJizenListPath As System.Windows.Forms.TextBox
        Friend WithEvents tabPageJigo As System.Windows.Forms.TabPage
        Friend WithEvents Label326 As System.Windows.Forms.Label
        Friend WithEvents Label325 As System.Windows.Forms.Label
        Friend WithEvents lblMenuCaution As System.Windows.Forms.Label
        Friend WithEvents grpMenuJizen As System.Windows.Forms.GroupBox
        Friend WithEvents btnMenuJizen As System.Windows.Forms.Button
        Friend WithEvents lblKMenuJizen As System.Windows.Forms.Label
        Friend WithEvents grpMenuDatacv As System.Windows.Forms.GroupBox
        Friend WithEvents btnMenuDatacv As System.Windows.Forms.Button
        Friend WithEvents lblKMenuDatacv As System.Windows.Forms.Label
        Friend WithEvents grpMenuGazocv As System.Windows.Forms.GroupBox
        Friend WithEvents btnMenuGazocv As System.Windows.Forms.Button
        Friend WithEvents lblMenuGazocv As System.Windows.Forms.Label
        Friend WithEvents grpMenuJigo As System.Windows.Forms.GroupBox
        Friend WithEvents btnMenuJigo As System.Windows.Forms.Button
        Friend WithEvents lblMenuJigo As System.Windows.Forms.Label
        Friend WithEvents lblMenuDescription1 As System.Windows.Forms.Label
        Friend WithEvents grpMenuHojyo As System.Windows.Forms.GroupBox
        Friend WithEvents btnMenuHojyo As System.Windows.Forms.Button
        Friend WithEvents lblMenuHojyo As System.Windows.Forms.Label
        Friend WithEvents btnJigoListDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtJigoListPath As System.Windows.Forms.TextBox
        Friend WithEvents tabCtrlJigo As System.Windows.Forms.TabControl
        Friend WithEvents tabPageJigo1 As System.Windows.Forms.TabPage
        Friend WithEvents lblJigoDescription1 As System.Windows.Forms.Label
        Friend WithEvents tabPageHojyo As System.Windows.Forms.TabPage
        Friend WithEvents btnKensyoListDirSeach As System.Windows.Forms.Button
        Friend WithEvents txtKensyoListPath As System.Windows.Forms.TextBox
        Friend WithEvents tabCtrlHojyo As System.Windows.Forms.TabControl
        Friend WithEvents tabPageHojyo2 As System.Windows.Forms.TabPage
        Friend WithEvents GroupBox25 As System.Windows.Forms.GroupBox
        Friend WithEvents Panel1 As System.Windows.Forms.Panel
        Friend WithEvents btnListSzenKysSorit As System.Windows.Forms.Button
        Friend WithEvents Label92 As System.Windows.Forms.Label
        Friend WithEvents Label95 As System.Windows.Forms.Label
        Friend WithEvents Label175 As System.Windows.Forms.Label
        Friend WithEvents lblCntSzenKysSorit As System.Windows.Forms.Label
        Friend WithEvents chkSzenKysSorit As System.Windows.Forms.CheckBox
        Friend WithEvents Label194 As System.Windows.Forms.Label
        Friend WithEvents Panel9 As System.Windows.Forms.Panel
        Friend WithEvents btnListBunkatumisyu As System.Windows.Forms.Button
        Friend WithEvents Label195 As System.Windows.Forms.Label
        Friend WithEvents Label196 As System.Windows.Forms.Label
        Friend WithEvents Label197 As System.Windows.Forms.Label
        Friend WithEvents lblCntBunkatumisyu As System.Windows.Forms.Label
        Friend WithEvents Panel10 As System.Windows.Forms.Panel
        Friend WithEvents btnListKojyosh As System.Windows.Forms.Button
        Friend WithEvents Label199 As System.Windows.Forms.Label
        Friend WithEvents Label200 As System.Windows.Forms.Label
        Friend WithEvents Label201 As System.Windows.Forms.Label
        Friend WithEvents lblCntKojyosh As System.Windows.Forms.Label
        Friend WithEvents chkBunkatumisyu As System.Windows.Forms.CheckBox
        Friend WithEvents Label207 As System.Windows.Forms.Label
        Friend WithEvents chkKojyosh As System.Windows.Forms.CheckBox
        Friend WithEvents Label208 As System.Windows.Forms.Label
        Friend WithEvents Label209 As System.Windows.Forms.Label
        Friend WithEvents lblHojyoDescription1 As System.Windows.Forms.Label
        Friend WithEvents Label393 As System.Windows.Forms.Label
        Friend WithEvents Label392 As System.Windows.Forms.Label
        Friend WithEvents CheckBox51 As System.Windows.Forms.CheckBox
        Friend WithEvents Label391 As System.Windows.Forms.Label
        Friend WithEvents CheckBox50 As System.Windows.Forms.CheckBox
        Friend WithEvents Panel24 As System.Windows.Forms.Panel
        Friend WithEvents Label390 As System.Windows.Forms.Label
        Friend WithEvents Label389 As System.Windows.Forms.Label
        Friend WithEvents Label388 As System.Windows.Forms.Label
        Friend WithEvents Label387 As System.Windows.Forms.Label
        Friend WithEvents Button21 As System.Windows.Forms.Button
        Friend WithEvents Panel23 As System.Windows.Forms.Panel
        Friend WithEvents Label386 As System.Windows.Forms.Label
        Friend WithEvents Label385 As System.Windows.Forms.Label
        Friend WithEvents Label384 As System.Windows.Forms.Label
        Friend WithEvents Label383 As System.Windows.Forms.Label
        Friend WithEvents Button20 As System.Windows.Forms.Button
        Friend WithEvents Label382 As System.Windows.Forms.Label
        Friend WithEvents CheckBox49 As System.Windows.Forms.CheckBox
        Friend WithEvents Panel22 As System.Windows.Forms.Panel
        Friend WithEvents Label381 As System.Windows.Forms.Label
        Friend WithEvents Label380 As System.Windows.Forms.Label
        Friend WithEvents Label379 As System.Windows.Forms.Label
        Friend WithEvents Label378 As System.Windows.Forms.Label
        Friend WithEvents Button19 As System.Windows.Forms.Button
        Friend WithEvents tabPageHojyo1 As System.Windows.Forms.TabPage
        Friend WithEvents GroupBox27 As System.Windows.Forms.GroupBox
        Friend WithEvents Panel25 As System.Windows.Forms.Panel
        Friend WithEvents btnRNRelRename As System.Windows.Forms.Button
        Friend WithEvents Label394 As System.Windows.Forms.Label
        Friend WithEvents lblRelRenameClaim As System.Windows.Forms.Label
        Friend WithEvents Label397 As System.Windows.Forms.Label
        Friend WithEvents txtAfRelRename As System.Windows.Forms.TextBox
        Friend WithEvents Label398 As System.Windows.Forms.Label
        Friend WithEvents txtBfRelRename As System.Windows.Forms.TextBox
        Friend WithEvents Label399 As System.Windows.Forms.Label
        Friend WithEvents Label400 As System.Windows.Forms.Label
        Friend WithEvents chkRelRename As System.Windows.Forms.CheckBox
        Friend WithEvents Label401 As System.Windows.Forms.Label
        Friend WithEvents Label402 As System.Windows.Forms.Label
        Friend WithEvents grpJigoUserSagyo As System.Windows.Forms.GroupBox
        Friend WithEvents Label333 As System.Windows.Forms.Label
        Friend WithEvents Label334 As System.Windows.Forms.Label
        Friend WithEvents Label336 As System.Windows.Forms.Label
        Friend WithEvents Label337 As System.Windows.Forms.Label
        Friend WithEvents Label303 As System.Windows.Forms.Label
        Friend WithEvents Label304 As System.Windows.Forms.Label
        Friend WithEvents Label305 As System.Windows.Forms.Label
        Friend WithEvents Label306 As System.Windows.Forms.Label
        Friend WithEvents tabPageJigo2 As System.Windows.Forms.TabPage
        Friend WithEvents grpJigoCmtTyuui As System.Windows.Forms.GroupBox
        Friend WithEvents Label23 As System.Windows.Forms.Label
        Friend WithEvents Label115 As System.Windows.Forms.Label
        Friend WithEvents grpJigoDonyuji As System.Windows.Forms.GroupBox
        Friend WithEvents Label108 As System.Windows.Forms.Label
        Friend WithEvents Label109 As System.Windows.Forms.Label
        Friend WithEvents Label110 As System.Windows.Forms.Label
        Friend WithEvents Label111 As System.Windows.Forms.Label
        Friend WithEvents Label112 As System.Windows.Forms.Label
        Friend WithEvents Label113 As System.Windows.Forms.Label
        Friend WithEvents Label98 As System.Windows.Forms.Label
        Friend WithEvents Label148 As System.Windows.Forms.Label
        Friend WithEvents Label147 As System.Windows.Forms.Label
        Friend WithEvents Label146 As System.Windows.Forms.Label
        Friend WithEvents Label143 As System.Windows.Forms.Label
        Friend WithEvents Label142 As System.Windows.Forms.Label
        Friend WithEvents Label99 As System.Windows.Forms.Label
        Friend WithEvents Label24 As System.Windows.Forms.Label
        Friend WithEvents grpKFirstNaiyo As System.Windows.Forms.GroupBox
        Friend WithEvents lblHajimeKi342 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi332 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi322 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi312 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi341 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi331 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi321 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi311 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi302 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi202 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi102 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi301 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi201 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi101 As System.Windows.Forms.Label
        Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
        Friend WithEvents chkOptionSelectKaikei As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectReform As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectClaim As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectFB As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectSyoki As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectSh As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectNk As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectSq As System.Windows.Forms.CheckBox
        Friend WithEvents chkOptionSelectKy As System.Windows.Forms.CheckBox
        Friend WithEvents lblSOpSeleNaiyo As System.Windows.Forms.Label
        Friend WithEvents pnlMenuMain As System.Windows.Forms.Panel
        Friend WithEvents pnlPrgMStart As System.Windows.Forms.Panel
        Friend WithEvents picIcoMStart2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoMStart1 As System.Windows.Forms.PictureBox
        Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgMStart As System.Windows.Forms.Label
        Friend WithEvents picArwMSession2 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwMSession1 As System.Windows.Forms.PictureBox
        Friend WithEvents pnlPrgMEnd As System.Windows.Forms.Panel
        Friend WithEvents picIcoMEnd2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoMEnd1 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoJizen1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgMEnd As System.Windows.Forms.Label
        Friend WithEvents pnlPrgMMenu As System.Windows.Forms.Panel
        Friend WithEvents picIcoMMenu2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoMMenu1 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoSession1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgMMenu As System.Windows.Forms.Label
        Friend WithEvents pnlPrgMSyoki As System.Windows.Forms.Panel
        Friend WithEvents picIcoMSyoki2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoMSyoki1 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoSyoki1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgMSyoki As System.Windows.Forms.Label
        Friend WithEvents pnlPrgMSession As System.Windows.Forms.Panel
        Friend WithEvents picIcoMSession2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoMSession1 As System.Windows.Forms.PictureBox
        Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgMSession As System.Windows.Forms.Label
        Friend WithEvents picArwMEnd2 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwMMenu2 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwMSyoki2 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwMEnd1 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwMMenu1 As System.Windows.Forms.PictureBox
        Friend WithEvents picArwMSyoki1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblHidden4 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi352 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeKi351 As System.Windows.Forms.Label
        Friend WithEvents Label57 As System.Windows.Forms.Label
        Friend WithEvents Label91 As System.Windows.Forms.Label
        Friend WithEvents Label58 As System.Windows.Forms.Label
        Friend WithEvents lblKOpSeleNaiyo As System.Windows.Forms.Label
        Friend WithEvents lblKOpSeleTitle As System.Windows.Forms.Label
        Friend WithEvents pnlOptionSelect As System.Windows.Forms.Panel
        Friend WithEvents Label137 As System.Windows.Forms.Label
        Friend WithEvents lblJizenPageNum1 As System.Windows.Forms.Label
        Friend WithEvents lblJizenPageNum2 As System.Windows.Forms.Label
        Friend WithEvents chkFBANSERSetuzoku As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstANSERArea As System.Windows.Forms.CheckBox
        Friend WithEvents chkMstANSERAccpoint As System.Windows.Forms.CheckBox
        Friend WithEvents chkSzenRelfile As System.Windows.Forms.CheckBox
        Friend WithEvents Label160 As System.Windows.Forms.Label
        Friend WithEvents Label162 As System.Windows.Forms.Label
        Friend WithEvents lblKiSqAzBaseCnt As System.Windows.Forms.Label
        Friend WithEvents pnlJizenYanuso As System.Windows.Forms.Panel
        Friend WithEvents Panel8 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenYanuso As System.Windows.Forms.Button
        Friend WithEvents Label188 As System.Windows.Forms.Label
        Friend WithEvents Label189 As System.Windows.Forms.Label
        Friend WithEvents Label190 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenYanuso As System.Windows.Forms.Label
        Friend WithEvents chkJizenYanuso As System.Windows.Forms.CheckBox
        Friend WithEvents Label96 As System.Windows.Forms.Label
        Friend WithEvents pnlJizenAzu As System.Windows.Forms.Panel
        Friend WithEvents pnlJizenKai As System.Windows.Forms.Panel
        Friend WithEvents Panel12 As System.Windows.Forms.Panel
        Friend WithEvents pnlJizenHasseiOw As System.Windows.Forms.Panel
        Friend WithEvents pnlJizenBkhourei As System.Windows.Forms.Panel
        Friend WithEvents pnlJizenMinus As System.Windows.Forms.Panel
        Friend WithEvents pnlJizenSzenKyshutan As System.Windows.Forms.Panel
        Friend WithEvents pnlJizenSo As System.Windows.Forms.Panel
        Friend WithEvents pnlJizenHurikae As System.Windows.Forms.Panel
        Friend WithEvents pnlKiKys As System.Windows.Forms.Panel
        Friend WithEvents pnlKiHy As System.Windows.Forms.Panel
        Friend WithEvents pnlKiBk As System.Windows.Forms.Panel
        Friend WithEvents pnlKiKy As System.Windows.Forms.Panel
        Friend WithEvents pnlKiSq As System.Windows.Forms.Panel
        Friend WithEvents pnlSzen As System.Windows.Forms.Panel
        Friend WithEvents pnlKiClaim As System.Windows.Forms.Panel
        Friend WithEvents pnlKiGyCyukaiBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiOw As System.Windows.Forms.Panel
        Friend WithEvents pnlKiJisya As System.Windows.Forms.Panel
        Friend WithEvents pnlKiSyskanriBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiGySyuzenBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiGyYatinhosyoBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiGyLifelineBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiGyHokenBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiGySisetuBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiGySekoBase As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstBus As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstHokenrui As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstTokuyaku As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstKasyoClaimrui As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstTitle As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstArea As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstSchool As System.Windows.Forms.Panel
        Friend WithEvents pnlKiMstHendo As System.Windows.Forms.Panel
        Friend WithEvents pnlRelRename As System.Windows.Forms.Panel
        Friend WithEvents pnlSzenKysSorit As System.Windows.Forms.Panel
        Friend WithEvents pnlKojyosh As System.Windows.Forms.Panel
        Friend WithEvents pnlBunkatumisyu As System.Windows.Forms.Panel
        Friend WithEvents lblHidden5 As System.Windows.Forms.Label
        Friend WithEvents pnlKiOpKys As System.Windows.Forms.Panel
        Friend WithEvents chkKiOpKys As System.Windows.Forms.CheckBox
        Friend WithEvents Label184 As System.Windows.Forms.Label
        Friend WithEvents Panel18 As System.Windows.Forms.Panel
        Friend WithEvents btnKiOpKys As System.Windows.Forms.Button
        Friend WithEvents Label187 As System.Windows.Forms.Label
        Friend WithEvents Label191 As System.Windows.Forms.Label
        Friend WithEvents Label193 As System.Windows.Forms.Label
        Friend WithEvents lblCntKiOpKys As System.Windows.Forms.Label
        Friend WithEvents pnlKiOpOw As System.Windows.Forms.Panel
        Friend WithEvents chkKiOpOw As System.Windows.Forms.CheckBox
        Friend WithEvents Label169 As System.Windows.Forms.Label
        Friend WithEvents Panel16 As System.Windows.Forms.Panel
        Friend WithEvents btnKiOpOw As System.Windows.Forms.Button
        Friend WithEvents Label172 As System.Windows.Forms.Label
        Friend WithEvents Label177 As System.Windows.Forms.Label
        Friend WithEvents Label178 As System.Windows.Forms.Label
        Friend WithEvents lblCntKiOpOw As System.Windows.Forms.Label
        Friend WithEvents Label206 As System.Windows.Forms.Label
        Friend WithEvents Label205 As System.Windows.Forms.Label
        Friend WithEvents btnListKiOpOw As System.Windows.Forms.Button
        Friend WithEvents btnListKiOpKys As System.Windows.Forms.Button
        Friend WithEvents cmbV10Networklib As System.Windows.Forms.ComboBox
        Friend WithEvents cmbV7Networklib As System.Windows.Forms.ComboBox
        Friend WithEvents Label179 As System.Windows.Forms.Label
        Friend WithEvents Label198 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenBkhoureiNotCV As System.Windows.Forms.Label
        Friend WithEvents Label231 As System.Windows.Forms.Label
        Friend WithEvents Label221 As System.Windows.Forms.Label
        Friend WithEvents lblRelRenameSzen As System.Windows.Forms.Label
        Friend WithEvents Label229 As System.Windows.Forms.Label
        Friend WithEvents Label217 As System.Windows.Forms.Label
        Friend WithEvents Label215 As System.Windows.Forms.Label
        Friend WithEvents Label212 As System.Windows.Forms.Label
        Friend WithEvents chkRelRenameSzen As System.Windows.Forms.CheckBox
        Friend WithEvents chkRelRenameClaim As System.Windows.Forms.CheckBox
        Friend WithEvents pnlRelationSet As System.Windows.Forms.Panel
        Friend WithEvents chkRelationSet As System.Windows.Forms.CheckBox
        Friend WithEvents Label225 As System.Windows.Forms.Label
        Friend WithEvents Panel17 As System.Windows.Forms.Panel
        Friend WithEvents btnListRelationSet As System.Windows.Forms.Button
        Friend WithEvents Label234 As System.Windows.Forms.Label
        Friend WithEvents Label237 As System.Windows.Forms.Label
        Friend WithEvents Label239 As System.Windows.Forms.Label
        Friend WithEvents lblCntRelationSet As System.Windows.Forms.Label
        Friend WithEvents Label242 As System.Windows.Forms.Label
        Friend WithEvents lblLine3 As System.Windows.Forms.Label
        Friend WithEvents lblLine4 As System.Windows.Forms.Label
        Friend WithEvents lblLine0 As System.Windows.Forms.Label
        Friend WithEvents Label245 As System.Windows.Forms.Label
        Friend WithEvents pnlRefresh As System.Windows.Forms.Panel
        Friend WithEvents Label251 As System.Windows.Forms.Label
        Friend WithEvents btnRefresh As System.Windows.Forms.Button
        Friend WithEvents pnlJizenListPath As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoListPath As System.Windows.Forms.Panel
        Friend WithEvents Label248 As System.Windows.Forms.Label
        Friend WithEvents Label254 As System.Windows.Forms.Label
        Friend WithEvents pnlKensyoListPath As System.Windows.Forms.Panel
        Friend WithEvents Label257 As System.Windows.Forms.Label
        Friend WithEvents pnlJizenKagi As System.Windows.Forms.Panel
        Friend WithEvents chkJizenKagi As System.Windows.Forms.CheckBox
        Friend WithEvents Label262 As System.Windows.Forms.Label
        Friend WithEvents Panel19 As System.Windows.Forms.Panel
        Friend WithEvents Label266 As System.Windows.Forms.Label
        Friend WithEvents Label269 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenKyKagi As System.Windows.Forms.Label
        Friend WithEvents btnListJizenHyKagi As System.Windows.Forms.Button
        Friend WithEvents Label276 As System.Windows.Forms.Label
        Friend WithEvents Label279 As System.Windows.Forms.Label
        Friend WithEvents Label283 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenHyKagi As System.Windows.Forms.Label
        Friend WithEvents btnListJizenKyKagi As System.Windows.Forms.Button
        Friend WithEvents Label291 As System.Windows.Forms.Label
        Friend WithEvents Label287 As System.Windows.Forms.Label
        Friend WithEvents Label274 As System.Windows.Forms.Label
        Friend WithEvents Label307 As System.Windows.Forms.Label
        Friend WithEvents Label297 As System.Windows.Forms.Label
        Friend WithEvents Label296 As System.Windows.Forms.Label
        Friend WithEvents Label293 As System.Windows.Forms.Label
        Friend WithEvents tabPageBase210 As System.Windows.Forms.TabPage
        Friend WithEvents grpRendo As System.Windows.Forms.GroupBox
        Friend WithEvents chkRendoSosinJisyaweb As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoSosinBase As System.Windows.Forms.CheckBox
        Friend WithEvents pnlRendo As System.Windows.Forms.Panel
        Friend WithEvents chkKiRendoBase As System.Windows.Forms.CheckBox
        Friend WithEvents lblKiRendoBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label314 As System.Windows.Forms.Label
        Friend WithEvents Label316 As System.Windows.Forms.Label
        Friend WithEvents lblHFirstLabel As System.Windows.Forms.Label
        Friend WithEvents grpHFirstNaiyo As System.Windows.Forms.GroupBox
        Friend WithEvents lblHajimeH342 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH322 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH341 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH321 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH302 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH202 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH102 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH301 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH201 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH101 As System.Windows.Forms.Label
        Friend WithEvents lblKSessionDescription1 As System.Windows.Forms.Label
        Friend WithEvents lblHSessionDescription1 As System.Windows.Forms.Label
        Friend WithEvents lblHOpSeleNaiyo As System.Windows.Forms.Label
        Friend WithEvents lblHOpSeleTitle As System.Windows.Forms.Label
        Friend WithEvents lblHMenuDatacv As System.Windows.Forms.Label
        Friend WithEvents pnlKJigoCmtSyudo As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoCmtCsvSyuturyoku As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoCmtSyosiki As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoCmtSyusi As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoCmtHeiko As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoCmtSoKotiku As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoCmtNkNyuryoku As System.Windows.Forms.Panel
        Friend WithEvents pnlJigoCmtSqKotiku As System.Windows.Forms.Panel
        Friend WithEvents Label121 As System.Windows.Forms.Label
        Friend WithEvents pnlHJigoCmtSyudo As System.Windows.Forms.Panel
        Friend WithEvents Label136 As System.Windows.Forms.Label
        Friend WithEvents Label317 As System.Windows.Forms.Label
        Friend WithEvents Label318 As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmt07 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmt10 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmt02 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmt08 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmt04 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmt03 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmt01 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmt09 As System.Windows.Forms.Panel
        Friend WithEvents pnlDcFstCmtK02 As System.Windows.Forms.Panel
        Friend WithEvents lblHDatacvHajimeniLabel As System.Windows.Forms.Label
        Friend WithEvents Label124 As System.Windows.Forms.Label
        Friend WithEvents Label328 As System.Windows.Forms.Label
        Friend WithEvents lblHDatacvEndLabel As System.Windows.Forms.Label
        Friend WithEvents tabPageHJizen1 As System.Windows.Forms.TabPage
        Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
        Friend WithEvents pnlHJizenGazoKeisiki As System.Windows.Forms.Panel
        Friend WithEvents chkHJizenGazoKeisiki As System.Windows.Forms.CheckBox
        Friend WithEvents Label344 As System.Windows.Forms.Label
        Friend WithEvents pnlHJizenTyukan As System.Windows.Forms.Panel
        Friend WithEvents chkHJizenTyukan As System.Windows.Forms.CheckBox
        Friend WithEvents Label350 As System.Windows.Forms.Label
        Friend WithEvents Panel26 As System.Windows.Forms.Panel
        Friend WithEvents btnHJizenTyukanOpen As System.Windows.Forms.Button
        Friend WithEvents Label355 As System.Windows.Forms.Label
        Friend WithEvents Label352 As System.Windows.Forms.Label
        Friend WithEvents Label322 As System.Windows.Forms.Label
        Friend WithEvents Label343 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH312 As System.Windows.Forms.Label
        Friend WithEvents lblHajimeH311 As System.Windows.Forms.Label
        Friend WithEvents chkRendoSosinSuumo As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoSosinAthome As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoSosinHomes As System.Windows.Forms.CheckBox
        Friend WithEvents lblKiMstBikotitleCnt As System.Windows.Forms.Label
        Friend WithEvents Label236 As System.Windows.Forms.Label
        Friend WithEvents Label235 As System.Windows.Forms.Label
        Friend WithEvents lblKiMstGazotitleCnt As System.Windows.Forms.Label
        Friend WithEvents Label346 As System.Windows.Forms.Label
        Friend WithEvents Label347 As System.Windows.Forms.Label
        Friend WithEvents chkMstGazotitle As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoKokokuSuumo As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoKokokuAthome As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoKokokuHomes As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoKokokuJisyaweb As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoHyrui As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoHysosin As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoBtoBgroup As System.Windows.Forms.CheckBox
        Friend WithEvents chkRendoMapdisp As System.Windows.Forms.CheckBox
        Friend WithEvents pnlPrgDCConv As System.Windows.Forms.Panel
        Friend WithEvents picIcoDCConv2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoDCConv1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgDCConv As System.Windows.Forms.Label
        Friend WithEvents pnlPrgDCRelation As System.Windows.Forms.Panel
        Friend WithEvents picIcoDCRelation2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoDCRelation1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgDCRelation As System.Windows.Forms.Label
        Friend WithEvents pnlPrgDCFileWrite As System.Windows.Forms.Panel
        Friend WithEvents picIcoDCFileWrite2 As System.Windows.Forms.PictureBox
        Friend WithEvents picIcoDCFileWrite1 As System.Windows.Forms.PictureBox
        Friend WithEvents lblPrgDCFileWrite As System.Windows.Forms.Label
        Friend WithEvents btnHJizenTyukanOpen2 As System.Windows.Forms.Button
        Friend WithEvents Label349 As System.Windows.Forms.Label
        Friend WithEvents lblMiddleFile2 As System.Windows.Forms.Label
        Friend WithEvents txtMidDirPath2 As System.Windows.Forms.TextBox
        Friend WithEvents btnMidDirSeach2 As System.Windows.Forms.Button
        Friend WithEvents tabPageJizen4 As System.Windows.Forms.TabPage
        Friend WithEvents grpJizen4 As System.Windows.Forms.GroupBox
        Friend WithEvents Panel21 As System.Windows.Forms.Panel
        Friend WithEvents chkJizenNonJisyaKoza As System.Windows.Forms.CheckBox
        Friend WithEvents Label365 As System.Windows.Forms.Label
        Friend WithEvents Panel27 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenNonJisyaKoza As System.Windows.Forms.Button
        Friend WithEvents Label366 As System.Windows.Forms.Label
        Friend WithEvents Label367 As System.Windows.Forms.Label
        Friend WithEvents Label368 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenNonJisyaKoza As System.Windows.Forms.Label
        Friend WithEvents Label370 As System.Windows.Forms.Label
        Friend WithEvents chkSyskanriinit As System.Windows.Forms.CheckBox
        Friend WithEvents Label83 As System.Windows.Forms.Label
        Friend WithEvents txtRendoID As System.Windows.Forms.TextBox
        Friend WithEvents Panel13 As System.Windows.Forms.Panel
        Friend WithEvents chkJizenSimeSokin As System.Windows.Forms.CheckBox
        Friend WithEvents Label149 As System.Windows.Forms.Label
        Friend WithEvents Panel14 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenSimeSokin As System.Windows.Forms.Button
        Friend WithEvents Label310 As System.Windows.Forms.Label
        Friend WithEvents Label348 As System.Windows.Forms.Label
        Friend WithEvents Label351 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenSimeSokin As System.Windows.Forms.Label
        Friend WithEvents Panel15 As System.Windows.Forms.Panel
        Friend WithEvents chkJizenKozameigikana As System.Windows.Forms.CheckBox
        Friend WithEvents Label353 As System.Windows.Forms.Label
        Friend WithEvents Panel20 As System.Windows.Forms.Panel
        Friend WithEvents btnListJizenKozameigikana As System.Windows.Forms.Button
        Friend WithEvents Label354 As System.Windows.Forms.Label
        Friend WithEvents Label356 As System.Windows.Forms.Label
        Friend WithEvents Label357 As System.Windows.Forms.Label
        Friend WithEvents lblCntJizenKozameigikana As System.Windows.Forms.Label
        Friend WithEvents Label358 As System.Windows.Forms.Label
        Friend WithEvents tabPageJizen5 As System.Windows.Forms.TabPage
        Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
        Friend WithEvents Label374 As System.Windows.Forms.Label
        Friend WithEvents Label359 As System.Windows.Forms.Label
        Friend WithEvents Label360 As System.Windows.Forms.Label
        Friend WithEvents btnExistMidToBaseMid As System.Windows.Forms.Button
        Friend WithEvents grpExistMidToBaseMid As System.Windows.Forms.GroupBox
        Friend WithEvents Label361 As System.Windows.Forms.Label
        Friend WithEvents txtExistMidToBaseMid As System.Windows.Forms.TextBox
        Friend WithEvents btnExistMidToBaseMidDirSerach As System.Windows.Forms.Button
        Friend WithEvents Label106 As System.Windows.Forms.Label
        Friend WithEvents Label107 As System.Windows.Forms.Label
        Friend WithEvents pnlCVType As System.Windows.Forms.Panel
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents optCVNew As System.Windows.Forms.RadioButton
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents optCVAdd As System.Windows.Forms.RadioButton
        Friend WithEvents pnlDcFstCmtK01 As System.Windows.Forms.Panel
        Friend WithEvents Label371 As System.Windows.Forms.Label
        Friend WithEvents Label372 As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmt06 As System.Windows.Forms.Panel
        Friend WithEvents Label364 As System.Windows.Forms.Label
        Friend WithEvents Label369 As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmt05 As System.Windows.Forms.Panel
        Friend WithEvents Label362 As System.Windows.Forms.Label
        Friend WithEvents Label363 As System.Windows.Forms.Label
        Friend WithEvents pnlUserName As System.Windows.Forms.Panel
        Friend WithEvents pnlLogPath As System.Windows.Forms.Panel
        Friend WithEvents pnlOptSelect As System.Windows.Forms.Panel
        Friend WithEvents lblTitleH As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmt11 As System.Windows.Forms.Panel
        Friend WithEvents Label21 As System.Windows.Forms.Label
        Friend WithEvents Label25 As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmtH02 As System.Windows.Forms.Panel
        Friend WithEvents Label27 As System.Windows.Forms.Label
        Friend WithEvents Label28 As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmtH01 As System.Windows.Forms.Panel
        Friend WithEvents Label158 As System.Windows.Forms.Label
        Friend WithEvents Label159 As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmtH99 As System.Windows.Forms.Panel
        Friend WithEvents Label301 As System.Windows.Forms.Label
        Friend WithEvents Label302 As System.Windows.Forms.Label
        Friend WithEvents pnlDcFstCmtK99 As System.Windows.Forms.Panel
        Friend WithEvents Label299 As System.Windows.Forms.Label
        Friend WithEvents Label300 As System.Windows.Forms.Label
        Friend WithEvents pnlUnyoKaisi As System.Windows.Forms.Panel
        Friend WithEvents pgbCheck As System.Windows.Forms.ProgressBar
        Friend WithEvents lblPgbCheck As System.Windows.Forms.Label
        Friend WithEvents lblCheckSituation As System.Windows.Forms.Label
        Friend WithEvents pnlPrgChk As System.Windows.Forms.Panel
        Friend WithEvents Label308 As System.Windows.Forms.Label
        Friend WithEvents lblKiSqOwKojoBaseCnt As System.Windows.Forms.Label
        Friend WithEvents Label311 As System.Windows.Forms.Label
        Friend WithEvents pnlConvertType As System.Windows.Forms.Panel
        Friend WithEvents grpDevSettingX As System.Windows.Forms.GroupBox
        Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
        Friend WithEvents lblHMenuJizen As System.Windows.Forms.Label

    End Class

End Namespace