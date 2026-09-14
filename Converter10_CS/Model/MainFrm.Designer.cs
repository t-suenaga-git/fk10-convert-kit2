using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace Converter10.Njc.Frm
{

    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class MainFrm : Form
    {

        // フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
        [DebuggerNonUserCode()]
        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing && components is not null)
                {
                    components.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        // Windows フォーム デザイナーで必要です。
        private System.ComponentModel.IContainer components;

        // メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
        // Windows フォーム デザイナーを使用して変更できます。  
        // コード エディターを使って変更しないでください。
        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainFrm));
            this.chkRelationFile = new System.Windows.Forms.CheckBox();
            this.Label9 = new System.Windows.Forms.Label();
            this.pnlMenuDatacv = new System.Windows.Forms.Panel();
            this.pnlPrgDCConv = new System.Windows.Forms.Panel();
            this.picIcoDCConv2 = new System.Windows.Forms.PictureBox();
            this.picIcoDCConv1 = new System.Windows.Forms.PictureBox();
            this.lblPrgDCConv = new System.Windows.Forms.Label();
            this.pnlPrgDCRelation = new System.Windows.Forms.Panel();
            this.picIcoDCRelation2 = new System.Windows.Forms.PictureBox();
            this.picIcoDCRelation1 = new System.Windows.Forms.PictureBox();
            this.lblPrgDCRelation = new System.Windows.Forms.Label();
            this.pnlPrgDCFileWrite = new System.Windows.Forms.Panel();
            this.picIcoDCFileWrite2 = new System.Windows.Forms.PictureBox();
            this.picIcoDCFileWrite1 = new System.Windows.Forms.PictureBox();
            this.lblPrgDCFileWrite = new System.Windows.Forms.Label();
            this.picArwDCEnd2 = new System.Windows.Forms.PictureBox();
            this.picArwDCJikko2 = new System.Windows.Forms.PictureBox();
            this.picArwDCSelect2 = new System.Windows.Forms.PictureBox();
            this.picArwDCEnd1 = new System.Windows.Forms.PictureBox();
            this.picArwDCJikko1 = new System.Windows.Forms.PictureBox();
            this.picArwDCSelect1 = new System.Windows.Forms.PictureBox();
            this.pnlPrgDCEnd = new System.Windows.Forms.Panel();
            this.picIcoDCEnd2 = new System.Windows.Forms.PictureBox();
            this.picIcoDCEnd1 = new System.Windows.Forms.PictureBox();
            this.lblPrgDCEnd = new System.Windows.Forms.Label();
            this.pnlPrgDCJikko = new System.Windows.Forms.Panel();
            this.picIcoDCJikko2 = new System.Windows.Forms.PictureBox();
            this.picIcoDCJikko1 = new System.Windows.Forms.PictureBox();
            this.lblPrgDCJikko = new System.Windows.Forms.Label();
            this.pnlPrgDCSelect = new System.Windows.Forms.Panel();
            this.picIcoDCSelect2 = new System.Windows.Forms.PictureBox();
            this.picIcoDCSelect1 = new System.Windows.Forms.PictureBox();
            this.lblPrgDCSelect = new System.Windows.Forms.Label();
            this.pnlPrgDCHajimeni = new System.Windows.Forms.Panel();
            this.picIcoDCHajimeni2 = new System.Windows.Forms.PictureBox();
            this.picIcoDCHajimeni1 = new System.Windows.Forms.PictureBox();
            this.lblPrgDCHajimeni = new System.Windows.Forms.Label();
            this.lblHidden1 = new System.Windows.Forms.Label();
            this.tabCtrlMain = new System.Windows.Forms.TabControl();
            this.tabPageDev = new System.Windows.Forms.TabPage();
            this.GroupBox6 = new System.Windows.Forms.GroupBox();
            this.Label361 = new System.Windows.Forms.Label();
            this.txtExistMidToBaseMid = new System.Windows.Forms.TextBox();
            this.btnExistMidToBaseMid = new System.Windows.Forms.Button();
            this.btnExistMidToBaseMidDirSerach = new System.Windows.Forms.Button();
            this.grpDevSettingX = new System.Windows.Forms.GroupBox();
            this.chkChildItemControl = new System.Windows.Forms.CheckBox();
            this.Label48 = new System.Windows.Forms.Label();
            this.Label41 = new System.Windows.Forms.Label();
            this.Label50 = new System.Windows.Forms.Label();
            this.Label49 = new System.Windows.Forms.Label();
            this.txtLogOutputCnt = new System.Windows.Forms.TextBox();
            this.PictureBox1 = new System.Windows.Forms.PictureBox();
            this.grpDevSetting1 = new System.Windows.Forms.GroupBox();
            this.Label51 = new System.Windows.Forms.Label();
            this.Label360 = new System.Windows.Forms.Label();
            this.chkKyNewest = new System.Windows.Forms.CheckBox();
            this.chkSyskanriinit = new System.Windows.Forms.CheckBox();
            this.Label46 = new System.Windows.Forms.Label();
            this.Label45 = new System.Windows.Forms.Label();
            this.Label44 = new System.Windows.Forms.Label();
            this.chkRelTblDrop = new System.Windows.Forms.CheckBox();
            this.Label43 = new System.Windows.Forms.Label();
            this.Label42 = new System.Windows.Forms.Label();
            this.chkV7ViewDrop = new System.Windows.Forms.CheckBox();
            this.chkLogTblDrop = new System.Windows.Forms.CheckBox();
            this.chkOverWrite = new System.Windows.Forms.CheckBox();
            this.Label52 = new System.Windows.Forms.Label();
            this.grpDevSetting2 = new System.Windows.Forms.GroupBox();
            this.btnRelDirSeach = new System.Windows.Forms.Button();
            this.txtRelationDirPath = new System.Windows.Forms.TextBox();
            this.lblRelationDir = new System.Windows.Forms.Label();
            this.Label11 = new System.Windows.Forms.Label();
            this.chkRelation = new System.Windows.Forms.CheckBox();
            this.Label47 = new System.Windows.Forms.Label();
            this.chkMidNotStop = new System.Windows.Forms.CheckBox();
            this.chkCVStart = new System.Windows.Forms.CheckBox();
            this.Label8 = new System.Windows.Forms.Label();
            this.Label10 = new System.Windows.Forms.Label();
            this.chkMiddleFile = new System.Windows.Forms.CheckBox();
            this.grpDevSetting3 = new System.Windows.Forms.GroupBox();
            this.Label55 = new System.Windows.Forms.Label();
            this.Label54 = new System.Windows.Forms.Label();
            this.chkRommDuplicate = new System.Windows.Forms.CheckBox();
            this.chkEmptyRoomNo = new System.Windows.Forms.CheckBox();
            this.Label53 = new System.Windows.Forms.Label();
            this.grpRelation = new System.Windows.Forms.GroupBox();
            this.GroupBox3 = new System.Windows.Forms.GroupBox();
            this.Label39 = new System.Windows.Forms.Label();
            this.Label33 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label56 = new System.Windows.Forms.Label();
            this.Label16 = new System.Windows.Forms.Label();
            this.Label15 = new System.Windows.Forms.Label();
            this.Label14 = new System.Windows.Forms.Label();
            this.Label12 = new System.Windows.Forms.Label();
            this.Label18 = new System.Windows.Forms.Label();
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.chkMstKasyorui = new System.Windows.Forms.CheckBox();
            this.chkMstGenjotokuyaku = new System.Windows.Forms.CheckBox();
            this.lblTokuyakuInfo = new System.Windows.Forms.Label();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.Label34 = new System.Windows.Forms.Label();
            this.Label29 = new System.Windows.Forms.Label();
            this.chkMstKozasyubetu = new System.Windows.Forms.CheckBox();
            this.chkMstSetubi = new System.Windows.Forms.CheckBox();
            this.chkMstKozo = new System.Windows.Forms.CheckBox();
            this.chkMstTorihikitaiyo = new System.Windows.Forms.CheckBox();
            this.chkMstNkinkbn = new System.Windows.Forms.CheckBox();
            this.chkMstHyrui = new System.Windows.Forms.CheckBox();
            this.chkMstNkinkomok = new System.Windows.Forms.CheckBox();
            this.chkMstBkrui = new System.Windows.Forms.CheckBox();
            this.grptaihi = new System.Windows.Forms.GroupBox();
            this.Label7 = new System.Windows.Forms.Label();
            this.btnFileSeach = new System.Windows.Forms.Button();
            this.Label5 = new System.Windows.Forms.Label();
            this.chkGyYatinhosyoKoza = new System.Windows.Forms.CheckBox();
            this.CheckBox66 = new System.Windows.Forms.CheckBox();
            this.lblGy = new System.Windows.Forms.Label();
            this.lblMst = new System.Windows.Forms.Label();
            this.lblKys = new System.Windows.Forms.Label();
            this.lblOwner = new System.Windows.Forms.Label();
            this.lblJisya = new System.Windows.Forms.Label();
            this.Label17 = new System.Windows.Forms.Label();
            this.Label19 = new System.Windows.Forms.Label();
            this.Label20 = new System.Windows.Forms.Label();
            this.chkMs25 = new System.Windows.Forms.CheckBox();
            this.chkMsKeiyakusyubetu = new System.Windows.Forms.CheckBox();
            this.chkMs23 = new System.Windows.Forms.CheckBox();
            this.chkMs24 = new System.Windows.Forms.CheckBox();
            this.chkMsOwevent = new System.Windows.Forms.CheckBox();
            this.chkMsYane = new System.Windows.Forms.CheckBox();
            this.chkMsKeikaikakunin = new System.Windows.Forms.CheckBox();
            this.chkMsKinyu = new System.Windows.Forms.CheckBox();
            this.chkMsKinyuten = new System.Windows.Forms.CheckBox();
            this.tabPageStart = new System.Windows.Forms.TabPage();
            this.lblFirstDescription1 = new System.Windows.Forms.Label();
            this.lblHFirstLabel = new System.Windows.Forms.Label();
            this.grpHFirstNaiyo = new System.Windows.Forms.GroupBox();
            this.lblHajimeH312 = new System.Windows.Forms.Label();
            this.lblHajimeH311 = new System.Windows.Forms.Label();
            this.lblHajimeH342 = new System.Windows.Forms.Label();
            this.lblHajimeH322 = new System.Windows.Forms.Label();
            this.lblHajimeH341 = new System.Windows.Forms.Label();
            this.lblHajimeH321 = new System.Windows.Forms.Label();
            this.lblHajimeH302 = new System.Windows.Forms.Label();
            this.lblHajimeH202 = new System.Windows.Forms.Label();
            this.lblHajimeH102 = new System.Windows.Forms.Label();
            this.lblHajimeH301 = new System.Windows.Forms.Label();
            this.lblHajimeH201 = new System.Windows.Forms.Label();
            this.lblHajimeH101 = new System.Windows.Forms.Label();
            this.PictureBox2 = new System.Windows.Forms.PictureBox();
            this.tabPageSession = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.pre_table_connection_string = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.grpTimeOut = new System.Windows.Forms.GroupBox();
            this.txtTimeOut = new System.Windows.Forms.TextBox();
            this.lblTimeOutSec = new System.Windows.Forms.Label();
            this.lblSessionCaution = new System.Windows.Forms.Label();
            this.lblHSessionDescription1 = new System.Windows.Forms.Label();
            this.Label36 = new System.Windows.Forms.Label();
            this.Label35 = new System.Windows.Forms.Label();
            this.lblNetworklib = new System.Windows.Forms.Label();
            this.Label30 = new System.Windows.Forms.Label();
            this.Label13 = new System.Windows.Forms.Label();
            this.grp10ConnectInfo = new System.Windows.Forms.GroupBox();
            this.cmbV10Networklib = new System.Windows.Forms.ComboBox();
            this.grpV10Authent = new System.Windows.Forms.GroupBox();
            this.optV10Authent2 = new System.Windows.Forms.RadioButton();
            this.optV10Authent1 = new System.Windows.Forms.RadioButton();
            this.txtV10Pass = new System.Windows.Forms.TextBox();
            this.txtV10User = new System.Windows.Forms.TextBox();
            this.txtV10Catalog = new System.Windows.Forms.TextBox();
            this.txtV10Server = new System.Windows.Forms.TextBox();
            this.lblV10 = new System.Windows.Forms.Label();
            this.btnDefConInfoRead = new System.Windows.Forms.Button();
            this.btnConnectTest = new System.Windows.Forms.Button();
            this.tabPageSyoki = new System.Windows.Forms.TabPage();
            this.pnlUserName = new System.Windows.Forms.Panel();
            this.lblExecutor = new System.Windows.Forms.Label();
            this.txtRecUser = new System.Windows.Forms.TextBox();
            this.Label58 = new System.Windows.Forms.Label();
            this.pnlLogPath = new System.Windows.Forms.Panel();
            this.lblLogDir = new System.Windows.Forms.Label();
            this.Label57 = new System.Windows.Forms.Label();
            this.btnLogDirSeach = new System.Windows.Forms.Button();
            this.txtLogDirPath = new System.Windows.Forms.TextBox();
            this.lblSyokiDescription1 = new System.Windows.Forms.Label();
            this.tabPageMenu = new System.Windows.Forms.TabPage();
            this.lblMenuCaution = new System.Windows.Forms.Label();
            this.grpMenuJizen = new System.Windows.Forms.GroupBox();
            this.lblKMenuJizen = new System.Windows.Forms.Label();
            this.lblHMenuJizen = new System.Windows.Forms.Label();
            this.btnMenuJizen = new System.Windows.Forms.Button();
            this.grpMenuDatacv = new System.Windows.Forms.GroupBox();
            this.lblKMenuDatacv = new System.Windows.Forms.Label();
            this.lblHMenuDatacv = new System.Windows.Forms.Label();
            this.btnMenuDatacv = new System.Windows.Forms.Button();
            this.grpMenuJigo = new System.Windows.Forms.GroupBox();
            this.btnMenuJigo = new System.Windows.Forms.Button();
            this.lblMenuJigo = new System.Windows.Forms.Label();
            this.lblMenuDescription1 = new System.Windows.Forms.Label();
            this.tabPageJizen = new System.Windows.Forms.TabPage();
            this.pnlJizenListPath = new System.Windows.Forms.Panel();
            this.Label254 = new System.Windows.Forms.Label();
            this.btnJizenListDirSeach = new System.Windows.Forms.Button();
            this.txtJizenListPath = new System.Windows.Forms.TextBox();
            this.lblLine1 = new System.Windows.Forms.Label();
            this.lblHidden3 = new System.Windows.Forms.Label();
            this.tabCtrlJizen = new System.Windows.Forms.TabControl();
            this.tabPageHJizen1 = new System.Windows.Forms.TabPage();
            this.GroupBox5 = new System.Windows.Forms.GroupBox();
            this.pnlHJizenTyukan = new System.Windows.Forms.Panel();
            this.txtMidDirPath = new System.Windows.Forms.TextBox();
            this.chkHJizenTyukan = new System.Windows.Forms.CheckBox();
            this.Label355 = new System.Windows.Forms.Label();
            this.lblCautionDescription = new System.Windows.Forms.Label();
            this.lblJizenDescription1 = new System.Windows.Forms.Label();
            this.tabPageJigo = new System.Windows.Forms.TabPage();
            this.Label245 = new System.Windows.Forms.Label();
            this.lblLine3 = new System.Windows.Forms.Label();
            this.lblHidden4 = new System.Windows.Forms.Label();
            this.tabCtrlJigo = new System.Windows.Forms.TabControl();
            this.tabPageJigo1 = new System.Windows.Forms.TabPage();
            this.grpJigoDonyuji = new System.Windows.Forms.GroupBox();
            this.pnlJigoCmtSoKotiku = new System.Windows.Forms.Panel();
            this.Label109 = new System.Windows.Forms.Label();
            this.Label108 = new System.Windows.Forms.Label();
            this.Label142 = new System.Windows.Forms.Label();
            this.pnlJigoCmtNkNyuryoku = new System.Windows.Forms.Panel();
            this.Label99 = new System.Windows.Forms.Label();
            this.Label111 = new System.Windows.Forms.Label();
            this.Label110 = new System.Windows.Forms.Label();
            this.pnlJigoCmtSqKotiku = new System.Windows.Forms.Panel();
            this.Label113 = new System.Windows.Forms.Label();
            this.Label112 = new System.Windows.Forms.Label();
            this.Label24 = new System.Windows.Forms.Label();
            this.grpJigoUserSagyo = new System.Windows.Forms.GroupBox();
            this.pnlHJigoCmtSyudo = new System.Windows.Forms.Panel();
            this.Label136 = new System.Windows.Forms.Label();
            this.Label317 = new System.Windows.Forms.Label();
            this.Label318 = new System.Windows.Forms.Label();
            this.lblJizenPageNum1 = new System.Windows.Forms.Label();
            this.lblJigoDescription1 = new System.Windows.Forms.Label();
            this.tabPageHojyo = new System.Windows.Forms.TabPage();
            this.tabPageHajimeni = new System.Windows.Forms.TabPage();
            this.lblDatacvHajimeniDescription1 = new System.Windows.Forms.Label();
            this.lblHDatacvHajimeniLabel = new System.Windows.Forms.Label();
            this.btnDoui = new System.Windows.Forms.Button();
            this.GroupBox10 = new System.Windows.Forms.GroupBox();
            this.pnlDcFstCmtH99 = new System.Windows.Forms.Panel();
            this.Label301 = new System.Windows.Forms.Label();
            this.Label302 = new System.Windows.Forms.Label();
            this.pnlDcFstCmtH01 = new System.Windows.Forms.Panel();
            this.Label158 = new System.Windows.Forms.Label();
            this.Label159 = new System.Windows.Forms.Label();
            this.pnlDcFstCmtH02 = new System.Windows.Forms.Panel();
            this.Label27 = new System.Windows.Forms.Label();
            this.Label28 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt11 = new System.Windows.Forms.Panel();
            this.Label21 = new System.Windows.Forms.Label();
            this.Label25 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt06 = new System.Windows.Forms.Panel();
            this.Label364 = new System.Windows.Forms.Label();
            this.Label369 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt05 = new System.Windows.Forms.Panel();
            this.Label362 = new System.Windows.Forms.Label();
            this.Label363 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt07 = new System.Windows.Forms.Panel();
            this.Label122 = new System.Windows.Forms.Label();
            this.Label104 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt10 = new System.Windows.Forms.Panel();
            this.Label123 = new System.Windows.Forms.Label();
            this.Label105 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt02 = new System.Windows.Forms.Panel();
            this.Label127 = new System.Windows.Forms.Label();
            this.Label133 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt08 = new System.Windows.Forms.Panel();
            this.Label126 = new System.Windows.Forms.Label();
            this.Label132 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt04 = new System.Windows.Forms.Panel();
            this.Label119 = new System.Windows.Forms.Label();
            this.Label116 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt03 = new System.Windows.Forms.Panel();
            this.Label97 = new System.Windows.Forms.Label();
            this.Label101 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt01 = new System.Windows.Forms.Panel();
            this.Label129 = new System.Windows.Forms.Label();
            this.Label135 = new System.Windows.Forms.Label();
            this.pnlDcFstCmt09 = new System.Windows.Forms.Panel();
            this.Label326 = new System.Windows.Forms.Label();
            this.Label325 = new System.Windows.Forms.Label();
            this.PictureBox14 = new System.Windows.Forms.PictureBox();
            this.tabPageSelect = new System.Windows.Forms.TabPage();
            this.grpExistMidToBaseMid = new System.Windows.Forms.GroupBox();
            this.btnAllChk = new System.Windows.Forms.Button();
            this.lblDatacvSelectCaution = new System.Windows.Forms.Label();
            this.lblLine2 = new System.Windows.Forms.Label();
            this.lblHidden2 = new System.Windows.Forms.Label();
            this.grpMiddleFile = new System.Windows.Forms.GroupBox();
            this.Label349 = new System.Windows.Forms.Label();
            this.Label328 = new System.Windows.Forms.Label();
            this.Label124 = new System.Windows.Forms.Label();
            this.btnMidFileCheck = new System.Windows.Forms.Button();
            this.btnMidDirLogSeach = new System.Windows.Forms.Button();
            this.txtMidDirLogPath = new System.Windows.Forms.TextBox();
            this.lblMiddleFileLog = new System.Windows.Forms.Label();
            this.tabCtrlCVItem = new System.Windows.Forms.TabControl();
            this.tabPageKizon110 = new System.Windows.Forms.TabPage();
            this.grpKizon1 = new System.Windows.Forms.GroupBox();
            this.pnlKiMstHendo = new System.Windows.Forms.Panel();
            this.chkKiMstHendo = new System.Windows.Forms.CheckBox();
            this.lblKiMstHendoCnt = new System.Windows.Forms.Label();
            this.Label230 = new System.Windows.Forms.Label();
            this.Label82 = new System.Windows.Forms.Label();
            this.Label81 = new System.Windows.Forms.Label();
            this.pnlKiMstTokuyaku = new System.Windows.Forms.Panel();
            this.chkKiMstTokuyaku = new System.Windows.Forms.CheckBox();
            this.lblKiMstTokuyakuCnt = new System.Windows.Forms.Label();
            this.Label223 = new System.Windows.Forms.Label();
            this.Label222 = new System.Windows.Forms.Label();
            this.Label78 = new System.Windows.Forms.Label();
            this.pnlKiMstKasyoClaimrui = new System.Windows.Forms.Panel();
            this.chkKiMstKasyoClaimrui = new System.Windows.Forms.CheckBox();
            this.lblKiMstKasyoClaimruiCnt = new System.Windows.Forms.Label();
            this.Label228 = new System.Windows.Forms.Label();
            this.Label227 = new System.Windows.Forms.Label();
            this.Label226 = new System.Windows.Forms.Label();
            this.pnlKiMstTitle = new System.Windows.Forms.Panel();
            this.Label346 = new System.Windows.Forms.Label();
            this.Label236 = new System.Windows.Forms.Label();
            this.Label233 = new System.Windows.Forms.Label();
            this.lblKiMstGazotitleCnt = new System.Windows.Forms.Label();
            this.Label347 = new System.Windows.Forms.Label();
            this.lblKiMstBikotitleCnt = new System.Windows.Forms.Label();
            this.Label235 = new System.Windows.Forms.Label();
            this.chkKiMstTitle = new System.Windows.Forms.CheckBox();
            this.lblKiMstKagititleCnt = new System.Windows.Forms.Label();
            this.Label232 = new System.Windows.Forms.Label();
            this.Label80 = new System.Windows.Forms.Label();
            this.pnlKiMstArea = new System.Windows.Forms.Panel();
            this.chkKiMstArea = new System.Windows.Forms.CheckBox();
            this.lblKiMstAreaCnt = new System.Windows.Forms.Label();
            this.Label216 = new System.Windows.Forms.Label();
            this.Label77 = new System.Windows.Forms.Label();
            this.Label76 = new System.Windows.Forms.Label();
            this.pnlKiMstSchool = new System.Windows.Forms.Panel();
            this.chkKiMstSchool = new System.Windows.Forms.CheckBox();
            this.lblKiMstSchoolCnt = new System.Windows.Forms.Label();
            this.Label214 = new System.Windows.Forms.Label();
            this.Label213 = new System.Windows.Forms.Label();
            this.Label75 = new System.Windows.Forms.Label();
            this.pnlKiMstHokenrui = new System.Windows.Forms.Panel();
            this.chkKiMstHokenrui = new System.Windows.Forms.CheckBox();
            this.lblKiMstHokenruiCnt = new System.Windows.Forms.Label();
            this.Label220 = new System.Windows.Forms.Label();
            this.Label219 = new System.Windows.Forms.Label();
            this.Label218 = new System.Windows.Forms.Label();
            this.pnlKiMstBus = new System.Windows.Forms.Panel();
            this.chkKiMstBus = new System.Windows.Forms.CheckBox();
            this.lblKiMstBusCnt = new System.Windows.Forms.Label();
            this.Label211 = new System.Windows.Forms.Label();
            this.Label181 = new System.Windows.Forms.Label();
            this.Label210 = new System.Windows.Forms.Label();
            this.lblSelectPageCnt1 = new System.Windows.Forms.Label();
            this.tabPageKizon120 = new System.Windows.Forms.TabPage();
            this.grpKizon3 = new System.Windows.Forms.GroupBox();
            this.pnlKiGySyuzenBase = new System.Windows.Forms.Panel();
            this.chkKiGySyuzenBase = new System.Windows.Forms.CheckBox();
            this.lblKiGySyuzenBaseCnt = new System.Windows.Forms.Label();
            this.Label241 = new System.Windows.Forms.Label();
            this.Label240 = new System.Windows.Forms.Label();
            this.pnlKiGyYatinhosyoBase = new System.Windows.Forms.Panel();
            this.chkKiGyYatinhosyoBase = new System.Windows.Forms.CheckBox();
            this.lblKiGyYatinhosyoBaseCnt = new System.Windows.Forms.Label();
            this.Label250 = new System.Windows.Forms.Label();
            this.Label249 = new System.Windows.Forms.Label();
            this.pnlKiGyLifelineBase = new System.Windows.Forms.Panel();
            this.chkKiGyLifelineBase = new System.Windows.Forms.CheckBox();
            this.lblKiGyLifelineBaseCnt = new System.Windows.Forms.Label();
            this.Label244 = new System.Windows.Forms.Label();
            this.Label243 = new System.Windows.Forms.Label();
            this.pnlKiGyHokenBase = new System.Windows.Forms.Panel();
            this.chkKiGyHokenBase = new System.Windows.Forms.CheckBox();
            this.lblKiGyHokenBaseCnt = new System.Windows.Forms.Label();
            this.Label247 = new System.Windows.Forms.Label();
            this.Label246 = new System.Windows.Forms.Label();
            this.pnlKiGySisetuBase = new System.Windows.Forms.Panel();
            this.chkKiGySisetuBase = new System.Windows.Forms.CheckBox();
            this.lblKiGySisetuBaseCnt = new System.Windows.Forms.Label();
            this.Label253 = new System.Windows.Forms.Label();
            this.Label252 = new System.Windows.Forms.Label();
            this.pnlKiGyCyukaiBase = new System.Windows.Forms.Panel();
            this.chkKiGyCyukaiBase = new System.Windows.Forms.CheckBox();
            this.lblKiGyCyukaiBaseCnt = new System.Windows.Forms.Label();
            this.Label238 = new System.Windows.Forms.Label();
            this.Label87 = new System.Windows.Forms.Label();
            this.pnlKiGySekoBase = new System.Windows.Forms.Panel();
            this.chkKiGySekoBase = new System.Windows.Forms.CheckBox();
            this.lblKiGySekoBaseCnt = new System.Windows.Forms.Label();
            this.Label256 = new System.Windows.Forms.Label();
            this.Label255 = new System.Windows.Forms.Label();
            this.grpKizon5 = new System.Windows.Forms.GroupBox();
            this.pnlKiOw = new System.Windows.Forms.Panel();
            this.chkKiOwBase = new System.Windows.Forms.CheckBox();
            this.lblKiOwBaseCnt = new System.Windows.Forms.Label();
            this.Label261 = new System.Windows.Forms.Label();
            this.Label268 = new System.Windows.Forms.Label();
            this.Label267 = new System.Windows.Forms.Label();
            this.pnlKiJisya = new System.Windows.Forms.Panel();
            this.chkKiJisyaBase = new System.Windows.Forms.CheckBox();
            this.lblKiJisyaBaseCnt = new System.Windows.Forms.Label();
            this.Label265 = new System.Windows.Forms.Label();
            this.Label264 = new System.Windows.Forms.Label();
            this.Label263 = new System.Windows.Forms.Label();
            this.pnlKiSyskanriBase = new System.Windows.Forms.Panel();
            this.chkKiSyskanriBase = new System.Windows.Forms.CheckBox();
            this.lblKiSyskanriBaseCnt = new System.Windows.Forms.Label();
            this.Label260 = new System.Windows.Forms.Label();
            this.Label259 = new System.Windows.Forms.Label();
            this.Label258 = new System.Windows.Forms.Label();
            this.lblSelectPageCnt2 = new System.Windows.Forms.Label();
            this.tabPageKizon130 = new System.Windows.Forms.TabPage();
            this.grpKizon2 = new System.Windows.Forms.GroupBox();
            this.Label359 = new System.Windows.Forms.Label();
            this.pnlRendo = new System.Windows.Forms.Panel();
            this.Label83 = new System.Windows.Forms.Label();
            this.txtRendoID = new System.Windows.Forms.TextBox();
            this.Label121 = new System.Windows.Forms.Label();
            this.chkKiRendoBase = new System.Windows.Forms.CheckBox();
            this.lblKiRendoBaseCnt = new System.Windows.Forms.Label();
            this.Label314 = new System.Windows.Forms.Label();
            this.Label316 = new System.Windows.Forms.Label();
            this.pnlKiSq = new System.Windows.Forms.Panel();
            this.Label308 = new System.Windows.Forms.Label();
            this.lblKiSqOwKojoBaseCnt = new System.Windows.Forms.Label();
            this.Label311 = new System.Windows.Forms.Label();
            this.Label290 = new System.Windows.Forms.Label();
            this.Label162 = new System.Windows.Forms.Label();
            this.chkKiSqBase = new System.Windows.Forms.CheckBox();
            this.lblKiSqMiBaseCnt = new System.Windows.Forms.Label();
            this.Label289 = new System.Windows.Forms.Label();
            this.Label288 = new System.Windows.Forms.Label();
            this.lblKiSqAzBaseCnt = new System.Windows.Forms.Label();
            this.Label160 = new System.Windows.Forms.Label();
            this.pnlSzen = new System.Windows.Forms.Panel();
            this.chkKiSzenBase = new System.Windows.Forms.CheckBox();
            this.lblKiSzenBaseCnt = new System.Windows.Forms.Label();
            this.Label295 = new System.Windows.Forms.Label();
            this.Label294 = new System.Windows.Forms.Label();
            this.Label174 = new System.Windows.Forms.Label();
            this.pnlKiClaim = new System.Windows.Forms.Panel();
            this.chkKiClaimBase = new System.Windows.Forms.CheckBox();
            this.lblKiClaimBaseCnt = new System.Windows.Forms.Label();
            this.Label292 = new System.Windows.Forms.Label();
            this.Label180 = new System.Windows.Forms.Label();
            this.pnlKiKy = new System.Windows.Forms.Panel();
            this.chkKiKyBase = new System.Windows.Forms.CheckBox();
            this.lblKiKyBaseCnt = new System.Windows.Forms.Label();
            this.Label286 = new System.Windows.Forms.Label();
            this.Label285 = new System.Windows.Forms.Label();
            this.Label284 = new System.Windows.Forms.Label();
            this.pnlKiKys = new System.Windows.Forms.Panel();
            this.chkKiKysBase = new System.Windows.Forms.CheckBox();
            this.lblKiKysBaseCnt = new System.Windows.Forms.Label();
            this.Label282 = new System.Windows.Forms.Label();
            this.Label281 = new System.Windows.Forms.Label();
            this.Label280 = new System.Windows.Forms.Label();
            this.pnlKiHy = new System.Windows.Forms.Panel();
            this.chkKiHyBase = new System.Windows.Forms.CheckBox();
            this.lblKiHyBaseCnt = new System.Windows.Forms.Label();
            this.Label275 = new System.Windows.Forms.Label();
            this.Label270 = new System.Windows.Forms.Label();
            this.chkKiHySetubi = new System.Windows.Forms.CheckBox();
            this.lblKiHySetubiCnt = new System.Windows.Forms.Label();
            this.Label278 = new System.Windows.Forms.Label();
            this.Label277 = new System.Windows.Forms.Label();
            this.pnlKiBk = new System.Windows.Forms.Panel();
            this.chkKiBkBase = new System.Windows.Forms.CheckBox();
            this.lblKiBkBaseCnt = new System.Windows.Forms.Label();
            this.Label273 = new System.Windows.Forms.Label();
            this.Label271 = new System.Windows.Forms.Label();
            this.grpKizonKagi = new System.Windows.Forms.GroupBox();
            this.optKiKyKagi = new System.Windows.Forms.RadioButton();
            this.optKiHyKagi = new System.Windows.Forms.RadioButton();
            this.Label86 = new System.Windows.Forms.Label();
            this.Label272 = new System.Windows.Forms.Label();
            this.tabPageBase110 = new System.Windows.Forms.TabPage();
            this.lblRelItemInfo = new System.Windows.Forms.Label();
            this.grpMst = new System.Windows.Forms.GroupBox();
            this.chkMstGazotitle = new System.Windows.Forms.CheckBox();
            this.chkMstBikolst = new System.Windows.Forms.CheckBox();
            this.chkMstBikotitle = new System.Windows.Forms.CheckBox();
            this.chkMstHendoitiran = new System.Windows.Forms.CheckBox();
            this.chkMstHendo = new System.Windows.Forms.CheckBox();
            this.chkMstKasyoClaimrui = new System.Windows.Forms.CheckBox();
            this.chkMstBus = new System.Windows.Forms.CheckBox();
            this.chkMstSchool = new System.Windows.Forms.CheckBox();
            this.chkMstTokuyaku = new System.Windows.Forms.CheckBox();
            this.chkMstHokenrui = new System.Windows.Forms.CheckBox();
            this.chkMstBusKotu = new System.Windows.Forms.CheckBox();
            this.chkMstKagititle = new System.Windows.Forms.CheckBox();
            this.chkMstArea = new System.Windows.Forms.CheckBox();
            this.tabPageBase120 = new System.Windows.Forms.TabPage();
            this.grpGy = new System.Windows.Forms.GroupBox();
            this.chkGySyuzenMemo = new System.Windows.Forms.CheckBox();
            this.chkGyHokenMemo = new System.Windows.Forms.CheckBox();
            this.chkGySyuzenKoza = new System.Windows.Forms.CheckBox();
            this.chkGySekoBase = new System.Windows.Forms.CheckBox();
            this.chkGySisetuBase = new System.Windows.Forms.CheckBox();
            this.chkGySyuzenBase = new System.Windows.Forms.CheckBox();
            this.chkGyYatinhosyoMemo = new System.Windows.Forms.CheckBox();
            this.chkGyYatinhosyoBase = new System.Windows.Forms.CheckBox();
            this.chkGyCyukaiKoza = new System.Windows.Forms.CheckBox();
            this.chkGyCyukaiMemo = new System.Windows.Forms.CheckBox();
            this.chkGyHokenBase = new System.Windows.Forms.CheckBox();
            this.chkGyLifelineBase = new System.Windows.Forms.CheckBox();
            this.chkGyHokenKoza = new System.Windows.Forms.CheckBox();
            this.chkGyCyukaiBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase130 = new System.Windows.Forms.TabPage();
            this.grpKys = new System.Windows.Forms.GroupBox();
            this.chkKysHosyonin = new System.Windows.Forms.CheckBox();
            this.chkKysSyogoKana = new System.Windows.Forms.CheckBox();
            this.chkKysKoza = new System.Windows.Forms.CheckBox();
            this.chkKysMemo = new System.Windows.Forms.CheckBox();
            this.chkKysBase = new System.Windows.Forms.CheckBox();
            this.grpOw = new System.Windows.Forms.GroupBox();
            this.chkOwEvent = new System.Windows.Forms.CheckBox();
            this.chkOwKoza = new System.Windows.Forms.CheckBox();
            this.chkOwMemo = new System.Windows.Forms.CheckBox();
            this.chkOwBase = new System.Windows.Forms.CheckBox();
            this.grpJisya = new System.Windows.Forms.GroupBox();
            this.chkFBANSERSetuzoku = new System.Windows.Forms.CheckBox();
            this.chkMstANSERArea = new System.Windows.Forms.CheckBox();
            this.chkMstANSERAccpoint = new System.Windows.Forms.CheckBox();
            this.chkJisyaTanto = new System.Windows.Forms.CheckBox();
            this.chkMstYatinKoza = new System.Windows.Forms.CheckBox();
            this.chkJisyaKoza = new System.Windows.Forms.CheckBox();
            this.chkJisyaBase = new System.Windows.Forms.CheckBox();
            this.chkFBFuriirai = new System.Windows.Forms.CheckBox();
            this.chkFBNsSyutoku = new System.Windows.Forms.CheckBox();
            this.chkJisyaMemo = new System.Windows.Forms.CheckBox();
            this.chkFBKozafurikae = new System.Windows.Forms.CheckBox();
            this.chkFBFuritesuryo = new System.Windows.Forms.CheckBox();
            this.tabPageBase140 = new System.Windows.Forms.TabPage();
            this.grpBk = new System.Windows.Forms.GroupBox();
            this.chkOpKys = new System.Windows.Forms.CheckBox();
            this.chkOpOw = new System.Windows.Forms.CheckBox();
            this.grpKagiSelect = new System.Windows.Forms.GroupBox();
            this.optKyKagi = new System.Windows.Forms.RadioButton();
            this.optHyKagi = new System.Windows.Forms.RadioButton();
            this.chkBkSyo = new System.Windows.Forms.CheckBox();
            this.chkBkHendo = new System.Windows.Forms.CheckBox();
            this.chkBkKinrincyusyajo = new System.Windows.Forms.CheckBox();
            this.chkBkSansyofile = new System.Windows.Forms.CheckBox();
            this.chkBkSzeniji = new System.Windows.Forms.CheckBox();
            this.chkBkSyuhen = new System.Windows.Forms.CheckBox();
            this.chkBkSetudo = new System.Windows.Forms.CheckBox();
            this.chkBkKotu = new System.Windows.Forms.CheckBox();
            this.chkBkKenri = new System.Windows.Forms.CheckBox();
            this.chkBkGomi = new System.Windows.Forms.CheckBox();
            this.chkBkSyosai = new System.Windows.Forms.CheckBox();
            this.chkBkKagi = new System.Windows.Forms.CheckBox();
            this.chkBkMemo = new System.Windows.Forms.CheckBox();
            this.chkBkBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase150 = new System.Windows.Forms.TabPage();
            this.grpHy = new System.Windows.Forms.GroupBox();
            this.chkHySyo = new System.Windows.Forms.CheckBox();
            this.chkHySansyofile = new System.Windows.Forms.CheckBox();
            this.chkHyGenjotanka = new System.Windows.Forms.CheckBox();
            this.chkHyKenri = new System.Windows.Forms.CheckBox();
            this.chkHyConfirm = new System.Windows.Forms.CheckBox();
            this.chkHyHendo = new System.Windows.Forms.CheckBox();
            this.chkHyCommonsalespoint = new System.Windows.Forms.CheckBox();
            this.chkHyMenseki = new System.Windows.Forms.CheckBox();
            this.chkHyNkinkomk = new System.Windows.Forms.CheckBox();
            this.chkHyMadoriutiwake = new System.Windows.Forms.CheckBox();
            this.chkHySzeniji = new System.Windows.Forms.CheckBox();
            this.chkHyTokuyaku = new System.Windows.Forms.CheckBox();
            this.chkHyParking = new System.Windows.Forms.CheckBox();
            this.chkHySyosai = new System.Windows.Forms.CheckBox();
            this.chkHyMemo = new System.Windows.Forms.CheckBox();
            this.chkHyKagi = new System.Windows.Forms.CheckBox();
            this.chkHySetubi = new System.Windows.Forms.CheckBox();
            this.chkHyBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase160 = new System.Windows.Forms.TabPage();
            this.grpSorule = new System.Windows.Forms.GroupBox();
            this.chkSoruleSosaki = new System.Windows.Forms.CheckBox();
            this.chkSoruleKojo = new System.Windows.Forms.CheckBox();
            this.chkSoruleNkin = new System.Windows.Forms.CheckBox();
            this.chkSoruleBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase170 = new System.Windows.Forms.TabPage();
            this.grpKy = new System.Windows.Forms.GroupBox();
            this.chkKySzenmeisai = new System.Windows.Forms.CheckBox();
            this.chkKySzen = new System.Windows.Forms.CheckBox();
            this.chkKyKai = new System.Windows.Forms.CheckBox();
            this.CheckBox14 = new System.Windows.Forms.CheckBox();
            this.CheckBox24 = new System.Windows.Forms.CheckBox();
            this.CheckBox23 = new System.Windows.Forms.CheckBox();
            this.chkKyHosyonin = new System.Windows.Forms.CheckBox();
            this.chkKyMemo = new System.Windows.Forms.CheckBox();
            this.chkKyTokuyaku = new System.Windows.Forms.CheckBox();
            this.CheckBox18 = new System.Windows.Forms.CheckBox();
            this.chkKySorule = new System.Windows.Forms.CheckBox();
            this.CheckBox16 = new System.Windows.Forms.CheckBox();
            this.CheckBox15 = new System.Windows.Forms.CheckBox();
            this.chkKyNyukyo = new System.Windows.Forms.CheckBox();
            this.chkKyNkinkomkNx = new System.Windows.Forms.CheckBox();
            this.chkKyNkinkomk = new System.Windows.Forms.CheckBox();
            this.chkKyHoken = new System.Windows.Forms.CheckBox();
            this.chkKyKojoRule = new System.Windows.Forms.CheckBox();
            this.chkKyKys = new System.Windows.Forms.CheckBox();
            this.CheckBox5 = new System.Windows.Forms.CheckBox();
            this.chkKyRireki = new System.Windows.Forms.CheckBox();
            this.chkKyHendo = new System.Windows.Forms.CheckBox();
            this.chkKyCar = new System.Windows.Forms.CheckBox();
            this.chkKyBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase180 = new System.Windows.Forms.TabPage();
            this.grpSq = new System.Windows.Forms.GroupBox();
            this.chkSqSqKojo = new System.Windows.Forms.CheckBox();
            this.chkSqKoteiKojo = new System.Windows.Forms.CheckBox();
            this.chkSqHendokensin = new System.Windows.Forms.CheckBox();
            this.chkSqUnyotaino = new System.Windows.Forms.CheckBox();
            this.chkSqSq = new System.Windows.Forms.CheckBox();
            this.chkSqKajyo = new System.Windows.Forms.CheckBox();
            this.tabPageBase190 = new System.Windows.Forms.TabPage();
            this.grpSzen = new System.Windows.Forms.GroupBox();
            this.chkSzenRelfile = new System.Windows.Forms.CheckBox();
            this.chkSzenClaim = new System.Windows.Forms.CheckBox();
            this.chkSzenMemo = new System.Windows.Forms.CheckBox();
            this.chkSzenSzen = new System.Windows.Forms.CheckBox();
            this.chkSzenSzenmeisai = new System.Windows.Forms.CheckBox();
            this.chkSzenBase = new System.Windows.Forms.CheckBox();
            this.grpClaim = new System.Windows.Forms.GroupBox();
            this.chkClaimTaiorireki = new System.Windows.Forms.CheckBox();
            this.chkClaimRelfile = new System.Windows.Forms.CheckBox();
            this.chkClaimBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase200 = new System.Windows.Forms.TabPage();
            this.grpSyskanri = new System.Windows.Forms.GroupBox();
            this.chkSyskanriNkinkomkmerge = new System.Windows.Forms.CheckBox();
            this.chkSyskanriHenkanmoji = new System.Windows.Forms.CheckBox();
            this.chkSyskanriZei = new System.Windows.Forms.CheckBox();
            this.chkSyskanriBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase210 = new System.Windows.Forms.TabPage();
            this.grpRendo = new System.Windows.Forms.GroupBox();
            this.chkRendoMapdisp = new System.Windows.Forms.CheckBox();
            this.chkRendoBtoBgroup = new System.Windows.Forms.CheckBox();
            this.chkRendoHysosin = new System.Windows.Forms.CheckBox();
            this.chkRendoHyrui = new System.Windows.Forms.CheckBox();
            this.chkRendoKokokuSuumo = new System.Windows.Forms.CheckBox();
            this.chkRendoKokokuAthome = new System.Windows.Forms.CheckBox();
            this.chkRendoKokokuHomes = new System.Windows.Forms.CheckBox();
            this.chkRendoKokokuJisyaweb = new System.Windows.Forms.CheckBox();
            this.chkRendoSosinSuumo = new System.Windows.Forms.CheckBox();
            this.chkRendoSosinAthome = new System.Windows.Forms.CheckBox();
            this.chkRendoSosinHomes = new System.Windows.Forms.CheckBox();
            this.chkRendoSosinJisyaweb = new System.Windows.Forms.CheckBox();
            this.chkRendoSosinBase = new System.Windows.Forms.CheckBox();
            this.tabPageBase900 = new System.Windows.Forms.TabPage();
            this.Label22 = new System.Windows.Forms.Label();
            this.tabPageHanyo110 = new System.Windows.Forms.TabPage();
            this.grpHMst = new System.Windows.Forms.GroupBox();
            this.Label138 = new System.Windows.Forms.Label();
            this.Label150 = new System.Windows.Forms.Label();
            this.Label151 = new System.Windows.Forms.Label();
            this.Label152 = new System.Windows.Forms.Label();
            this.Label153 = new System.Windows.Forms.Label();
            this.Label154 = new System.Windows.Forms.Label();
            this.Label155 = new System.Windows.Forms.Label();
            this.Label156 = new System.Windows.Forms.Label();
            this.Label157 = new System.Windows.Forms.Label();
            this.CheckBox2 = new System.Windows.Forms.CheckBox();
            this.CheckBox3 = new System.Windows.Forms.CheckBox();
            this.CheckBox6 = new System.Windows.Forms.CheckBox();
            this.CheckBox9 = new System.Windows.Forms.CheckBox();
            this.CheckBox10 = new System.Windows.Forms.CheckBox();
            this.CheckBox11 = new System.Windows.Forms.CheckBox();
            this.CheckBox12 = new System.Windows.Forms.CheckBox();
            this.CheckBox13 = new System.Windows.Forms.CheckBox();
            this.CheckBox17 = new System.Windows.Forms.CheckBox();
            this.tabPageHanyo120 = new System.Windows.Forms.TabPage();
            this.grpHGy = new System.Windows.Forms.GroupBox();
            this.CheckBox4 = new System.Windows.Forms.CheckBox();
            this.CheckBox30 = new System.Windows.Forms.CheckBox();
            this.CheckBox31 = new System.Windows.Forms.CheckBox();
            this.CheckBox32 = new System.Windows.Forms.CheckBox();
            this.CheckBox33 = new System.Windows.Forms.CheckBox();
            this.CheckBox34 = new System.Windows.Forms.CheckBox();
            this.CheckBox35 = new System.Windows.Forms.CheckBox();
            this.grpHOw = new System.Windows.Forms.GroupBox();
            this.CheckBox1 = new System.Windows.Forms.CheckBox();
            this.CheckBox26 = new System.Windows.Forms.CheckBox();
            this.CheckBox27 = new System.Windows.Forms.CheckBox();
            this.CheckBox28 = new System.Windows.Forms.CheckBox();
            this.CheckBox29 = new System.Windows.Forms.CheckBox();
            this.grpHJisya = new System.Windows.Forms.GroupBox();
            this.CheckBox19 = new System.Windows.Forms.CheckBox();
            this.CheckBox20 = new System.Windows.Forms.CheckBox();
            this.CheckBox21 = new System.Windows.Forms.CheckBox();
            this.CheckBox22 = new System.Windows.Forms.CheckBox();
            this.CheckBox25 = new System.Windows.Forms.CheckBox();
            this.tabPageHanyo130 = new System.Windows.Forms.TabPage();
            this.tabPageHanyo140 = new System.Windows.Forms.TabPage();
            this.tabPageHanyo150 = new System.Windows.Forms.TabPage();
            this.TabPage1 = new System.Windows.Forms.TabPage();
            this.grpKizon4 = new System.Windows.Forms.GroupBox();
            this.lblSelectPageCnt3 = new System.Windows.Forms.Label();
            this.Label26 = new System.Windows.Forms.Label();
            this.Label73 = new System.Windows.Forms.Label();
            this.Label70 = new System.Windows.Forms.Label();
            this.Label71 = new System.Windows.Forms.Label();
            this.Label72 = new System.Windows.Forms.Label();
            this.Label74 = new System.Windows.Forms.Label();
            this.chkHSongai = new System.Windows.Forms.CheckBox();
            this.chkHJisyaKoza = new System.Windows.Forms.CheckBox();
            this.chkHSetubi = new System.Windows.Forms.CheckBox();
            this.chkHDataFmt = new System.Windows.Forms.CheckBox();
            this.Label69 = new System.Windows.Forms.Label();
            this.Label68 = new System.Windows.Forms.Label();
            this.Label67 = new System.Windows.Forms.Label();
            this.Label66 = new System.Windows.Forms.Label();
            this.Label65 = new System.Windows.Forms.Label();
            this.Label64 = new System.Windows.Forms.Label();
            this.Label63 = new System.Windows.Forms.Label();
            this.Label62 = new System.Windows.Forms.Label();
            this.Label61 = new System.Windows.Forms.Label();
            this.Label60 = new System.Windows.Forms.Label();
            this.Label59 = new System.Windows.Forms.Label();
            this.chkHNkinKomk = new System.Windows.Forms.CheckBox();
            this.chkHTaiyo = new System.Windows.Forms.CheckBox();
            this.chkHKozo = new System.Windows.Forms.CheckBox();
            this.chkHKozaSyu = new System.Windows.Forms.CheckBox();
            this.chkHEki = new System.Windows.Forms.CheckBox();
            this.chkHEnsen = new System.Windows.Forms.CheckBox();
            this.chkHKinyuSiten = new System.Windows.Forms.CheckBox();
            this.chkHYouto = new System.Windows.Forms.CheckBox();
            this.chkHBkBunrui = new System.Windows.Forms.CheckBox();
            this.chkHKinyu = new System.Windows.Forms.CheckBox();
            this.chkHKyBunrui = new System.Windows.Forms.CheckBox();
            this.chkHKagi = new System.Windows.Forms.CheckBox();
            this.chkHHouKenri = new System.Windows.Forms.CheckBox();
            this.chkHNkinKbn = new System.Windows.Forms.CheckBox();
            this.chkHHyBunrui = new System.Windows.Forms.CheckBox();
            this.pnlRekiClear = new System.Windows.Forms.Panel();
            this.btnRekiClear = new System.Windows.Forms.Button();
            this.Label79 = new System.Windows.Forms.Label();
            this.Label298 = new System.Windows.Forms.Label();
            this.lblDatacvSelectDescription1 = new System.Windows.Forms.Label();
            this.tabPageJikko = new System.Windows.Forms.TabPage();
            this.lblDatacvJikkoCaution = new System.Windows.Forms.Label();
            this.lblDatacvJikkoDescription1 = new System.Windows.Forms.Label();
            this.grpTotalProcess = new System.Windows.Forms.GroupBox();
            this.lblCVItem = new System.Windows.Forms.Label();
            this.txtPartialSituation = new System.Windows.Forms.TextBox();
            this.Label31 = new System.Windows.Forms.Label();
            this.lblPgbPartial = new System.Windows.Forms.Label();
            this.lblPgbTotal = new System.Windows.Forms.Label();
            this.txtTotalSituation = new System.Windows.Forms.TextBox();
            this.pgbpartial = new System.Windows.Forms.ProgressBar();
            this.pgbTotal = new System.Windows.Forms.ProgressBar();
            this.Label32 = new System.Windows.Forms.Label();
            this.lblTotalSituation = new System.Windows.Forms.Label();
            this.tabPageEndOK = new System.Windows.Forms.TabPage();
            this.lblDatacvEndDescription1 = new System.Windows.Forms.Label();
            this.lblHDatacvEndLabel = new System.Windows.Forms.Label();
            this.PictureBox15 = new System.Windows.Forms.PictureBox();
            this.tabPageEndError = new System.Windows.Forms.TabPage();
            this.GroupBox23 = new System.Windows.Forms.GroupBox();
            this.Label163 = new System.Windows.Forms.Label();
            this.lblDatacvErrorDescription1 = new System.Windows.Forms.Label();
            this.GroupBox19 = new System.Windows.Forms.GroupBox();
            this.Label170 = new System.Windows.Forms.Label();
            this.GroupBox20 = new System.Windows.Forms.GroupBox();
            this.Label176 = new System.Windows.Forms.Label();
            this.lblDatacvErrorLabel = new System.Windows.Forms.Label();
            this.PictureBox23 = new System.Windows.Forms.PictureBox();
            this.tabPageEndCancel = new System.Windows.Forms.TabPage();
            this.lblDatacvCancelDescription1 = new System.Windows.Forms.Label();
            this.GroupBox21 = new System.Windows.Forms.GroupBox();
            this.Label186 = new System.Windows.Forms.Label();
            this.GroupBox22 = new System.Windows.Forms.GroupBox();
            this.Label192 = new System.Windows.Forms.Label();
            this.lblDatacvCancelLabel = new System.Windows.Forms.Label();
            this.PictureBox25 = new System.Windows.Forms.PictureBox();
            this.tabPageIkkatu = new System.Windows.Forms.TabPage();
            this.tabPageHanyoJizen = new System.Windows.Forms.TabPage();
            this.btnDevTabChange = new System.Windows.Forms.Button();
            this.Label393 = new System.Windows.Forms.Label();
            this.Label392 = new System.Windows.Forms.Label();
            this.CheckBox51 = new System.Windows.Forms.CheckBox();
            this.Label391 = new System.Windows.Forms.Label();
            this.CheckBox50 = new System.Windows.Forms.CheckBox();
            this.Panel24 = new System.Windows.Forms.Panel();
            this.Label390 = new System.Windows.Forms.Label();
            this.Label389 = new System.Windows.Forms.Label();
            this.Label388 = new System.Windows.Forms.Label();
            this.Label387 = new System.Windows.Forms.Label();
            this.Panel23 = new System.Windows.Forms.Panel();
            this.Label386 = new System.Windows.Forms.Label();
            this.Label385 = new System.Windows.Forms.Label();
            this.Label384 = new System.Windows.Forms.Label();
            this.Label383 = new System.Windows.Forms.Label();
            this.Label382 = new System.Windows.Forms.Label();
            this.CheckBox49 = new System.Windows.Forms.CheckBox();
            this.Panel22 = new System.Windows.Forms.Panel();
            this.Label381 = new System.Windows.Forms.Label();
            this.Label380 = new System.Windows.Forms.Label();
            this.Label379 = new System.Windows.Forms.Label();
            this.Label378 = new System.Windows.Forms.Label();
            this.lblLine0 = new System.Windows.Forms.Label();
            this.pnlRefresh = new System.Windows.Forms.Panel();
            this.Label251 = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnEnd = new System.Windows.Forms.Button();
            this.Button21 = new System.Windows.Forms.Button();
            this.Button20 = new System.Windows.Forms.Button();
            this.Button19 = new System.Windows.Forms.Button();
            this.lblTitleH = new System.Windows.Forms.Label();
            this.pgbCheck = new System.Windows.Forms.ProgressBar();
            this.lblPgbCheck = new System.Windows.Forms.Label();
            this.lblCheckSituation = new System.Windows.Forms.Label();
            this.pnlPrgChk = new System.Windows.Forms.Panel();
            this.pnlMenuDatacv.SuspendLayout();
            this.pnlPrgDCConv.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCConv2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCConv1)).BeginInit();
            this.pnlPrgDCRelation.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCRelation2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCRelation1)).BeginInit();
            this.pnlPrgDCFileWrite.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCFileWrite2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCFileWrite1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCEnd2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCJikko2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCSelect2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCEnd1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCJikko1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCSelect1)).BeginInit();
            this.pnlPrgDCEnd.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCEnd2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCEnd1)).BeginInit();
            this.pnlPrgDCJikko.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCJikko2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCJikko1)).BeginInit();
            this.pnlPrgDCSelect.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCSelect2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCSelect1)).BeginInit();
            this.pnlPrgDCHajimeni.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCHajimeni2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCHajimeni1)).BeginInit();
            this.tabCtrlMain.SuspendLayout();
            this.tabPageDev.SuspendLayout();
            this.GroupBox6.SuspendLayout();
            this.grpDevSettingX.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox1)).BeginInit();
            this.grpDevSetting1.SuspendLayout();
            this.grpDevSetting2.SuspendLayout();
            this.grpDevSetting3.SuspendLayout();
            this.GroupBox3.SuspendLayout();
            this.GroupBox2.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            this.grptaihi.SuspendLayout();
            this.tabPageStart.SuspendLayout();
            this.grpHFirstNaiyo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox2)).BeginInit();
            this.tabPageSession.SuspendLayout();
            this.grpTimeOut.SuspendLayout();
            this.grp10ConnectInfo.SuspendLayout();
            this.grpV10Authent.SuspendLayout();
            this.tabPageSyoki.SuspendLayout();
            this.pnlUserName.SuspendLayout();
            this.pnlLogPath.SuspendLayout();
            this.tabPageMenu.SuspendLayout();
            this.grpMenuJizen.SuspendLayout();
            this.grpMenuDatacv.SuspendLayout();
            this.grpMenuJigo.SuspendLayout();
            this.tabPageJizen.SuspendLayout();
            this.pnlJizenListPath.SuspendLayout();
            this.tabCtrlJizen.SuspendLayout();
            this.tabPageHJizen1.SuspendLayout();
            this.GroupBox5.SuspendLayout();
            this.pnlHJizenTyukan.SuspendLayout();
            this.tabPageJigo.SuspendLayout();
            this.tabCtrlJigo.SuspendLayout();
            this.tabPageJigo1.SuspendLayout();
            this.grpJigoDonyuji.SuspendLayout();
            this.pnlJigoCmtSoKotiku.SuspendLayout();
            this.pnlJigoCmtNkNyuryoku.SuspendLayout();
            this.pnlJigoCmtSqKotiku.SuspendLayout();
            this.grpJigoUserSagyo.SuspendLayout();
            this.pnlHJigoCmtSyudo.SuspendLayout();
            this.tabPageHajimeni.SuspendLayout();
            this.GroupBox10.SuspendLayout();
            this.pnlDcFstCmtH99.SuspendLayout();
            this.pnlDcFstCmtH01.SuspendLayout();
            this.pnlDcFstCmtH02.SuspendLayout();
            this.pnlDcFstCmt11.SuspendLayout();
            this.pnlDcFstCmt06.SuspendLayout();
            this.pnlDcFstCmt05.SuspendLayout();
            this.pnlDcFstCmt07.SuspendLayout();
            this.pnlDcFstCmt10.SuspendLayout();
            this.pnlDcFstCmt02.SuspendLayout();
            this.pnlDcFstCmt08.SuspendLayout();
            this.pnlDcFstCmt04.SuspendLayout();
            this.pnlDcFstCmt03.SuspendLayout();
            this.pnlDcFstCmt01.SuspendLayout();
            this.pnlDcFstCmt09.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox14)).BeginInit();
            this.tabPageSelect.SuspendLayout();
            this.grpMiddleFile.SuspendLayout();
            this.tabCtrlCVItem.SuspendLayout();
            this.tabPageKizon110.SuspendLayout();
            this.grpKizon1.SuspendLayout();
            this.pnlKiMstHendo.SuspendLayout();
            this.pnlKiMstTokuyaku.SuspendLayout();
            this.pnlKiMstKasyoClaimrui.SuspendLayout();
            this.pnlKiMstTitle.SuspendLayout();
            this.pnlKiMstArea.SuspendLayout();
            this.pnlKiMstSchool.SuspendLayout();
            this.pnlKiMstHokenrui.SuspendLayout();
            this.pnlKiMstBus.SuspendLayout();
            this.tabPageKizon120.SuspendLayout();
            this.grpKizon3.SuspendLayout();
            this.pnlKiGySyuzenBase.SuspendLayout();
            this.pnlKiGyYatinhosyoBase.SuspendLayout();
            this.pnlKiGyLifelineBase.SuspendLayout();
            this.pnlKiGyHokenBase.SuspendLayout();
            this.pnlKiGySisetuBase.SuspendLayout();
            this.pnlKiGyCyukaiBase.SuspendLayout();
            this.pnlKiGySekoBase.SuspendLayout();
            this.grpKizon5.SuspendLayout();
            this.pnlKiOw.SuspendLayout();
            this.pnlKiJisya.SuspendLayout();
            this.pnlKiSyskanriBase.SuspendLayout();
            this.tabPageKizon130.SuspendLayout();
            this.grpKizon2.SuspendLayout();
            this.pnlRendo.SuspendLayout();
            this.pnlKiSq.SuspendLayout();
            this.pnlSzen.SuspendLayout();
            this.pnlKiClaim.SuspendLayout();
            this.pnlKiKy.SuspendLayout();
            this.pnlKiKys.SuspendLayout();
            this.pnlKiHy.SuspendLayout();
            this.pnlKiBk.SuspendLayout();
            this.grpKizonKagi.SuspendLayout();
            this.tabPageBase110.SuspendLayout();
            this.grpMst.SuspendLayout();
            this.tabPageBase120.SuspendLayout();
            this.grpGy.SuspendLayout();
            this.tabPageBase130.SuspendLayout();
            this.grpKys.SuspendLayout();
            this.grpOw.SuspendLayout();
            this.grpJisya.SuspendLayout();
            this.tabPageBase140.SuspendLayout();
            this.grpBk.SuspendLayout();
            this.grpKagiSelect.SuspendLayout();
            this.tabPageBase150.SuspendLayout();
            this.grpHy.SuspendLayout();
            this.tabPageBase160.SuspendLayout();
            this.grpSorule.SuspendLayout();
            this.tabPageBase170.SuspendLayout();
            this.grpKy.SuspendLayout();
            this.tabPageBase180.SuspendLayout();
            this.grpSq.SuspendLayout();
            this.tabPageBase190.SuspendLayout();
            this.grpSzen.SuspendLayout();
            this.grpClaim.SuspendLayout();
            this.tabPageBase200.SuspendLayout();
            this.grpSyskanri.SuspendLayout();
            this.tabPageBase210.SuspendLayout();
            this.grpRendo.SuspendLayout();
            this.tabPageBase900.SuspendLayout();
            this.tabPageHanyo110.SuspendLayout();
            this.grpHMst.SuspendLayout();
            this.tabPageHanyo120.SuspendLayout();
            this.grpHGy.SuspendLayout();
            this.grpHOw.SuspendLayout();
            this.grpHJisya.SuspendLayout();
            this.TabPage1.SuspendLayout();
            this.grpKizon4.SuspendLayout();
            this.pnlRekiClear.SuspendLayout();
            this.tabPageJikko.SuspendLayout();
            this.grpTotalProcess.SuspendLayout();
            this.tabPageEndOK.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox15)).BeginInit();
            this.tabPageEndError.SuspendLayout();
            this.GroupBox23.SuspendLayout();
            this.GroupBox19.SuspendLayout();
            this.GroupBox20.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox23)).BeginInit();
            this.tabPageEndCancel.SuspendLayout();
            this.GroupBox21.SuspendLayout();
            this.GroupBox22.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox25)).BeginInit();
            this.pnlRefresh.SuspendLayout();
            this.pnlPrgChk.SuspendLayout();
            this.SuspendLayout();
            // 
            // chkRelationFile
            // 
            this.chkRelationFile.AutoSize = true;
            this.chkRelationFile.Font = new System.Drawing.Font("メイリオ", 6.75F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkRelationFile.Location = new System.Drawing.Point(8, 126);
            this.chkRelationFile.Name = "chkRelationFile";
            this.chkRelationFile.Size = new System.Drawing.Size(125, 18);
            this.chkRelationFile.TabIndex = 110;
            this.chkRelationFile.Text = "紐付けファイル作成済み";
            this.chkRelationFile.UseVisualStyleBackColor = true;
            // 
            // Label9
            // 
            this.Label9.Font = new System.Drawing.Font("メイリオ", 6.75F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label9.Location = new System.Drawing.Point(8, 147);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(211, 31);
            this.Label9.TabIndex = 116;
            this.Label9.Text = "紐付け設定を仮TBLへ登録済みの場合にチェックを付けて下さい。";
            this.Label9.UseCompatibleTextRendering = true;
            // 
            // pnlMenuDatacv
            // 
            this.pnlMenuDatacv.BackColor = System.Drawing.Color.Gainsboro;
            this.pnlMenuDatacv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnlMenuDatacv.Controls.Add(this.pnlPrgDCConv);
            this.pnlMenuDatacv.Controls.Add(this.pnlPrgDCRelation);
            this.pnlMenuDatacv.Controls.Add(this.pnlPrgDCFileWrite);
            this.pnlMenuDatacv.Controls.Add(this.picArwDCEnd2);
            this.pnlMenuDatacv.Controls.Add(this.picArwDCJikko2);
            this.pnlMenuDatacv.Controls.Add(this.picArwDCSelect2);
            this.pnlMenuDatacv.Controls.Add(this.picArwDCEnd1);
            this.pnlMenuDatacv.Controls.Add(this.picArwDCJikko1);
            this.pnlMenuDatacv.Controls.Add(this.picArwDCSelect1);
            this.pnlMenuDatacv.Controls.Add(this.pnlPrgDCEnd);
            this.pnlMenuDatacv.Controls.Add(this.pnlPrgDCJikko);
            this.pnlMenuDatacv.Controls.Add(this.pnlPrgDCSelect);
            this.pnlMenuDatacv.Controls.Add(this.pnlPrgDCHajimeni);
            this.pnlMenuDatacv.Location = new System.Drawing.Point(33, 60);
            this.pnlMenuDatacv.Name = "pnlMenuDatacv";
            this.pnlMenuDatacv.Size = new System.Drawing.Size(250, 561);
            this.pnlMenuDatacv.TabIndex = 1;
            // 
            // pnlPrgDCConv
            // 
            this.pnlPrgDCConv.BackColor = System.Drawing.Color.Gainsboro;
            this.pnlPrgDCConv.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlPrgDCConv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgDCConv.CausesValidation = false;
            this.pnlPrgDCConv.Controls.Add(this.picIcoDCConv2);
            this.pnlPrgDCConv.Controls.Add(this.picIcoDCConv1);
            this.pnlPrgDCConv.Controls.Add(this.lblPrgDCConv);
            this.pnlPrgDCConv.Location = new System.Drawing.Point(31, 336);
            this.pnlPrgDCConv.Name = "pnlPrgDCConv";
            this.pnlPrgDCConv.Size = new System.Drawing.Size(184, 33);
            this.pnlPrgDCConv.TabIndex = 9;
            // 
            // picIcoDCConv2
            // 
            this.picIcoDCConv2.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCConv2.Image")));
            this.picIcoDCConv2.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCConv2.Name = "picIcoDCConv2";
            this.picIcoDCConv2.Size = new System.Drawing.Size(25, 25);
            this.picIcoDCConv2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCConv2.TabIndex = 4;
            this.picIcoDCConv2.TabStop = false;
            // 
            // picIcoDCConv1
            // 
            this.picIcoDCConv1.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCConv1.Image")));
            this.picIcoDCConv1.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCConv1.Name = "picIcoDCConv1";
            this.picIcoDCConv1.Size = new System.Drawing.Size(25, 26);
            this.picIcoDCConv1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCConv1.TabIndex = 3;
            this.picIcoDCConv1.TabStop = false;
            // 
            // lblPrgDCConv
            // 
            this.lblPrgDCConv.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPrgDCConv.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblPrgDCConv.Location = new System.Drawing.Point(33, 5);
            this.lblPrgDCConv.Name = "lblPrgDCConv";
            this.lblPrgDCConv.Size = new System.Drawing.Size(149, 20);
            this.lblPrgDCConv.TabIndex = 0;
            this.lblPrgDCConv.Text = " データコンバート";
            // 
            // pnlPrgDCRelation
            // 
            this.pnlPrgDCRelation.BackColor = System.Drawing.Color.Gainsboro;
            this.pnlPrgDCRelation.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlPrgDCRelation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgDCRelation.CausesValidation = false;
            this.pnlPrgDCRelation.Controls.Add(this.picIcoDCRelation2);
            this.pnlPrgDCRelation.Controls.Add(this.picIcoDCRelation1);
            this.pnlPrgDCRelation.Controls.Add(this.lblPrgDCRelation);
            this.pnlPrgDCRelation.Location = new System.Drawing.Point(31, 297);
            this.pnlPrgDCRelation.Name = "pnlPrgDCRelation";
            this.pnlPrgDCRelation.Size = new System.Drawing.Size(184, 33);
            this.pnlPrgDCRelation.TabIndex = 8;
            // 
            // picIcoDCRelation2
            // 
            this.picIcoDCRelation2.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCRelation2.Image")));
            this.picIcoDCRelation2.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCRelation2.Name = "picIcoDCRelation2";
            this.picIcoDCRelation2.Size = new System.Drawing.Size(25, 25);
            this.picIcoDCRelation2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCRelation2.TabIndex = 4;
            this.picIcoDCRelation2.TabStop = false;
            // 
            // picIcoDCRelation1
            // 
            this.picIcoDCRelation1.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCRelation1.Image")));
            this.picIcoDCRelation1.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCRelation1.Name = "picIcoDCRelation1";
            this.picIcoDCRelation1.Size = new System.Drawing.Size(25, 26);
            this.picIcoDCRelation1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCRelation1.TabIndex = 3;
            this.picIcoDCRelation1.TabStop = false;
            // 
            // lblPrgDCRelation
            // 
            this.lblPrgDCRelation.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPrgDCRelation.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblPrgDCRelation.Location = new System.Drawing.Point(33, 5);
            this.lblPrgDCRelation.Name = "lblPrgDCRelation";
            this.lblPrgDCRelation.Size = new System.Drawing.Size(149, 20);
            this.lblPrgDCRelation.TabIndex = 0;
            this.lblPrgDCRelation.Text = " 紐 付 設 定 作 業";
            // 
            // pnlPrgDCFileWrite
            // 
            this.pnlPrgDCFileWrite.BackColor = System.Drawing.Color.Gainsboro;
            this.pnlPrgDCFileWrite.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlPrgDCFileWrite.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgDCFileWrite.CausesValidation = false;
            this.pnlPrgDCFileWrite.Controls.Add(this.picIcoDCFileWrite2);
            this.pnlPrgDCFileWrite.Controls.Add(this.picIcoDCFileWrite1);
            this.pnlPrgDCFileWrite.Controls.Add(this.lblPrgDCFileWrite);
            this.pnlPrgDCFileWrite.Location = new System.Drawing.Point(31, 258);
            this.pnlPrgDCFileWrite.Name = "pnlPrgDCFileWrite";
            this.pnlPrgDCFileWrite.Size = new System.Drawing.Size(184, 33);
            this.pnlPrgDCFileWrite.TabIndex = 7;
            // 
            // picIcoDCFileWrite2
            // 
            this.picIcoDCFileWrite2.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCFileWrite2.Image")));
            this.picIcoDCFileWrite2.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCFileWrite2.Name = "picIcoDCFileWrite2";
            this.picIcoDCFileWrite2.Size = new System.Drawing.Size(25, 25);
            this.picIcoDCFileWrite2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCFileWrite2.TabIndex = 4;
            this.picIcoDCFileWrite2.TabStop = false;
            // 
            // picIcoDCFileWrite1
            // 
            this.picIcoDCFileWrite1.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCFileWrite1.Image")));
            this.picIcoDCFileWrite1.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCFileWrite1.Name = "picIcoDCFileWrite1";
            this.picIcoDCFileWrite1.Size = new System.Drawing.Size(25, 26);
            this.picIcoDCFileWrite1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCFileWrite1.TabIndex = 3;
            this.picIcoDCFileWrite1.TabStop = false;
            // 
            // lblPrgDCFileWrite
            // 
            this.lblPrgDCFileWrite.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPrgDCFileWrite.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblPrgDCFileWrite.Location = new System.Drawing.Point(33, 5);
            this.lblPrgDCFileWrite.Name = "lblPrgDCFileWrite";
            this.lblPrgDCFileWrite.Size = new System.Drawing.Size(149, 20);
            this.lblPrgDCFileWrite.TabIndex = 0;
            this.lblPrgDCFileWrite.Text = " 中間ファイル書込";
            // 
            // picArwDCEnd2
            // 
            this.picArwDCEnd2.Image = ((System.Drawing.Image)(resources.GetObject("picArwDCEnd2.Image")));
            this.picArwDCEnd2.Location = new System.Drawing.Point(103, 375);
            this.picArwDCEnd2.Name = "picArwDCEnd2";
            this.picArwDCEnd2.Size = new System.Drawing.Size(40, 35);
            this.picArwDCEnd2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picArwDCEnd2.TabIndex = 12;
            this.picArwDCEnd2.TabStop = false;
            // 
            // picArwDCJikko2
            // 
            this.picArwDCJikko2.Image = ((System.Drawing.Image)(resources.GetObject("picArwDCJikko2.Image")));
            this.picArwDCJikko2.Location = new System.Drawing.Point(103, 157);
            this.picArwDCJikko2.Name = "picArwDCJikko2";
            this.picArwDCJikko2.Size = new System.Drawing.Size(40, 35);
            this.picArwDCJikko2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picArwDCJikko2.TabIndex = 11;
            this.picArwDCJikko2.TabStop = false;
            // 
            // picArwDCSelect2
            // 
            this.picArwDCSelect2.Image = ((System.Drawing.Image)(resources.GetObject("picArwDCSelect2.Image")));
            this.picArwDCSelect2.Location = new System.Drawing.Point(103, 62);
            this.picArwDCSelect2.Name = "picArwDCSelect2";
            this.picArwDCSelect2.Size = new System.Drawing.Size(40, 35);
            this.picArwDCSelect2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picArwDCSelect2.TabIndex = 10;
            this.picArwDCSelect2.TabStop = false;
            // 
            // picArwDCEnd1
            // 
            this.picArwDCEnd1.Image = ((System.Drawing.Image)(resources.GetObject("picArwDCEnd1.Image")));
            this.picArwDCEnd1.Location = new System.Drawing.Point(103, 375);
            this.picArwDCEnd1.Name = "picArwDCEnd1";
            this.picArwDCEnd1.Size = new System.Drawing.Size(40, 35);
            this.picArwDCEnd1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picArwDCEnd1.TabIndex = 9;
            this.picArwDCEnd1.TabStop = false;
            // 
            // picArwDCJikko1
            // 
            this.picArwDCJikko1.Image = ((System.Drawing.Image)(resources.GetObject("picArwDCJikko1.Image")));
            this.picArwDCJikko1.Location = new System.Drawing.Point(103, 157);
            this.picArwDCJikko1.Name = "picArwDCJikko1";
            this.picArwDCJikko1.Size = new System.Drawing.Size(40, 35);
            this.picArwDCJikko1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picArwDCJikko1.TabIndex = 8;
            this.picArwDCJikko1.TabStop = false;
            // 
            // picArwDCSelect1
            // 
            this.picArwDCSelect1.Image = ((System.Drawing.Image)(resources.GetObject("picArwDCSelect1.Image")));
            this.picArwDCSelect1.Location = new System.Drawing.Point(103, 62);
            this.picArwDCSelect1.Name = "picArwDCSelect1";
            this.picArwDCSelect1.Size = new System.Drawing.Size(40, 35);
            this.picArwDCSelect1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picArwDCSelect1.TabIndex = 7;
            this.picArwDCSelect1.TabStop = false;
            // 
            // pnlPrgDCEnd
            // 
            this.pnlPrgDCEnd.BackColor = System.Drawing.Color.Gainsboro;
            this.pnlPrgDCEnd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlPrgDCEnd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgDCEnd.CausesValidation = false;
            this.pnlPrgDCEnd.Controls.Add(this.picIcoDCEnd2);
            this.pnlPrgDCEnd.Controls.Add(this.picIcoDCEnd1);
            this.pnlPrgDCEnd.Controls.Add(this.lblPrgDCEnd);
            this.pnlPrgDCEnd.Location = new System.Drawing.Point(12, 416);
            this.pnlPrgDCEnd.Name = "pnlPrgDCEnd";
            this.pnlPrgDCEnd.Size = new System.Drawing.Size(221, 47);
            this.pnlPrgDCEnd.TabIndex = 6;
            // 
            // picIcoDCEnd2
            // 
            this.picIcoDCEnd2.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCEnd2.Image")));
            this.picIcoDCEnd2.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCEnd2.Name = "picIcoDCEnd2";
            this.picIcoDCEnd2.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCEnd2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCEnd2.TabIndex = 4;
            this.picIcoDCEnd2.TabStop = false;
            // 
            // picIcoDCEnd1
            // 
            this.picIcoDCEnd1.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCEnd1.Image")));
            this.picIcoDCEnd1.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCEnd1.Name = "picIcoDCEnd1";
            this.picIcoDCEnd1.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCEnd1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picIcoDCEnd1.TabIndex = 3;
            this.picIcoDCEnd1.TabStop = false;
            // 
            // lblPrgDCEnd
            // 
            this.lblPrgDCEnd.Font = new System.Drawing.Font("メイリオ", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPrgDCEnd.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblPrgDCEnd.Location = new System.Drawing.Point(60, 8);
            this.lblPrgDCEnd.Name = "lblPrgDCEnd";
            this.lblPrgDCEnd.Size = new System.Drawing.Size(140, 30);
            this.lblPrgDCEnd.TabIndex = 0;
            this.lblPrgDCEnd.Text = "  終      了  ";
            this.lblPrgDCEnd.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlPrgDCJikko
            // 
            this.pnlPrgDCJikko.BackColor = System.Drawing.Color.Gainsboro;
            this.pnlPrgDCJikko.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlPrgDCJikko.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgDCJikko.CausesValidation = false;
            this.pnlPrgDCJikko.Controls.Add(this.picIcoDCJikko2);
            this.pnlPrgDCJikko.Controls.Add(this.picIcoDCJikko1);
            this.pnlPrgDCJikko.Controls.Add(this.lblPrgDCJikko);
            this.pnlPrgDCJikko.Location = new System.Drawing.Point(12, 200);
            this.pnlPrgDCJikko.Name = "pnlPrgDCJikko";
            this.pnlPrgDCJikko.Size = new System.Drawing.Size(221, 47);
            this.pnlPrgDCJikko.TabIndex = 5;
            // 
            // picIcoDCJikko2
            // 
            this.picIcoDCJikko2.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCJikko2.Image")));
            this.picIcoDCJikko2.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCJikko2.Name = "picIcoDCJikko2";
            this.picIcoDCJikko2.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCJikko2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picIcoDCJikko2.TabIndex = 4;
            this.picIcoDCJikko2.TabStop = false;
            // 
            // picIcoDCJikko1
            // 
            this.picIcoDCJikko1.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCJikko1.Image")));
            this.picIcoDCJikko1.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCJikko1.Name = "picIcoDCJikko1";
            this.picIcoDCJikko1.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCJikko1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picIcoDCJikko1.TabIndex = 3;
            this.picIcoDCJikko1.TabStop = false;
            // 
            // lblPrgDCJikko
            // 
            this.lblPrgDCJikko.Font = new System.Drawing.Font("メイリオ", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPrgDCJikko.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblPrgDCJikko.Location = new System.Drawing.Point(60, 8);
            this.lblPrgDCJikko.Name = "lblPrgDCJikko";
            this.lblPrgDCJikko.Size = new System.Drawing.Size(140, 30);
            this.lblPrgDCJikko.TabIndex = 0;
            this.lblPrgDCJikko.Text = " 移 行 処 理";
            this.lblPrgDCJikko.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlPrgDCSelect
            // 
            this.pnlPrgDCSelect.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlPrgDCSelect.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgDCSelect.CausesValidation = false;
            this.pnlPrgDCSelect.Controls.Add(this.picIcoDCSelect2);
            this.pnlPrgDCSelect.Controls.Add(this.picIcoDCSelect1);
            this.pnlPrgDCSelect.Controls.Add(this.lblPrgDCSelect);
            this.pnlPrgDCSelect.Location = new System.Drawing.Point(12, 103);
            this.pnlPrgDCSelect.Name = "pnlPrgDCSelect";
            this.pnlPrgDCSelect.Size = new System.Drawing.Size(221, 47);
            this.pnlPrgDCSelect.TabIndex = 4;
            // 
            // picIcoDCSelect2
            // 
            this.picIcoDCSelect2.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCSelect2.Image")));
            this.picIcoDCSelect2.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCSelect2.Name = "picIcoDCSelect2";
            this.picIcoDCSelect2.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCSelect2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picIcoDCSelect2.TabIndex = 4;
            this.picIcoDCSelect2.TabStop = false;
            // 
            // picIcoDCSelect1
            // 
            this.picIcoDCSelect1.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCSelect1.Image")));
            this.picIcoDCSelect1.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCSelect1.Name = "picIcoDCSelect1";
            this.picIcoDCSelect1.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCSelect1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picIcoDCSelect1.TabIndex = 3;
            this.picIcoDCSelect1.TabStop = false;
            // 
            // lblPrgDCSelect
            // 
            this.lblPrgDCSelect.Font = new System.Drawing.Font("メイリオ", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPrgDCSelect.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblPrgDCSelect.Location = new System.Drawing.Point(60, 8);
            this.lblPrgDCSelect.Name = "lblPrgDCSelect";
            this.lblPrgDCSelect.Size = new System.Drawing.Size(140, 30);
            this.lblPrgDCSelect.TabIndex = 0;
            this.lblPrgDCSelect.Text = "対象項目選択";
            this.lblPrgDCSelect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlPrgDCHajimeni
            // 
            this.pnlPrgDCHajimeni.BackColor = System.Drawing.Color.Gainsboro;
            this.pnlPrgDCHajimeni.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlPrgDCHajimeni.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgDCHajimeni.CausesValidation = false;
            this.pnlPrgDCHajimeni.Controls.Add(this.picIcoDCHajimeni2);
            this.pnlPrgDCHajimeni.Controls.Add(this.picIcoDCHajimeni1);
            this.pnlPrgDCHajimeni.Controls.Add(this.lblPrgDCHajimeni);
            this.pnlPrgDCHajimeni.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.pnlPrgDCHajimeni.Location = new System.Drawing.Point(12, 8);
            this.pnlPrgDCHajimeni.Name = "pnlPrgDCHajimeni";
            this.pnlPrgDCHajimeni.Size = new System.Drawing.Size(221, 47);
            this.pnlPrgDCHajimeni.TabIndex = 0;
            // 
            // picIcoDCHajimeni2
            // 
            this.picIcoDCHajimeni2.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCHajimeni2.Image")));
            this.picIcoDCHajimeni2.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCHajimeni2.Name = "picIcoDCHajimeni2";
            this.picIcoDCHajimeni2.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCHajimeni2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picIcoDCHajimeni2.TabIndex = 3;
            this.picIcoDCHajimeni2.TabStop = false;
            // 
            // picIcoDCHajimeni1
            // 
            this.picIcoDCHajimeni1.Image = ((System.Drawing.Image)(resources.GetObject("picIcoDCHajimeni1.Image")));
            this.picIcoDCHajimeni1.Location = new System.Drawing.Point(3, 3);
            this.picIcoDCHajimeni1.Name = "picIcoDCHajimeni1";
            this.picIcoDCHajimeni1.Size = new System.Drawing.Size(40, 40);
            this.picIcoDCHajimeni1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picIcoDCHajimeni1.TabIndex = 2;
            this.picIcoDCHajimeni1.TabStop = false;
            // 
            // lblPrgDCHajimeni
            // 
            this.lblPrgDCHajimeni.Font = new System.Drawing.Font("メイリオ", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPrgDCHajimeni.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblPrgDCHajimeni.Location = new System.Drawing.Point(60, 8);
            this.lblPrgDCHajimeni.Name = "lblPrgDCHajimeni";
            this.lblPrgDCHajimeni.Size = new System.Drawing.Size(140, 30);
            this.lblPrgDCHajimeni.TabIndex = 0;
            this.lblPrgDCHajimeni.Text = " は じ め に";
            this.lblPrgDCHajimeni.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblHidden1
            // 
            this.lblHidden1.Location = new System.Drawing.Point(289, 11);
            this.lblHidden1.Name = "lblHidden1";
            this.lblHidden1.Size = new System.Drawing.Size(21, 55);
            this.lblHidden1.TabIndex = 2;
            this.lblHidden1.Text = "　";
            // 
            // tabCtrlMain
            // 
            this.tabCtrlMain.Controls.Add(this.tabPageDev);
            this.tabCtrlMain.Controls.Add(this.tabPageStart);
            this.tabCtrlMain.Controls.Add(this.tabPageSession);
            this.tabCtrlMain.Controls.Add(this.tabPageSyoki);
            this.tabCtrlMain.Controls.Add(this.tabPageMenu);
            this.tabCtrlMain.Controls.Add(this.tabPageJizen);
            this.tabCtrlMain.Controls.Add(this.tabPageJigo);
            this.tabCtrlMain.Controls.Add(this.tabPageHojyo);
            this.tabCtrlMain.Controls.Add(this.tabPageHajimeni);
            this.tabCtrlMain.Controls.Add(this.tabPageSelect);
            this.tabCtrlMain.Controls.Add(this.tabPageJikko);
            this.tabCtrlMain.Controls.Add(this.tabPageEndOK);
            this.tabCtrlMain.Controls.Add(this.tabPageEndError);
            this.tabCtrlMain.Controls.Add(this.tabPageEndCancel);
            this.tabCtrlMain.Controls.Add(this.tabPageIkkatu);
            this.tabCtrlMain.Controls.Add(this.tabPageHanyoJizen);
            this.tabCtrlMain.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabCtrlMain.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.tabCtrlMain.Location = new System.Drawing.Point(316, 40);
            this.tabCtrlMain.Name = "tabCtrlMain";
            this.tabCtrlMain.SelectedIndex = 0;
            this.tabCtrlMain.Size = new System.Drawing.Size(920, 580);
            this.tabCtrlMain.TabIndex = 3;
            // 
            // tabPageDev
            // 
            this.tabPageDev.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageDev.Controls.Add(this.GroupBox6);
            this.tabPageDev.Controls.Add(this.grpDevSettingX);
            this.tabPageDev.Controls.Add(this.PictureBox1);
            this.tabPageDev.Controls.Add(this.grpDevSetting1);
            this.tabPageDev.Controls.Add(this.grpDevSetting2);
            this.tabPageDev.Controls.Add(this.grpDevSetting3);
            this.tabPageDev.Controls.Add(this.grpRelation);
            this.tabPageDev.Controls.Add(this.GroupBox3);
            this.tabPageDev.Controls.Add(this.GroupBox2);
            this.tabPageDev.Controls.Add(this.GroupBox1);
            this.tabPageDev.Controls.Add(this.grptaihi);
            this.tabPageDev.Location = new System.Drawing.Point(4, 27);
            this.tabPageDev.Name = "tabPageDev";
            this.tabPageDev.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageDev.Size = new System.Drawing.Size(912, 549);
            this.tabPageDev.TabIndex = 5;
            this.tabPageDev.Text = " 開発用";
            // 
            // GroupBox6
            // 
            this.GroupBox6.Controls.Add(this.Label361);
            this.GroupBox6.Controls.Add(this.txtExistMidToBaseMid);
            this.GroupBox6.Controls.Add(this.btnExistMidToBaseMid);
            this.GroupBox6.Controls.Add(this.btnExistMidToBaseMidDirSerach);
            this.GroupBox6.Location = new System.Drawing.Point(670, 226);
            this.GroupBox6.Name = "GroupBox6";
            this.GroupBox6.Size = new System.Drawing.Size(228, 108);
            this.GroupBox6.TabIndex = 168;
            this.GroupBox6.TabStop = false;
            this.GroupBox6.Text = "【テスト用汎用中間ファイル作成】";
            // 
            // Label361
            // 
            this.Label361.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label361.Location = new System.Drawing.Point(18, 21);
            this.Label361.Name = "Label361";
            this.Label361.Size = new System.Drawing.Size(186, 20);
            this.Label361.TabIndex = 145;
            this.Label361.Text = "中間ファイル作成元格納先";
            this.Label361.UseCompatibleTextRendering = true;
            // 
            // txtExistMidToBaseMid
            // 
            this.txtExistMidToBaseMid.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtExistMidToBaseMid.Location = new System.Drawing.Point(18, 42);
            this.txtExistMidToBaseMid.Name = "txtExistMidToBaseMid";
            this.txtExistMidToBaseMid.Size = new System.Drawing.Size(164, 24);
            this.txtExistMidToBaseMid.TabIndex = 143;
            // 
            // btnExistMidToBaseMid
            // 
            this.btnExistMidToBaseMid.BackColor = System.Drawing.SystemColors.Menu;
            this.btnExistMidToBaseMid.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExistMidToBaseMid.Image = ((System.Drawing.Image)(resources.GetObject("btnExistMidToBaseMid.Image")));
            this.btnExistMidToBaseMid.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExistMidToBaseMid.Location = new System.Drawing.Point(59, 72);
            this.btnExistMidToBaseMid.Name = "btnExistMidToBaseMid";
            this.btnExistMidToBaseMid.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnExistMidToBaseMid.Size = new System.Drawing.Size(100, 25);
            this.btnExistMidToBaseMid.TabIndex = 46;
            this.btnExistMidToBaseMid.Text = "既存→汎用";
            this.btnExistMidToBaseMid.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnExistMidToBaseMid.UseVisualStyleBackColor = true;
            this.btnExistMidToBaseMid.Click += new System.EventHandler(this.btnExistMidToBaseMid_Click);
            // 
            // btnExistMidToBaseMidDirSerach
            // 
            this.btnExistMidToBaseMidDirSerach.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExistMidToBaseMidDirSerach.Location = new System.Drawing.Point(192, 42);
            this.btnExistMidToBaseMidDirSerach.Name = "btnExistMidToBaseMidDirSerach";
            this.btnExistMidToBaseMidDirSerach.Size = new System.Drawing.Size(30, 25);
            this.btnExistMidToBaseMidDirSerach.TabIndex = 144;
            this.btnExistMidToBaseMidDirSerach.Text = "...";
            this.btnExistMidToBaseMidDirSerach.UseVisualStyleBackColor = true;
            this.btnExistMidToBaseMidDirSerach.Click += new System.EventHandler(this.btnDirSeach_Click);
            // 
            // grpDevSettingX
            // 
            this.grpDevSettingX.Controls.Add(this.chkRelationFile);
            this.grpDevSettingX.Controls.Add(this.Label9);
            this.grpDevSettingX.Controls.Add(this.chkChildItemControl);
            this.grpDevSettingX.Controls.Add(this.Label48);
            this.grpDevSettingX.Controls.Add(this.Label41);
            this.grpDevSettingX.Controls.Add(this.Label50);
            this.grpDevSettingX.Controls.Add(this.Label49);
            this.grpDevSettingX.Controls.Add(this.txtLogOutputCnt);
            this.grpDevSettingX.Location = new System.Drawing.Point(670, 354);
            this.grpDevSettingX.Name = "grpDevSettingX";
            this.grpDevSettingX.Size = new System.Drawing.Size(228, 188);
            this.grpDevSettingX.TabIndex = 167;
            this.grpDevSettingX.TabStop = false;
            this.grpDevSettingX.Text = "【没】";
            // 
            // chkChildItemControl
            // 
            this.chkChildItemControl.AutoSize = true;
            this.chkChildItemControl.Font = new System.Drawing.Font("メイリオ", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkChildItemControl.Location = new System.Drawing.Point(7, 19);
            this.chkChildItemControl.Name = "chkChildItemControl";
            this.chkChildItemControl.Size = new System.Drawing.Size(152, 18);
            this.chkChildItemControl.TabIndex = 161;
            this.chkChildItemControl.Text = "子項目のチェックボックス制御";
            this.chkChildItemControl.UseVisualStyleBackColor = true;
            // 
            // Label48
            // 
            this.Label48.Font = new System.Drawing.Font("メイリオ", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label48.Location = new System.Drawing.Point(8, 39);
            this.Label48.Name = "Label48";
            this.Label48.Size = new System.Drawing.Size(211, 30);
            this.Label48.TabIndex = 162;
            this.Label48.Text = "対象項目選択画面内のチェックボックスの親子関係によるチェックを制御します。";
            this.Label48.UseCompatibleTextRendering = true;
            // 
            // Label41
            // 
            this.Label41.Font = new System.Drawing.Font("メイリオ", 7F);
            this.Label41.Location = new System.Drawing.Point(7, 71);
            this.Label41.Name = "Label41";
            this.Label41.Size = new System.Drawing.Size(73, 16);
            this.Label41.TabIndex = 147;
            this.Label41.Text = "● ログ出力件数";
            this.Label41.UseCompatibleTextRendering = true;
            // 
            // Label50
            // 
            this.Label50.Font = new System.Drawing.Font("メイリオ", 7F);
            this.Label50.Location = new System.Drawing.Point(188, 86);
            this.Label50.Name = "Label50";
            this.Label50.Size = new System.Drawing.Size(32, 16);
            this.Label50.TabIndex = 149;
            this.Label50.Text = "件毎";
            this.Label50.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label50.UseCompatibleTextRendering = true;
            // 
            // Label49
            // 
            this.Label49.Font = new System.Drawing.Font("メイリオ", 7F);
            this.Label49.Location = new System.Drawing.Point(8, 87);
            this.Label49.Name = "Label49";
            this.Label49.Size = new System.Drawing.Size(115, 29);
            this.Label49.TabIndex = 148;
            this.Label49.Text = "右記に設定された件数毎にログを出力します。";
            this.Label49.UseCompatibleTextRendering = true;
            // 
            // txtLogOutputCnt
            // 
            this.txtLogOutputCnt.Font = new System.Drawing.Font("メイリオ", 7F);
            this.txtLogOutputCnt.Location = new System.Drawing.Point(129, 84);
            this.txtLogOutputCnt.Name = "txtLogOutputCnt";
            this.txtLogOutputCnt.Size = new System.Drawing.Size(53, 21);
            this.txtLogOutputCnt.TabIndex = 146;
            this.txtLogOutputCnt.Text = "1";
            this.txtLogOutputCnt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // PictureBox1
            // 
            this.PictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("PictureBox1.Image")));
            this.PictureBox1.Location = new System.Drawing.Point(770, 27);
            this.PictureBox1.Name = "PictureBox1";
            this.PictureBox1.Size = new System.Drawing.Size(128, 100);
            this.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.PictureBox1.TabIndex = 165;
            this.PictureBox1.TabStop = false;
            // 
            // grpDevSetting1
            // 
            this.grpDevSetting1.Controls.Add(this.Label51);
            this.grpDevSetting1.Controls.Add(this.Label360);
            this.grpDevSetting1.Controls.Add(this.chkKyNewest);
            this.grpDevSetting1.Controls.Add(this.chkSyskanriinit);
            this.grpDevSetting1.Controls.Add(this.Label46);
            this.grpDevSetting1.Controls.Add(this.Label45);
            this.grpDevSetting1.Controls.Add(this.Label44);
            this.grpDevSetting1.Controls.Add(this.chkRelTblDrop);
            this.grpDevSetting1.Controls.Add(this.Label43);
            this.grpDevSetting1.Controls.Add(this.Label42);
            this.grpDevSetting1.Controls.Add(this.chkV7ViewDrop);
            this.grpDevSetting1.Controls.Add(this.chkLogTblDrop);
            this.grpDevSetting1.Controls.Add(this.chkOverWrite);
            this.grpDevSetting1.Controls.Add(this.Label52);
            this.grpDevSetting1.Location = new System.Drawing.Point(23, 27);
            this.grpDevSetting1.Name = "grpDevSetting1";
            this.grpDevSetting1.Size = new System.Drawing.Size(726, 191);
            this.grpDevSetting1.TabIndex = 164;
            this.grpDevSetting1.TabStop = false;
            this.grpDevSetting1.Text = "【開発用設定1】";
            // 
            // Label51
            // 
            this.Label51.Location = new System.Drawing.Point(240, 156);
            this.Label51.Name = "Label51";
            this.Label51.Size = new System.Drawing.Size(465, 22);
            this.Label51.TabIndex = 164;
            this.Label51.Text = "解約されていない最新の契約情報のみをコンバートします。(V7受託CVと同等)";
            this.Label51.UseCompatibleTextRendering = true;
            // 
            // Label360
            // 
            this.Label360.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label360.Location = new System.Drawing.Point(240, 135);
            this.Label360.Name = "Label360";
            this.Label360.Size = new System.Drawing.Size(444, 22);
            this.Label360.TabIndex = 166;
            this.Label360.Text = "賃貸革命10の初期設定(XML)値を初期化します。";
            this.Label360.UseCompatibleTextRendering = true;
            // 
            // chkKyNewest
            // 
            this.chkKyNewest.Location = new System.Drawing.Point(25, 155);
            this.chkKyNewest.Name = "chkKyNewest";
            this.chkKyNewest.Size = new System.Drawing.Size(195, 22);
            this.chkKyNewest.TabIndex = 163;
            this.chkKyNewest.Text = "最新契約情報のみ対象";
            this.chkKyNewest.UseVisualStyleBackColor = true;
            // 
            // chkSyskanriinit
            // 
            this.chkSyskanriinit.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkSyskanriinit.ForeColor = System.Drawing.Color.Black;
            this.chkSyskanriinit.Location = new System.Drawing.Point(25, 133);
            this.chkSyskanriinit.Name = "chkSyskanriinit";
            this.chkSyskanriinit.Size = new System.Drawing.Size(195, 22);
            this.chkSyskanriinit.TabIndex = 161;
            this.chkSyskanriinit.Text = "初期設定初期化";
            this.chkSyskanriinit.UseVisualStyleBackColor = true;
            // 
            // Label46
            // 
            this.Label46.Location = new System.Drawing.Point(25, 69);
            this.Label46.Name = "Label46";
            this.Label46.Size = new System.Drawing.Size(195, 20);
            this.Label46.TabIndex = 160;
            this.Label46.Text = "● 仮TBL削除 (10DB作成)";
            this.Label46.UseCompatibleTextRendering = true;
            // 
            // Label45
            // 
            this.Label45.Location = new System.Drawing.Point(240, 69);
            this.Label45.Name = "Label45";
            this.Label45.Size = new System.Drawing.Size(444, 22);
            this.Label45.TabIndex = 159;
            this.Label45.Text = "賃貸革命10へコンバートする為に作成した以下の作業用TBLを削除します。";
            this.Label45.UseCompatibleTextRendering = true;
            // 
            // Label44
            // 
            this.Label44.Location = new System.Drawing.Point(240, 92);
            this.Label44.Name = "Label44";
            this.Label44.Size = new System.Drawing.Size(444, 24);
            this.Label44.TabIndex = 158;
            this.Label44.Text = "作業用として作成した紐付設定TBLを削除(ドロップ)します。";
            this.Label44.UseCompatibleTextRendering = true;
            // 
            // chkRelTblDrop
            // 
            this.chkRelTblDrop.AutoSize = true;
            this.chkRelTblDrop.Location = new System.Drawing.Point(45, 91);
            this.chkRelTblDrop.Name = "chkRelTblDrop";
            this.chkRelTblDrop.Size = new System.Drawing.Size(111, 22);
            this.chkRelTblDrop.TabIndex = 157;
            this.chkRelTblDrop.Text = "紐付設定データ";
            this.chkRelTblDrop.UseVisualStyleBackColor = true;
            // 
            // Label43
            // 
            this.Label43.Location = new System.Drawing.Point(240, 46);
            this.Label43.Name = "Label43";
            this.Label43.Size = new System.Drawing.Size(444, 24);
            this.Label43.TabIndex = 156;
            this.Label43.Text = "賃貸革命V7から必要項目を抽出する為に作成したVIEWを削除します。";
            this.Label43.UseCompatibleTextRendering = true;
            // 
            // Label42
            // 
            this.Label42.Location = new System.Drawing.Point(240, 113);
            this.Label42.Name = "Label42";
            this.Label42.Size = new System.Drawing.Size(444, 22);
            this.Label42.TabIndex = 155;
            this.Label42.Text = "作業用として作成したコンバートログTBLを削除(ドロップ)します。";
            this.Label42.UseCompatibleTextRendering = true;
            // 
            // chkV7ViewDrop
            // 
            this.chkV7ViewDrop.AutoSize = true;
            this.chkV7ViewDrop.Location = new System.Drawing.Point(25, 46);
            this.chkV7ViewDrop.Name = "chkV7ViewDrop";
            this.chkV7ViewDrop.Size = new System.Drawing.Size(196, 22);
            this.chkV7ViewDrop.TabIndex = 154;
            this.chkV7ViewDrop.Text = "仮作成VIEWを削除 (V7側作成)";
            this.chkV7ViewDrop.UseVisualStyleBackColor = true;
            // 
            // chkLogTblDrop
            // 
            this.chkLogTblDrop.AutoSize = true;
            this.chkLogTblDrop.Location = new System.Drawing.Point(45, 112);
            this.chkLogTblDrop.Name = "chkLogTblDrop";
            this.chkLogTblDrop.Size = new System.Drawing.Size(87, 22);
            this.chkLogTblDrop.TabIndex = 153;
            this.chkLogTblDrop.Text = "ログデータ";
            this.chkLogTblDrop.UseVisualStyleBackColor = true;
            // 
            // chkOverWrite
            // 
            this.chkOverWrite.AutoSize = true;
            this.chkOverWrite.Location = new System.Drawing.Point(25, 25);
            this.chkOverWrite.Name = "chkOverWrite";
            this.chkOverWrite.Size = new System.Drawing.Size(123, 22);
            this.chkOverWrite.TabIndex = 152;
            this.chkOverWrite.Text = "上書きコンバート";
            this.chkOverWrite.UseVisualStyleBackColor = true;
            // 
            // Label52
            // 
            this.Label52.Location = new System.Drawing.Point(240, 25);
            this.Label52.Name = "Label52";
            this.Label52.Size = new System.Drawing.Size(448, 22);
            this.Label52.TabIndex = 151;
            this.Label52.Text = "該当するデータが既に登録されている場合、上書きでコンバートを行います。";
            this.Label52.UseCompatibleTextRendering = true;
            // 
            // grpDevSetting2
            // 
            this.grpDevSetting2.Controls.Add(this.btnRelDirSeach);
            this.grpDevSetting2.Controls.Add(this.txtRelationDirPath);
            this.grpDevSetting2.Controls.Add(this.lblRelationDir);
            this.grpDevSetting2.Controls.Add(this.Label11);
            this.grpDevSetting2.Controls.Add(this.chkRelation);
            this.grpDevSetting2.Controls.Add(this.Label47);
            this.grpDevSetting2.Controls.Add(this.chkMidNotStop);
            this.grpDevSetting2.Controls.Add(this.chkCVStart);
            this.grpDevSetting2.Controls.Add(this.Label8);
            this.grpDevSetting2.Controls.Add(this.Label10);
            this.grpDevSetting2.Controls.Add(this.chkMiddleFile);
            this.grpDevSetting2.Location = new System.Drawing.Point(23, 224);
            this.grpDevSetting2.Name = "grpDevSetting2";
            this.grpDevSetting2.Size = new System.Drawing.Size(611, 165);
            this.grpDevSetting2.TabIndex = 163;
            this.grpDevSetting2.TabStop = false;
            this.grpDevSetting2.Text = "【開発用設定2】";
            // 
            // btnRelDirSeach
            // 
            this.btnRelDirSeach.Location = new System.Drawing.Point(526, 130);
            this.btnRelDirSeach.Name = "btnRelDirSeach";
            this.btnRelDirSeach.Size = new System.Drawing.Size(30, 25);
            this.btnRelDirSeach.TabIndex = 168;
            this.btnRelDirSeach.Text = "...";
            this.btnRelDirSeach.UseVisualStyleBackColor = true;
            this.btnRelDirSeach.Click += new System.EventHandler(this.btnDirSeach_Click);
            // 
            // txtRelationDirPath
            // 
            this.txtRelationDirPath.Location = new System.Drawing.Point(175, 130);
            this.txtRelationDirPath.Name = "txtRelationDirPath";
            this.txtRelationDirPath.Size = new System.Drawing.Size(345, 25);
            this.txtRelationDirPath.TabIndex = 167;
            // 
            // lblRelationDir
            // 
            this.lblRelationDir.Location = new System.Drawing.Point(25, 133);
            this.lblRelationDir.Name = "lblRelationDir";
            this.lblRelationDir.Size = new System.Drawing.Size(144, 21);
            this.lblRelationDir.TabIndex = 166;
            this.lblRelationDir.Text = "● 紐付ファイル格納先";
            this.lblRelationDir.UseCompatibleTextRendering = true;
            // 
            // Label11
            // 
            this.Label11.Location = new System.Drawing.Point(175, 89);
            this.Label11.Name = "Label11";
            this.Label11.Size = new System.Drawing.Size(386, 38);
            this.Label11.TabIndex = 165;
            this.Label11.Text = "入金項目などの移行元と移行先のデータの紐付けを行います。\r\n既に紐付け処理を行っている場合はチェックをOFFにして下さい。";
            this.Label11.UseCompatibleTextRendering = true;
            // 
            // chkRelation
            // 
            this.chkRelation.AutoSize = true;
            this.chkRelation.Location = new System.Drawing.Point(25, 88);
            this.chkRelation.Name = "chkRelation";
            this.chkRelation.Size = new System.Drawing.Size(111, 22);
            this.chkRelation.TabIndex = 164;
            this.chkRelation.Text = "項目紐付け設定";
            this.chkRelation.UseVisualStyleBackColor = true;
            // 
            // Label47
            // 
            this.Label47.Location = new System.Drawing.Point(175, 68);
            this.Label47.Name = "Label47";
            this.Label47.Size = new System.Drawing.Size(376, 21);
            this.Label47.TabIndex = 144;
            this.Label47.Text = "移行元→中間→移行先を連続で行います。";
            this.Label47.UseCompatibleTextRendering = true;
            // 
            // chkMidNotStop
            // 
            this.chkMidNotStop.AutoSize = true;
            this.chkMidNotStop.Location = new System.Drawing.Point(25, 67);
            this.chkMidNotStop.Name = "chkMidNotStop";
            this.chkMidNotStop.Size = new System.Drawing.Size(75, 22);
            this.chkMidNotStop.TabIndex = 143;
            this.chkMidNotStop.Text = "連続実行";
            this.chkMidNotStop.UseVisualStyleBackColor = true;
            // 
            // chkCVStart
            // 
            this.chkCVStart.AutoSize = true;
            this.chkCVStart.Location = new System.Drawing.Point(25, 46);
            this.chkCVStart.Name = "chkCVStart";
            this.chkCVStart.Size = new System.Drawing.Size(111, 22);
            this.chkCVStart.TabIndex = 133;
            this.chkCVStart.Text = "コンバート実施";
            this.chkCVStart.UseVisualStyleBackColor = true;
            // 
            // Label8
            // 
            this.Label8.Location = new System.Drawing.Point(175, 47);
            this.Label8.Name = "Label8";
            this.Label8.Size = new System.Drawing.Size(345, 21);
            this.Label8.TabIndex = 132;
            this.Label8.Text = "コンバートを実行します。";
            this.Label8.UseCompatibleTextRendering = true;
            // 
            // Label10
            // 
            this.Label10.Location = new System.Drawing.Point(175, 27);
            this.Label10.Name = "Label10";
            this.Label10.Size = new System.Drawing.Size(365, 20);
            this.Label10.TabIndex = 131;
            this.Label10.Text = "移行元のデータを中間ファイルに出力します。";
            this.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label10.UseCompatibleTextRendering = true;
            // 
            // chkMiddleFile
            // 
            this.chkMiddleFile.AutoSize = true;
            this.chkMiddleFile.Location = new System.Drawing.Point(25, 25);
            this.chkMiddleFile.Name = "chkMiddleFile";
            this.chkMiddleFile.Size = new System.Drawing.Size(123, 22);
            this.chkMiddleFile.TabIndex = 130;
            this.chkMiddleFile.Text = "中間ファイル出力";
            this.chkMiddleFile.UseVisualStyleBackColor = true;
            // 
            // grpDevSetting3
            // 
            this.grpDevSetting3.Controls.Add(this.Label55);
            this.grpDevSetting3.Controls.Add(this.Label54);
            this.grpDevSetting3.Controls.Add(this.chkRommDuplicate);
            this.grpDevSetting3.Controls.Add(this.chkEmptyRoomNo);
            this.grpDevSetting3.Controls.Add(this.Label53);
            this.grpDevSetting3.Location = new System.Drawing.Point(23, 400);
            this.grpDevSetting3.Name = "grpDevSetting3";
            this.grpDevSetting3.Size = new System.Drawing.Size(632, 144);
            this.grpDevSetting3.TabIndex = 145;
            this.grpDevSetting3.TabStop = false;
            this.grpDevSetting3.Text = "【開発用設定3】";
            // 
            // Label55
            // 
            this.Label55.Location = new System.Drawing.Point(240, 71);
            this.Label55.Name = "Label55";
            this.Label55.Size = new System.Drawing.Size(386, 22);
            this.Label55.TabIndex = 160;
            this.Label55.Text = "重複部屋番号が存在する場合、任意の部屋番号に変更します。";
            this.Label55.UseCompatibleTextRendering = true;
            // 
            // Label54
            // 
            this.Label54.Location = new System.Drawing.Point(240, 50);
            this.Label54.Name = "Label54";
            this.Label54.Size = new System.Drawing.Size(386, 22);
            this.Label54.TabIndex = 159;
            this.Label54.Text = "移行元の部屋番号が未設定の場合、任意の部屋番号を設定します。";
            this.Label54.UseCompatibleTextRendering = true;
            // 
            // chkRommDuplicate
            // 
            this.chkRommDuplicate.Location = new System.Drawing.Point(45, 71);
            this.chkRommDuplicate.Name = "chkRommDuplicate";
            this.chkRommDuplicate.Size = new System.Drawing.Size(171, 22);
            this.chkRommDuplicate.TabIndex = 158;
            this.chkRommDuplicate.Text = "重複部屋番号変更設定";
            this.chkRommDuplicate.UseVisualStyleBackColor = true;
            // 
            // chkEmptyRoomNo
            // 
            this.chkEmptyRoomNo.AutoSize = true;
            this.chkEmptyRoomNo.Location = new System.Drawing.Point(45, 49);
            this.chkEmptyRoomNo.Name = "chkEmptyRoomNo";
            this.chkEmptyRoomNo.Size = new System.Drawing.Size(171, 22);
            this.chkEmptyRoomNo.TabIndex = 157;
            this.chkEmptyRoomNo.Text = "未設定部屋番号初期値設定";
            this.chkEmptyRoomNo.UseVisualStyleBackColor = true;
            // 
            // Label53
            // 
            this.Label53.Location = new System.Drawing.Point(25, 26);
            this.Label53.Name = "Label53";
            this.Label53.Size = new System.Drawing.Size(131, 20);
            this.Label53.TabIndex = 156;
            this.Label53.Text = "● 部屋番号設定";
            this.Label53.UseCompatibleTextRendering = true;
            // 
            // grpRelation
            // 
            this.grpRelation.Location = new System.Drawing.Point(391, 423);
            this.grpRelation.Name = "grpRelation";
            this.grpRelation.Size = new System.Drawing.Size(64, 72);
            this.grpRelation.TabIndex = 158;
            this.grpRelation.TabStop = false;
            this.grpRelation.Text = "紐付設定";
            this.grpRelation.Visible = false;
            // 
            // GroupBox3
            // 
            this.GroupBox3.Controls.Add(this.Label39);
            this.GroupBox3.Controls.Add(this.Label33);
            this.GroupBox3.Controls.Add(this.Label6);
            this.GroupBox3.Controls.Add(this.Label56);
            this.GroupBox3.Controls.Add(this.Label16);
            this.GroupBox3.Controls.Add(this.Label15);
            this.GroupBox3.Controls.Add(this.Label14);
            this.GroupBox3.Controls.Add(this.Label12);
            this.GroupBox3.Controls.Add(this.Label18);
            this.GroupBox3.Location = new System.Drawing.Point(301, 406);
            this.GroupBox3.Name = "GroupBox3";
            this.GroupBox3.Size = new System.Drawing.Size(77, 68);
            this.GroupBox3.TabIndex = 157;
            this.GroupBox3.TabStop = false;
            this.GroupBox3.Text = "退避4";
            this.GroupBox3.Visible = false;
            // 
            // Label39
            // 
            this.Label39.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label39.Location = new System.Drawing.Point(40, -6);
            this.Label39.Name = "Label39";
            this.Label39.Size = new System.Drawing.Size(32, 96);
            this.Label39.TabIndex = 136;
            this.Label39.Text = "├\r\n|\r\n├\r\n|\r\n├\r\n|\r\n└";
            this.Label39.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label33
            // 
            this.Label33.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label33.Location = new System.Drawing.Point(44, -14);
            this.Label33.Name = "Label33";
            this.Label33.Size = new System.Drawing.Size(32, 64);
            this.Label33.TabIndex = 135;
            this.Label33.Text = "├\r\n|\r\n├\r\n|\r\n└";
            this.Label33.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label6
            // 
            this.Label6.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label6.Location = new System.Drawing.Point(28, -14);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(32, 40);
            this.Label6.TabIndex = 134;
            this.Label6.Text = "├\r\n|\r\n└";
            this.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label56
            // 
            this.Label56.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label56.Location = new System.Drawing.Point(4, -26);
            this.Label56.Name = "Label56";
            this.Label56.Size = new System.Drawing.Size(32, 16);
            this.Label56.TabIndex = 133;
            this.Label56.Text = "└";
            this.Label56.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Label16
            // 
            this.Label16.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label16.Location = new System.Drawing.Point(-180, -26);
            this.Label16.Name = "Label16";
            this.Label16.Size = new System.Drawing.Size(32, 16);
            this.Label16.TabIndex = 132;
            this.Label16.Text = "└";
            this.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Label15
            // 
            this.Label15.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label15.Location = new System.Drawing.Point(112, 32);
            this.Label15.Name = "Label15";
            this.Label15.Size = new System.Drawing.Size(32, 40);
            this.Label15.TabIndex = 131;
            this.Label15.Text = "├\r\n|\r\n└";
            this.Label15.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label14
            // 
            this.Label14.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label14.Location = new System.Drawing.Point(16, 32);
            this.Label14.Name = "Label14";
            this.Label14.Size = new System.Drawing.Size(32, 40);
            this.Label14.TabIndex = 130;
            this.Label14.Text = "├\r\n|\r\n└";
            this.Label14.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label12
            // 
            this.Label12.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label12.Location = new System.Drawing.Point(64, 32);
            this.Label12.Name = "Label12";
            this.Label12.Size = new System.Drawing.Size(32, 40);
            this.Label12.TabIndex = 129;
            this.Label12.Text = "├\r\n|\r\n└";
            this.Label12.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label18
            // 
            this.Label18.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label18.Location = new System.Drawing.Point(16, 16);
            this.Label18.Name = "Label18";
            this.Label18.Size = new System.Drawing.Size(32, 16);
            this.Label18.TabIndex = 128;
            this.Label18.Text = "└";
            this.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.chkMstKasyorui);
            this.GroupBox2.Controls.Add(this.chkMstGenjotokuyaku);
            this.GroupBox2.Controls.Add(this.lblTokuyakuInfo);
            this.GroupBox2.Location = new System.Drawing.Point(183, 401);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(96, 103);
            this.GroupBox2.TabIndex = 156;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "退避3";
            this.GroupBox2.Visible = false;
            // 
            // chkMstKasyorui
            // 
            this.chkMstKasyorui.AutoSize = true;
            this.chkMstKasyorui.Location = new System.Drawing.Point(8, 112);
            this.chkMstKasyorui.Name = "chkMstKasyorui";
            this.chkMstKasyorui.Size = new System.Drawing.Size(111, 22);
            this.chkMstKasyorui.TabIndex = 127;
            this.chkMstKasyorui.Text = "箇所分類マスタ";
            this.chkMstKasyorui.UseVisualStyleBackColor = true;
            // 
            // chkMstGenjotokuyaku
            // 
            this.chkMstGenjotokuyaku.AutoSize = true;
            this.chkMstGenjotokuyaku.Location = new System.Drawing.Point(8, 88);
            this.chkMstGenjotokuyaku.Name = "chkMstGenjotokuyaku";
            this.chkMstGenjotokuyaku.Size = new System.Drawing.Size(135, 22);
            this.chkMstGenjotokuyaku.TabIndex = 126;
            this.chkMstGenjotokuyaku.Text = "原状回復特約マスタ";
            this.chkMstGenjotokuyaku.UseVisualStyleBackColor = true;
            // 
            // lblTokuyakuInfo
            // 
            this.lblTokuyakuInfo.Location = new System.Drawing.Point(-72, 22);
            this.lblTokuyakuInfo.Name = "lblTokuyakuInfo";
            this.lblTokuyakuInfo.Size = new System.Drawing.Size(344, 56);
            this.lblTokuyakuInfo.TabIndex = 124;
            this.lblTokuyakuInfo.Text = "※1.特約事項マスタおよび原状回復特約マスタは特約事項内容設定へ\r\n      統合されます。\r\n※2.箇所分類マスタおよびクレーム分類マスタはクレーム管理設定へ" +
    "\r\n      統合されます。";
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.Label34);
            this.GroupBox1.Controls.Add(this.Label29);
            this.GroupBox1.Controls.Add(this.chkMstKozasyubetu);
            this.GroupBox1.Controls.Add(this.chkMstSetubi);
            this.GroupBox1.Controls.Add(this.chkMstKozo);
            this.GroupBox1.Controls.Add(this.chkMstTorihikitaiyo);
            this.GroupBox1.Controls.Add(this.chkMstNkinkbn);
            this.GroupBox1.Controls.Add(this.chkMstHyrui);
            this.GroupBox1.Controls.Add(this.chkMstNkinkomok);
            this.GroupBox1.Controls.Add(this.chkMstBkrui);
            this.GroupBox1.Location = new System.Drawing.Point(71, 407);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(88, 40);
            this.GroupBox1.TabIndex = 124;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "退避2";
            this.GroupBox1.Visible = false;
            // 
            // Label34
            // 
            this.Label34.Location = new System.Drawing.Point(48, 216);
            this.Label34.Name = "Label34";
            this.Label34.Size = new System.Drawing.Size(160, 16);
            this.Label34.TabIndex = 128;
            this.Label34.Text = "項目別進捗";
            this.Label34.UseCompatibleTextRendering = true;
            // 
            // Label29
            // 
            this.Label29.Location = new System.Drawing.Point(40, 176);
            this.Label29.Name = "Label29";
            this.Label29.Size = new System.Drawing.Size(160, 16);
            this.Label29.TabIndex = 127;
            this.Label29.Text = "全体進捗";
            this.Label29.UseCompatibleTextRendering = true;
            // 
            // chkMstKozasyubetu
            // 
            this.chkMstKozasyubetu.AutoSize = true;
            this.chkMstKozasyubetu.Location = new System.Drawing.Point(167, 45);
            this.chkMstKozasyubetu.Name = "chkMstKozasyubetu";
            this.chkMstKozasyubetu.Size = new System.Drawing.Size(111, 22);
            this.chkMstKozasyubetu.TabIndex = 126;
            this.chkMstKozasyubetu.Text = "口座種別マスタ";
            this.chkMstKozasyubetu.UseVisualStyleBackColor = true;
            // 
            // chkMstSetubi
            // 
            this.chkMstSetubi.AutoSize = true;
            this.chkMstSetubi.Location = new System.Drawing.Point(167, 132);
            this.chkMstSetubi.Name = "chkMstSetubi";
            this.chkMstSetubi.Size = new System.Drawing.Size(87, 22);
            this.chkMstSetubi.TabIndex = 125;
            this.chkMstSetubi.Text = "設備マスタ";
            this.chkMstSetubi.UseVisualStyleBackColor = true;
            // 
            // chkMstKozo
            // 
            this.chkMstKozo.AutoSize = true;
            this.chkMstKozo.Location = new System.Drawing.Point(39, 103);
            this.chkMstKozo.Name = "chkMstKozo";
            this.chkMstKozo.Size = new System.Drawing.Size(87, 22);
            this.chkMstKozo.TabIndex = 124;
            this.chkMstKozo.Text = "構造マスタ";
            this.chkMstKozo.UseVisualStyleBackColor = true;
            // 
            // chkMstTorihikitaiyo
            // 
            this.chkMstTorihikitaiyo.AutoSize = true;
            this.chkMstTorihikitaiyo.Location = new System.Drawing.Point(39, 132);
            this.chkMstTorihikitaiyo.Name = "chkMstTorihikitaiyo";
            this.chkMstTorihikitaiyo.Size = new System.Drawing.Size(111, 22);
            this.chkMstTorihikitaiyo.TabIndex = 123;
            this.chkMstTorihikitaiyo.Text = "取引態様マスタ";
            this.chkMstTorihikitaiyo.UseVisualStyleBackColor = true;
            // 
            // chkMstNkinkbn
            // 
            this.chkMstNkinkbn.AutoSize = true;
            this.chkMstNkinkbn.Location = new System.Drawing.Point(167, 74);
            this.chkMstNkinkbn.Name = "chkMstNkinkbn";
            this.chkMstNkinkbn.Size = new System.Drawing.Size(111, 22);
            this.chkMstNkinkbn.TabIndex = 122;
            this.chkMstNkinkbn.Text = "入金区分マスタ";
            this.chkMstNkinkbn.UseVisualStyleBackColor = true;
            // 
            // chkMstHyrui
            // 
            this.chkMstHyrui.AutoSize = true;
            this.chkMstHyrui.Location = new System.Drawing.Point(39, 74);
            this.chkMstHyrui.Name = "chkMstHyrui";
            this.chkMstHyrui.Size = new System.Drawing.Size(111, 22);
            this.chkMstHyrui.TabIndex = 121;
            this.chkMstHyrui.Text = "部屋分類マスタ";
            this.chkMstHyrui.UseVisualStyleBackColor = true;
            // 
            // chkMstNkinkomok
            // 
            this.chkMstNkinkomok.AutoSize = true;
            this.chkMstNkinkomok.Location = new System.Drawing.Point(167, 103);
            this.chkMstNkinkomok.Name = "chkMstNkinkomok";
            this.chkMstNkinkomok.Size = new System.Drawing.Size(111, 22);
            this.chkMstNkinkomok.TabIndex = 120;
            this.chkMstNkinkomok.Text = "入金項目マスタ";
            this.chkMstNkinkomok.UseVisualStyleBackColor = true;
            // 
            // chkMstBkrui
            // 
            this.chkMstBkrui.AutoSize = true;
            this.chkMstBkrui.Location = new System.Drawing.Point(39, 45);
            this.chkMstBkrui.Name = "chkMstBkrui";
            this.chkMstBkrui.Size = new System.Drawing.Size(111, 22);
            this.chkMstBkrui.TabIndex = 119;
            this.chkMstBkrui.Text = "物件分類マスタ";
            this.chkMstBkrui.UseVisualStyleBackColor = true;
            // 
            // grptaihi
            // 
            this.grptaihi.Controls.Add(this.Label7);
            this.grptaihi.Controls.Add(this.btnFileSeach);
            this.grptaihi.Controls.Add(this.Label5);
            this.grptaihi.Controls.Add(this.chkGyYatinhosyoKoza);
            this.grptaihi.Controls.Add(this.CheckBox66);
            this.grptaihi.Controls.Add(this.lblGy);
            this.grptaihi.Controls.Add(this.lblMst);
            this.grptaihi.Controls.Add(this.lblKys);
            this.grptaihi.Controls.Add(this.lblOwner);
            this.grptaihi.Controls.Add(this.lblJisya);
            this.grptaihi.Controls.Add(this.Label17);
            this.grptaihi.Controls.Add(this.Label19);
            this.grptaihi.Controls.Add(this.Label20);
            this.grptaihi.Controls.Add(this.chkMs25);
            this.grptaihi.Controls.Add(this.chkMsKeiyakusyubetu);
            this.grptaihi.Controls.Add(this.chkMs23);
            this.grptaihi.Controls.Add(this.chkMs24);
            this.grptaihi.Controls.Add(this.chkMsOwevent);
            this.grptaihi.Controls.Add(this.chkMsYane);
            this.grptaihi.Controls.Add(this.chkMsKeikaikakunin);
            this.grptaihi.Controls.Add(this.chkMsKinyu);
            this.grptaihi.Controls.Add(this.chkMsKinyuten);
            this.grptaihi.Location = new System.Drawing.Point(71, 452);
            this.grptaihi.Name = "grptaihi";
            this.grptaihi.Size = new System.Drawing.Size(64, 48);
            this.grptaihi.TabIndex = 115;
            this.grptaihi.TabStop = false;
            this.grptaihi.Text = "退避";
            this.grptaihi.Visible = false;
            // 
            // Label7
            // 
            this.Label7.Location = new System.Drawing.Point(24, 136);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(150, 25);
            this.Label7.TabIndex = 126;
            this.Label7.Text = "中間ファイル登録データチェック";
            this.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Label7.UseCompatibleTextRendering = true;
            // 
            // btnFileSeach
            // 
            this.btnFileSeach.Location = new System.Drawing.Point(368, 128);
            this.btnFileSeach.Name = "btnFileSeach";
            this.btnFileSeach.Size = new System.Drawing.Size(27, 21);
            this.btnFileSeach.TabIndex = 125;
            this.btnFileSeach.Text = "...";
            this.btnFileSeach.UseVisualStyleBackColor = true;
            // 
            // Label5
            // 
            this.Label5.Location = new System.Drawing.Point(304, 96);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(374, 16);
            this.Label5.TabIndex = 117;
            this.Label5.Text = "既存データを上書きして移行する場合はチェックをONにして下さい。";
            this.Label5.UseCompatibleTextRendering = true;
            // 
            // chkGyYatinhosyoKoza
            // 
            this.chkGyYatinhosyoKoza.AutoSize = true;
            this.chkGyYatinhosyoKoza.Location = new System.Drawing.Point(129, 86);
            this.chkGyYatinhosyoKoza.Name = "chkGyYatinhosyoKoza";
            this.chkGyYatinhosyoKoza.Size = new System.Drawing.Size(147, 22);
            this.chkGyYatinhosyoKoza.TabIndex = 116;
            this.chkGyYatinhosyoKoza.Text = "家賃保証業者口座情報";
            this.chkGyYatinhosyoKoza.UseVisualStyleBackColor = true;
            // 
            // CheckBox66
            // 
            this.CheckBox66.AutoSize = true;
            this.CheckBox66.Location = new System.Drawing.Point(8, 16);
            this.CheckBox66.Name = "CheckBox66";
            this.CheckBox66.Size = new System.Drawing.Size(104, 22);
            this.CheckBox66.TabIndex = 115;
            this.CheckBox66.Text = "src内部に記載";
            this.CheckBox66.UseVisualStyleBackColor = true;
            // 
            // lblGy
            // 
            this.lblGy.AutoSize = true;
            this.lblGy.Location = new System.Drawing.Point(195, 171);
            this.lblGy.Name = "lblGy";
            this.lblGy.Size = new System.Drawing.Size(79, 18);
            this.lblGy.TabIndex = 113;
            this.lblGy.Text = "2.業者マスタ";
            // 
            // lblMst
            // 
            this.lblMst.AutoSize = true;
            this.lblMst.Location = new System.Drawing.Point(195, 178);
            this.lblMst.Name = "lblMst";
            this.lblMst.Size = new System.Drawing.Size(67, 18);
            this.lblMst.TabIndex = 114;
            this.lblMst.Text = "1.マスタ系";
            // 
            // lblKys
            // 
            this.lblKys.AutoSize = true;
            this.lblKys.Location = new System.Drawing.Point(185, 123);
            this.lblKys.Name = "lblKys";
            this.lblKys.Size = new System.Drawing.Size(79, 18);
            this.lblKys.TabIndex = 107;
            this.lblKys.Text = "5.契約者情報";
            // 
            // lblOwner
            // 
            this.lblOwner.AutoSize = true;
            this.lblOwner.Location = new System.Drawing.Point(185, 123);
            this.lblOwner.Name = "lblOwner";
            this.lblOwner.Size = new System.Drawing.Size(67, 18);
            this.lblOwner.TabIndex = 108;
            this.lblOwner.Text = "4.家主情報";
            // 
            // lblJisya
            // 
            this.lblJisya.AutoSize = true;
            this.lblJisya.Location = new System.Drawing.Point(195, 146);
            this.lblJisya.Name = "lblJisya";
            this.lblJisya.Size = new System.Drawing.Size(67, 18);
            this.lblJisya.TabIndex = 109;
            this.lblJisya.Text = "3.自社情報";
            // 
            // Label17
            // 
            this.Label17.AutoSize = true;
            this.Label17.Location = new System.Drawing.Point(185, 123);
            this.Label17.Name = "Label17";
            this.Label17.Size = new System.Drawing.Size(67, 18);
            this.Label17.TabIndex = 110;
            this.Label17.Text = "1.マスタ系";
            // 
            // Label19
            // 
            this.Label19.AutoSize = true;
            this.Label19.Location = new System.Drawing.Point(185, 123);
            this.Label19.Name = "Label19";
            this.Label19.Size = new System.Drawing.Size(67, 18);
            this.Label19.TabIndex = 111;
            this.Label19.Text = "1.マスタ系";
            // 
            // Label20
            // 
            this.Label20.AutoSize = true;
            this.Label20.Location = new System.Drawing.Point(185, 123);
            this.Label20.Name = "Label20";
            this.Label20.Size = new System.Drawing.Size(67, 18);
            this.Label20.TabIndex = 112;
            this.Label20.Text = "1.マスタ系";
            // 
            // chkMs25
            // 
            this.chkMs25.AutoSize = true;
            this.chkMs25.Location = new System.Drawing.Point(183, 56);
            this.chkMs25.Name = "chkMs25";
            this.chkMs25.Size = new System.Drawing.Size(104, 22);
            this.chkMs25.TabIndex = 106;
            this.chkMs25.Text = "src内部に記載";
            this.chkMs25.UseVisualStyleBackColor = true;
            // 
            // chkMsKeiyakusyubetu
            // 
            this.chkMsKeiyakusyubetu.AutoSize = true;
            this.chkMsKeiyakusyubetu.Location = new System.Drawing.Point(153, 104);
            this.chkMsKeiyakusyubetu.Name = "chkMsKeiyakusyubetu";
            this.chkMsKeiyakusyubetu.Size = new System.Drawing.Size(104, 22);
            this.chkMsKeiyakusyubetu.TabIndex = 104;
            this.chkMsKeiyakusyubetu.Text = "src内部に記載";
            this.chkMsKeiyakusyubetu.UseVisualStyleBackColor = true;
            // 
            // chkMs23
            // 
            this.chkMs23.AutoSize = true;
            this.chkMs23.Location = new System.Drawing.Point(183, 10);
            this.chkMs23.Name = "chkMs23";
            this.chkMs23.Size = new System.Drawing.Size(104, 22);
            this.chkMs23.TabIndex = 105;
            this.chkMs23.Text = "src内部に記載";
            this.chkMs23.UseVisualStyleBackColor = true;
            // 
            // chkMs24
            // 
            this.chkMs24.AutoSize = true;
            this.chkMs24.Location = new System.Drawing.Point(183, 34);
            this.chkMs24.Name = "chkMs24";
            this.chkMs24.Size = new System.Drawing.Size(104, 22);
            this.chkMs24.TabIndex = 104;
            this.chkMs24.Text = "src内部に記載";
            this.chkMs24.UseVisualStyleBackColor = true;
            // 
            // chkMsOwevent
            // 
            this.chkMsOwevent.AutoSize = true;
            this.chkMsOwevent.Location = new System.Drawing.Point(50, 18);
            this.chkMsOwevent.Name = "chkMsOwevent";
            this.chkMsOwevent.Size = new System.Drawing.Size(104, 22);
            this.chkMsOwevent.TabIndex = 82;
            this.chkMsOwevent.Text = "src内部に記載";
            this.chkMsOwevent.UseVisualStyleBackColor = true;
            // 
            // chkMsYane
            // 
            this.chkMsYane.AutoSize = true;
            this.chkMsYane.Location = new System.Drawing.Point(50, 40);
            this.chkMsYane.Name = "chkMsYane";
            this.chkMsYane.Size = new System.Drawing.Size(104, 22);
            this.chkMsYane.TabIndex = 93;
            this.chkMsYane.Text = "src内部に記載";
            this.chkMsYane.UseVisualStyleBackColor = true;
            // 
            // chkMsKeikaikakunin
            // 
            this.chkMsKeikaikakunin.AutoSize = true;
            this.chkMsKeikaikakunin.Location = new System.Drawing.Point(50, 62);
            this.chkMsKeikaikakunin.Name = "chkMsKeikaikakunin";
            this.chkMsKeikaikakunin.Size = new System.Drawing.Size(104, 22);
            this.chkMsKeikaikakunin.TabIndex = 84;
            this.chkMsKeikaikakunin.Text = "src内部に記載";
            this.chkMsKeikaikakunin.UseVisualStyleBackColor = true;
            // 
            // chkMsKinyu
            // 
            this.chkMsKinyu.AutoSize = true;
            this.chkMsKinyu.Location = new System.Drawing.Point(50, 104);
            this.chkMsKinyu.Name = "chkMsKinyu";
            this.chkMsKinyu.Size = new System.Drawing.Size(104, 22);
            this.chkMsKinyu.TabIndex = 95;
            this.chkMsKinyu.Text = "src内部に記載";
            this.chkMsKinyu.UseVisualStyleBackColor = true;
            // 
            // chkMsKinyuten
            // 
            this.chkMsKinyuten.AutoSize = true;
            this.chkMsKinyuten.Location = new System.Drawing.Point(50, 82);
            this.chkMsKinyuten.Name = "chkMsKinyuten";
            this.chkMsKinyuten.Size = new System.Drawing.Size(104, 22);
            this.chkMsKinyuten.TabIndex = 96;
            this.chkMsKinyuten.Text = "src内部に記載";
            this.chkMsKinyuten.UseVisualStyleBackColor = true;
            // 
            // tabPageStart
            // 
            this.tabPageStart.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageStart.Controls.Add(this.lblFirstDescription1);
            this.tabPageStart.Controls.Add(this.lblHFirstLabel);
            this.tabPageStart.Controls.Add(this.grpHFirstNaiyo);
            this.tabPageStart.Controls.Add(this.PictureBox2);
            this.tabPageStart.Location = new System.Drawing.Point(4, 27);
            this.tabPageStart.Name = "tabPageStart";
            this.tabPageStart.Size = new System.Drawing.Size(912, 549);
            this.tabPageStart.TabIndex = 13;
            this.tabPageStart.Text = "開始";
            // 
            // lblFirstDescription1
            // 
            this.lblFirstDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblFirstDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblFirstDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblFirstDescription1.Location = new System.Drawing.Point(35, 49);
            this.lblFirstDescription1.Name = "lblFirstDescription1";
            this.lblFirstDescription1.Size = new System.Drawing.Size(820, 35);
            this.lblFirstDescription1.TabIndex = 4;
            this.lblFirstDescription1.Text = "「次へ」ボタンを押して下さい。";
            this.lblFirstDescription1.UseCompatibleTextRendering = true;
            // 
            // lblHFirstLabel
            // 
            this.lblHFirstLabel.BackColor = System.Drawing.SystemColors.Menu;
            this.lblHFirstLabel.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHFirstLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lblHFirstLabel.Location = new System.Drawing.Point(35, 20);
            this.lblHFirstLabel.Name = "lblHFirstLabel";
            this.lblHFirstLabel.Size = new System.Drawing.Size(820, 80);
            this.lblHFirstLabel.TabIndex = 149;
            this.lblHFirstLabel.Text = "「中間テーブル」から「賃貸革命10」のコンバート作業を行います。";
            this.lblHFirstLabel.UseCompatibleTextRendering = true;
            // 
            // grpHFirstNaiyo
            // 
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH312);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH311);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH342);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH322);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH341);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH321);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH302);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH202);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH102);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH301);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH201);
            this.grpHFirstNaiyo.Controls.Add(this.lblHajimeH101);
            this.grpHFirstNaiyo.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpHFirstNaiyo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.grpHFirstNaiyo.Location = new System.Drawing.Point(65, 115);
            this.grpHFirstNaiyo.Name = "grpHFirstNaiyo";
            this.grpHFirstNaiyo.Size = new System.Drawing.Size(767, 259);
            this.grpHFirstNaiyo.TabIndex = 149;
            this.grpHFirstNaiyo.TabStop = false;
            this.grpHFirstNaiyo.Text = "【コンバート作業内容】";
            // 
            // lblHajimeH312
            // 
            this.lblHajimeH312.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH312.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH312.Location = new System.Drawing.Point(290, 114);
            this.lblHajimeH312.Name = "lblHajimeH312";
            this.lblHajimeH312.Size = new System.Drawing.Size(454, 20);
            this.lblHajimeH312.TabIndex = 64;
            this.lblHajimeH312.Text = "… コンバートを実行する前の作業を行います。";
            this.lblHajimeH312.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH311
            // 
            this.lblHajimeH311.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH311.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH311.Location = new System.Drawing.Point(124, 114);
            this.lblHajimeH311.Name = "lblHajimeH311";
            this.lblHajimeH311.Size = new System.Drawing.Size(160, 20);
            this.lblHajimeH311.TabIndex = 63;
            this.lblHajimeH311.Text = "　　・事前作業";
            this.lblHajimeH311.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH342
            // 
            this.lblHajimeH342.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH342.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH342.Location = new System.Drawing.Point(290, 168);
            this.lblHajimeH342.Name = "lblHajimeH342";
            this.lblHajimeH342.Size = new System.Drawing.Size(454, 20);
            this.lblHajimeH342.TabIndex = 62;
            this.lblHajimeH342.Text = "… コンバートを実行した後の「賃貸革命10」データの調整を行います。";
            this.lblHajimeH342.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH322
            // 
            this.lblHajimeH322.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH322.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH322.Location = new System.Drawing.Point(290, 141);
            this.lblHajimeH322.Name = "lblHajimeH322";
            this.lblHajimeH322.Size = new System.Drawing.Size(454, 20);
            this.lblHajimeH322.TabIndex = 60;
            this.lblHajimeH322.Text = "… 中間ファイルのチェックとデータコンバートを行います。";
            this.lblHajimeH322.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH341
            // 
            this.lblHajimeH341.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH341.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH341.Location = new System.Drawing.Point(124, 168);
            this.lblHajimeH341.Name = "lblHajimeH341";
            this.lblHajimeH341.Size = new System.Drawing.Size(160, 20);
            this.lblHajimeH341.TabIndex = 58;
            this.lblHajimeH341.Text = "　　・事後作業";
            this.lblHajimeH341.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH321
            // 
            this.lblHajimeH321.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH321.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH321.Location = new System.Drawing.Point(124, 141);
            this.lblHajimeH321.Name = "lblHajimeH321";
            this.lblHajimeH321.Size = new System.Drawing.Size(160, 20);
            this.lblHajimeH321.TabIndex = 56;
            this.lblHajimeH321.Text = "　　・データコンバート";
            this.lblHajimeH321.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH302
            // 
            this.lblHajimeH302.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH302.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH302.Location = new System.Drawing.Point(124, 85);
            this.lblHajimeH302.Name = "lblHajimeH302";
            this.lblHajimeH302.Size = new System.Drawing.Size(620, 20);
            this.lblHajimeH302.TabIndex = 54;
            this.lblHajimeH302.Text = "… コンバート作業を行います。(以下の内容から選択)";
            this.lblHajimeH302.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH202
            // 
            this.lblHajimeH202.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH202.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH202.Location = new System.Drawing.Point(124, 57);
            this.lblHajimeH202.Name = "lblHajimeH202";
            this.lblHajimeH202.Size = new System.Drawing.Size(620, 20);
            this.lblHajimeH202.TabIndex = 53;
            this.lblHajimeH202.Text = "… 「ログファイル格納パス」の設定など、簡単な初期設定を行います。";
            this.lblHajimeH202.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH102
            // 
            this.lblHajimeH102.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH102.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH102.Location = new System.Drawing.Point(124, 28);
            this.lblHajimeH102.Name = "lblHajimeH102";
            this.lblHajimeH102.Size = new System.Drawing.Size(620, 20);
            this.lblHajimeH102.TabIndex = 52;
            this.lblHajimeH102.Text = "… 「賃貸革命10」のデータベースの接続設定を行います。\r\n";
            this.lblHajimeH102.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH301
            // 
            this.lblHajimeH301.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH301.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH301.Location = new System.Drawing.Point(18, 85);
            this.lblHajimeH301.Name = "lblHajimeH301";
            this.lblHajimeH301.Size = new System.Drawing.Size(100, 20);
            this.lblHajimeH301.TabIndex = 51;
            this.lblHajimeH301.Text = "　作業選択";
            this.lblHajimeH301.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH201
            // 
            this.lblHajimeH201.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH201.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH201.Location = new System.Drawing.Point(18, 57);
            this.lblHajimeH201.Name = "lblHajimeH201";
            this.lblHajimeH201.Size = new System.Drawing.Size(100, 20);
            this.lblHajimeH201.TabIndex = 50;
            this.lblHajimeH201.Text = "　初期設定";
            this.lblHajimeH201.UseCompatibleTextRendering = true;
            // 
            // lblHajimeH101
            // 
            this.lblHajimeH101.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHajimeH101.ForeColor = System.Drawing.Color.Black;
            this.lblHajimeH101.Location = new System.Drawing.Point(18, 29);
            this.lblHajimeH101.Name = "lblHajimeH101";
            this.lblHajimeH101.Size = new System.Drawing.Size(100, 20);
            this.lblHajimeH101.TabIndex = 49;
            this.lblHajimeH101.Text = "　接続設定";
            this.lblHajimeH101.UseCompatibleTextRendering = true;
            // 
            // PictureBox2
            // 
            this.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.PictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("PictureBox2.Image")));
            this.PictureBox2.Location = new System.Drawing.Point(698, 390);
            this.PictureBox2.Name = "PictureBox2";
            this.PictureBox2.Size = new System.Drawing.Size(191, 138);
            this.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PictureBox2.TabIndex = 147;
            this.PictureBox2.TabStop = false;
            // 
            // tabPageSession
            // 
            this.tabPageSession.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageSession.Controls.Add(this.label2);
            this.tabPageSession.Controls.Add(this.pre_table_connection_string);
            this.tabPageSession.Controls.Add(this.label1);
            this.tabPageSession.Controls.Add(this.grpTimeOut);
            this.tabPageSession.Controls.Add(this.lblSessionCaution);
            this.tabPageSession.Controls.Add(this.lblHSessionDescription1);
            this.tabPageSession.Controls.Add(this.Label36);
            this.tabPageSession.Controls.Add(this.Label35);
            this.tabPageSession.Controls.Add(this.lblNetworklib);
            this.tabPageSession.Controls.Add(this.Label30);
            this.tabPageSession.Controls.Add(this.Label13);
            this.tabPageSession.Controls.Add(this.grp10ConnectInfo);
            this.tabPageSession.Controls.Add(this.btnDefConInfoRead);
            this.tabPageSession.Controls.Add(this.btnConnectTest);
            this.tabPageSession.Location = new System.Drawing.Point(4, 27);
            this.tabPageSession.Name = "tabPageSession";
            this.tabPageSession.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSession.Size = new System.Drawing.Size(912, 549);
            this.tabPageSession.TabIndex = 1;
            this.tabPageSession.Text = " 接続設定";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(119, 521);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(373, 18);
            this.label2.TabIndex = 18;
            this.label2.Text = "空白にすると、【移行先接続情報】指定のDBから読み込まれます。";
            // 
            // pre_table_connection_string
            // 
            this.pre_table_connection_string.Location = new System.Drawing.Point(122, 495);
            this.pre_table_connection_string.Name = "pre_table_connection_string";
            this.pre_table_connection_string.Size = new System.Drawing.Size(504, 25);
            this.pre_table_connection_string.TabIndex = 17;
            this.pre_table_connection_string.Text = "Persist Security Info=True;Data Source = ;Initial Catalog = ;User ID = ;Password " +
    "= ;Connection Timeout = 40";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 497);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 18);
            this.label1.TabIndex = 16;
            this.label1.Text = "Preテーブル読込元";
            // 
            // grpTimeOut
            // 
            this.grpTimeOut.Controls.Add(this.txtTimeOut);
            this.grpTimeOut.Controls.Add(this.lblTimeOutSec);
            this.grpTimeOut.Location = new System.Drawing.Point(769, 20);
            this.grpTimeOut.Name = "grpTimeOut";
            this.grpTimeOut.Size = new System.Drawing.Size(118, 64);
            this.grpTimeOut.TabIndex = 14;
            this.grpTimeOut.TabStop = false;
            this.grpTimeOut.Text = " タイムアウト値 ";
            // 
            // txtTimeOut
            // 
            this.txtTimeOut.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTimeOut.Location = new System.Drawing.Point(33, 24);
            this.txtTimeOut.Name = "txtTimeOut";
            this.txtTimeOut.Size = new System.Drawing.Size(43, 27);
            this.txtTimeOut.TabIndex = 10;
            this.txtTimeOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblTimeOutSec
            // 
            this.lblTimeOutSec.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblTimeOutSec.Location = new System.Drawing.Point(82, 28);
            this.lblTimeOutSec.Name = "lblTimeOutSec";
            this.lblTimeOutSec.Size = new System.Drawing.Size(29, 19);
            this.lblTimeOutSec.TabIndex = 11;
            this.lblTimeOutSec.Text = "秒";
            this.lblTimeOutSec.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblTimeOutSec.UseCompatibleTextRendering = true;
            // 
            // lblSessionCaution
            // 
            this.lblSessionCaution.BackColor = System.Drawing.SystemColors.Menu;
            this.lblSessionCaution.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblSessionCaution.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.lblSessionCaution.Location = new System.Drawing.Point(30, 90);
            this.lblSessionCaution.Name = "lblSessionCaution";
            this.lblSessionCaution.Size = new System.Drawing.Size(820, 24);
            this.lblSessionCaution.TabIndex = 1;
            this.lblSessionCaution.Text = "※正常接続が確認できない場合、[次へ] に進むことはできません。(正常接続できない場合はサポートへお問い合わせ下さい)";
            this.lblSessionCaution.UseCompatibleTextRendering = true;
            // 
            // lblHSessionDescription1
            // 
            this.lblHSessionDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblHSessionDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHSessionDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblHSessionDescription1.Location = new System.Drawing.Point(30, 21);
            this.lblHSessionDescription1.Name = "lblHSessionDescription1";
            this.lblHSessionDescription1.Size = new System.Drawing.Size(820, 80);
            this.lblHSessionDescription1.TabIndex = 15;
            this.lblHSessionDescription1.Text = "データベースの接続設定を行います。([初期値読込] を押すと接続情報の初期設定値を読み込みます)\r\n「賃貸革命10」の接続情報を設定し、[接続確認] ボタンを押し" +
    "て下さい。\r\n正常接続が確認できたら [次へ] ボタンを押して下さい。";
            this.lblHSessionDescription1.UseCompatibleTextRendering = true;
            // 
            // Label36
            // 
            this.Label36.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label36.Location = new System.Drawing.Point(34, 401);
            this.Label36.Name = "Label36";
            this.Label36.Size = new System.Drawing.Size(140, 20);
            this.Label36.TabIndex = 5;
            this.Label36.Text = "パスワード";
            this.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Label36.UseCompatibleTextRendering = true;
            // 
            // Label35
            // 
            this.Label35.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label35.Location = new System.Drawing.Point(34, 361);
            this.Label35.Name = "Label35";
            this.Label35.Size = new System.Drawing.Size(140, 20);
            this.Label35.TabIndex = 4;
            this.Label35.Text = "ログイン";
            this.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Label35.UseCompatibleTextRendering = true;
            // 
            // lblNetworklib
            // 
            this.lblNetworklib.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNetworklib.Location = new System.Drawing.Point(34, 444);
            this.lblNetworklib.Name = "lblNetworklib";
            this.lblNetworklib.Size = new System.Drawing.Size(140, 20);
            this.lblNetworklib.TabIndex = 6;
            this.lblNetworklib.Text = "ネットワークライブラリ";
            this.lblNetworklib.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblNetworklib.UseCompatibleTextRendering = true;
            // 
            // Label30
            // 
            this.Label30.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label30.Location = new System.Drawing.Point(34, 281);
            this.Label30.Name = "Label30";
            this.Label30.Size = new System.Drawing.Size(140, 20);
            this.Label30.TabIndex = 2;
            this.Label30.Text = "サーバー名";
            this.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Label30.UseCompatibleTextRendering = true;
            // 
            // Label13
            // 
            this.Label13.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label13.Location = new System.Drawing.Point(34, 321);
            this.Label13.Name = "Label13";
            this.Label13.Size = new System.Drawing.Size(140, 20);
            this.Label13.TabIndex = 3;
            this.Label13.Text = "カタログ名";
            this.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Label13.UseCompatibleTextRendering = true;
            // 
            // grp10ConnectInfo
            // 
            this.grp10ConnectInfo.Controls.Add(this.cmbV10Networklib);
            this.grp10ConnectInfo.Controls.Add(this.grpV10Authent);
            this.grp10ConnectInfo.Controls.Add(this.txtV10Pass);
            this.grp10ConnectInfo.Controls.Add(this.txtV10User);
            this.grp10ConnectInfo.Controls.Add(this.txtV10Catalog);
            this.grp10ConnectInfo.Controls.Add(this.txtV10Server);
            this.grp10ConnectInfo.Controls.Add(this.lblV10);
            this.grp10ConnectInfo.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grp10ConnectInfo.Location = new System.Drawing.Point(190, 158);
            this.grp10ConnectInfo.Name = "grp10ConnectInfo";
            this.grp10ConnectInfo.Size = new System.Drawing.Size(300, 330);
            this.grp10ConnectInfo.TabIndex = 8;
            this.grp10ConnectInfo.TabStop = false;
            this.grp10ConnectInfo.Text = "【移行先接続情報】";
            // 
            // cmbV10Networklib
            // 
            this.cmbV10Networklib.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbV10Networklib.FormattingEnabled = true;
            this.cmbV10Networklib.Location = new System.Drawing.Point(10, 280);
            this.cmbV10Networklib.Name = "cmbV10Networklib";
            this.cmbV10Networklib.Size = new System.Drawing.Size(280, 28);
            this.cmbV10Networklib.TabIndex = 7;
            // 
            // grpV10Authent
            // 
            this.grpV10Authent.Controls.Add(this.optV10Authent2);
            this.grpV10Authent.Controls.Add(this.optV10Authent1);
            this.grpV10Authent.Location = new System.Drawing.Point(6, 56);
            this.grpV10Authent.Name = "grpV10Authent";
            this.grpV10Authent.Size = new System.Drawing.Size(288, 50);
            this.grpV10Authent.TabIndex = 1;
            this.grpV10Authent.TabStop = false;
            // 
            // optV10Authent2
            // 
            this.optV10Authent2.AutoSize = true;
            this.optV10Authent2.Location = new System.Drawing.Point(160, 18);
            this.optV10Authent2.Name = "optV10Authent2";
            this.optV10Authent2.Size = new System.Drawing.Size(114, 24);
            this.optV10Authent2.TabIndex = 1;
            this.optV10Authent2.Text = "Windows 認証";
            this.optV10Authent2.UseVisualStyleBackColor = true;
            this.optV10Authent2.CheckedChanged += new System.EventHandler(this.opt_CheckedChanged);
            // 
            // optV10Authent1
            // 
            this.optV10Authent1.AutoSize = true;
            this.optV10Authent1.Checked = true;
            this.optV10Authent1.Location = new System.Drawing.Point(12, 18);
            this.optV10Authent1.Name = "optV10Authent1";
            this.optV10Authent1.Size = new System.Drawing.Size(123, 24);
            this.optV10Authent1.TabIndex = 0;
            this.optV10Authent1.TabStop = true;
            this.optV10Authent1.Text = "SQLServer 認証";
            this.optV10Authent1.UseVisualStyleBackColor = true;
            this.optV10Authent1.CheckedChanged += new System.EventHandler(this.opt_CheckedChanged);
            // 
            // txtV10Pass
            // 
            this.txtV10Pass.Location = new System.Drawing.Point(10, 238);
            this.txtV10Pass.Name = "txtV10Pass";
            this.txtV10Pass.Size = new System.Drawing.Size(280, 27);
            this.txtV10Pass.TabIndex = 5;
            this.txtV10Pass.Text = "***************";
            // 
            // txtV10User
            // 
            this.txtV10User.Location = new System.Drawing.Point(10, 201);
            this.txtV10User.Name = "txtV10User";
            this.txtV10User.Size = new System.Drawing.Size(280, 27);
            this.txtV10User.TabIndex = 4;
            this.txtV10User.Text = "sa";
            // 
            // txtV10Catalog
            // 
            this.txtV10Catalog.Location = new System.Drawing.Point(10, 160);
            this.txtV10Catalog.Name = "txtV10Catalog";
            this.txtV10Catalog.Size = new System.Drawing.Size(280, 27);
            this.txtV10Catalog.TabIndex = 3;
            this.txtV10Catalog.Text = "fk8db";
            // 
            // txtV10Server
            // 
            this.txtV10Server.Location = new System.Drawing.Point(10, 120);
            this.txtV10Server.Name = "txtV10Server";
            this.txtV10Server.Size = new System.Drawing.Size(280, 27);
            this.txtV10Server.TabIndex = 2;
            this.txtV10Server.Text = "PC-NJC\\SQL2012";
            // 
            // lblV10
            // 
            this.lblV10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblV10.Image = ((System.Drawing.Image)(resources.GetObject("lblV10.Image")));
            this.lblV10.Location = new System.Drawing.Point(6, 30);
            this.lblV10.Name = "lblV10";
            this.lblV10.Size = new System.Drawing.Size(288, 24);
            this.lblV10.TabIndex = 0;
            this.lblV10.Text = "賃貸革命10";
            this.lblV10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnDefConInfoRead
            // 
            this.btnDefConInfoRead.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDefConInfoRead.Image = ((System.Drawing.Image)(resources.GetObject("btnDefConInfoRead.Image")));
            this.btnDefConInfoRead.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDefConInfoRead.Location = new System.Drawing.Point(632, 506);
            this.btnDefConInfoRead.Name = "btnDefConInfoRead";
            this.btnDefConInfoRead.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnDefConInfoRead.Size = new System.Drawing.Size(120, 30);
            this.btnDefConInfoRead.TabIndex = 12;
            this.btnDefConInfoRead.Text = "設定値再読込";
            this.btnDefConInfoRead.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnDefConInfoRead.UseVisualStyleBackColor = true;
            this.btnDefConInfoRead.Click += new System.EventHandler(this.btnDefConInfoRead_Click);
            // 
            // btnConnectTest
            // 
            this.btnConnectTest.Font = new System.Drawing.Font("メイリオ", 11.25F);
            this.btnConnectTest.Image = ((System.Drawing.Image)(resources.GetObject("btnConnectTest.Image")));
            this.btnConnectTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnConnectTest.Location = new System.Drawing.Point(769, 506);
            this.btnConnectTest.Name = "btnConnectTest";
            this.btnConnectTest.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnConnectTest.Size = new System.Drawing.Size(120, 30);
            this.btnConnectTest.TabIndex = 13;
            this.btnConnectTest.Text = " 接続確認";
            this.btnConnectTest.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnConnectTest.UseVisualStyleBackColor = true;
            this.btnConnectTest.Click += new System.EventHandler(this.btnConnectTest_Click);
            // 
            // tabPageSyoki
            // 
            this.tabPageSyoki.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageSyoki.Controls.Add(this.pnlUserName);
            this.tabPageSyoki.Controls.Add(this.pnlLogPath);
            this.tabPageSyoki.Controls.Add(this.lblSyokiDescription1);
            this.tabPageSyoki.Location = new System.Drawing.Point(4, 27);
            this.tabPageSyoki.Name = "tabPageSyoki";
            this.tabPageSyoki.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSyoki.Size = new System.Drawing.Size(912, 549);
            this.tabPageSyoki.TabIndex = 0;
            this.tabPageSyoki.Text = " 初期設定";
            // 
            // pnlUserName
            // 
            this.pnlUserName.Controls.Add(this.lblExecutor);
            this.pnlUserName.Controls.Add(this.txtRecUser);
            this.pnlUserName.Controls.Add(this.Label58);
            this.pnlUserName.Location = new System.Drawing.Point(68, 215);
            this.pnlUserName.Name = "pnlUserName";
            this.pnlUserName.Size = new System.Drawing.Size(414, 82);
            this.pnlUserName.TabIndex = 22;
            // 
            // lblExecutor
            // 
            this.lblExecutor.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblExecutor.Location = new System.Drawing.Point(3, 6);
            this.lblExecutor.Name = "lblExecutor";
            this.lblExecutor.Size = new System.Drawing.Size(128, 19);
            this.lblExecutor.TabIndex = 16;
            this.lblExecutor.Text = "作業者名";
            this.lblExecutor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblExecutor.UseCompatibleTextRendering = true;
            // 
            // txtRecUser
            // 
            this.txtRecUser.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtRecUser.Location = new System.Drawing.Point(23, 46);
            this.txtRecUser.Name = "txtRecUser";
            this.txtRecUser.Size = new System.Drawing.Size(384, 25);
            this.txtRecUser.TabIndex = 18;
            // 
            // Label58
            // 
            this.Label58.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label58.Location = new System.Drawing.Point(23, 25);
            this.Label58.Name = "Label58";
            this.Label58.Size = new System.Drawing.Size(384, 20);
            this.Label58.TabIndex = 17;
            this.Label58.Text = "設定値は、コンバートしたデータベースの履歴に登録されます。";
            this.Label58.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label58.UseCompatibleTextRendering = true;
            // 
            // pnlLogPath
            // 
            this.pnlLogPath.Controls.Add(this.lblLogDir);
            this.pnlLogPath.Controls.Add(this.Label57);
            this.pnlLogPath.Controls.Add(this.btnLogDirSeach);
            this.pnlLogPath.Controls.Add(this.txtLogDirPath);
            this.pnlLogPath.Location = new System.Drawing.Point(68, 107);
            this.pnlLogPath.Name = "pnlLogPath";
            this.pnlLogPath.Size = new System.Drawing.Size(414, 82);
            this.pnlLogPath.TabIndex = 21;
            // 
            // lblLogDir
            // 
            this.lblLogDir.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblLogDir.Location = new System.Drawing.Point(3, 6);
            this.lblLogDir.Name = "lblLogDir";
            this.lblLogDir.Size = new System.Drawing.Size(128, 19);
            this.lblLogDir.TabIndex = 12;
            this.lblLogDir.Text = "ログファイル格納先";
            this.lblLogDir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblLogDir.UseCompatibleTextRendering = true;
            // 
            // Label57
            // 
            this.Label57.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label57.Location = new System.Drawing.Point(23, 25);
            this.Label57.Name = "Label57";
            this.Label57.Size = new System.Drawing.Size(384, 20);
            this.Label57.TabIndex = 13;
            this.Label57.Text = "コンバート結果を出力するファイルの格納先を設定して下さい。";
            this.Label57.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label57.UseCompatibleTextRendering = true;
            // 
            // btnLogDirSeach
            // 
            this.btnLogDirSeach.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnLogDirSeach.Location = new System.Drawing.Point(377, 46);
            this.btnLogDirSeach.Name = "btnLogDirSeach";
            this.btnLogDirSeach.Size = new System.Drawing.Size(30, 25);
            this.btnLogDirSeach.TabIndex = 15;
            this.btnLogDirSeach.Text = "...";
            this.btnLogDirSeach.UseVisualStyleBackColor = true;
            this.btnLogDirSeach.Click += new System.EventHandler(this.btnDirSeach_Click);
            // 
            // txtLogDirPath
            // 
            this.txtLogDirPath.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtLogDirPath.Location = new System.Drawing.Point(23, 46);
            this.txtLogDirPath.Name = "txtLogDirPath";
            this.txtLogDirPath.Size = new System.Drawing.Size(345, 25);
            this.txtLogDirPath.TabIndex = 14;
            // 
            // lblSyokiDescription1
            // 
            this.lblSyokiDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblSyokiDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblSyokiDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblSyokiDescription1.Location = new System.Drawing.Point(30, 20);
            this.lblSyokiDescription1.Name = "lblSyokiDescription1";
            this.lblSyokiDescription1.Size = new System.Drawing.Size(820, 80);
            this.lblSyokiDescription1.TabIndex = 0;
            this.lblSyokiDescription1.Text = "コンバーターの初期設定を行います。\r\n設定が完了したら「次へ」ボタンを押してください。";
            this.lblSyokiDescription1.UseCompatibleTextRendering = true;
            // 
            // tabPageMenu
            // 
            this.tabPageMenu.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageMenu.Controls.Add(this.lblMenuCaution);
            this.tabPageMenu.Controls.Add(this.grpMenuJizen);
            this.tabPageMenu.Controls.Add(this.grpMenuDatacv);
            this.tabPageMenu.Controls.Add(this.grpMenuJigo);
            this.tabPageMenu.Controls.Add(this.lblMenuDescription1);
            this.tabPageMenu.Location = new System.Drawing.Point(4, 27);
            this.tabPageMenu.Name = "tabPageMenu";
            this.tabPageMenu.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMenu.Size = new System.Drawing.Size(912, 549);
            this.tabPageMenu.TabIndex = 12;
            this.tabPageMenu.Text = "tabPageSelect";
            // 
            // lblMenuCaution
            // 
            this.lblMenuCaution.BackColor = System.Drawing.SystemColors.Menu;
            this.lblMenuCaution.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblMenuCaution.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.lblMenuCaution.Location = new System.Drawing.Point(30, 63);
            this.lblMenuCaution.Name = "lblMenuCaution";
            this.lblMenuCaution.Size = new System.Drawing.Size(701, 24);
            this.lblMenuCaution.TabIndex = 147;
            this.lblMenuCaution.Text = "※先に作業を完了させないと選択できない項目があります。";
            this.lblMenuCaution.UseCompatibleTextRendering = true;
            // 
            // grpMenuJizen
            // 
            this.grpMenuJizen.Controls.Add(this.lblKMenuJizen);
            this.grpMenuJizen.Controls.Add(this.lblHMenuJizen);
            this.grpMenuJizen.Controls.Add(this.btnMenuJizen);
            this.grpMenuJizen.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpMenuJizen.ForeColor = System.Drawing.Color.Navy;
            this.grpMenuJizen.Location = new System.Drawing.Point(53, 108);
            this.grpMenuJizen.Name = "grpMenuJizen";
            this.grpMenuJizen.Size = new System.Drawing.Size(691, 80);
            this.grpMenuJizen.TabIndex = 144;
            this.grpMenuJizen.TabStop = false;
            this.grpMenuJizen.Text = " 1. 事前作業 ";
            // 
            // lblKMenuJizen
            // 
            this.lblKMenuJizen.BackColor = System.Drawing.SystemColors.Menu;
            this.lblKMenuJizen.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblKMenuJizen.ForeColor = System.Drawing.Color.Black;
            this.lblKMenuJizen.Location = new System.Drawing.Point(148, 19);
            this.lblKMenuJizen.Name = "lblKMenuJizen";
            this.lblKMenuJizen.Size = new System.Drawing.Size(518, 55);
            this.lblKMenuJizen.TabIndex = 9;
            this.lblKMenuJizen.Text = "コンバートを実行する前の賃貸革命V7データの調整を行います。\r\n調整が必要と思われる内容のリストを出力し、それを元に調整を行って下さい。\r\n※コンバートを行う前に" +
    "必ず行って下さい。";
            this.lblKMenuJizen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblKMenuJizen.UseCompatibleTextRendering = true;
            // 
            // lblHMenuJizen
            // 
            this.lblHMenuJizen.BackColor = System.Drawing.SystemColors.Menu;
            this.lblHMenuJizen.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblHMenuJizen.ForeColor = System.Drawing.Color.Black;
            this.lblHMenuJizen.Location = new System.Drawing.Point(148, 19);
            this.lblHMenuJizen.Name = "lblHMenuJizen";
            this.lblHMenuJizen.Size = new System.Drawing.Size(518, 55);
            this.lblHMenuJizen.TabIndex = 11;
            this.lblHMenuJizen.Text = "中間ファイルの作成を行います。";
            this.lblHMenuJizen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblHMenuJizen.UseCompatibleTextRendering = true;
            // 
            // btnMenuJizen
            // 
            this.btnMenuJizen.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMenuJizen.ForeColor = System.Drawing.Color.Black;
            this.btnMenuJizen.Location = new System.Drawing.Point(32, 30);
            this.btnMenuJizen.Name = "btnMenuJizen";
            this.btnMenuJizen.Size = new System.Drawing.Size(101, 33);
            this.btnMenuJizen.TabIndex = 10;
            this.btnMenuJizen.Text = "事前調整";
            this.btnMenuJizen.UseVisualStyleBackColor = true;
            this.btnMenuJizen.Click += new System.EventHandler(this.btnMenuJizen_Click);
            // 
            // grpMenuDatacv
            // 
            this.grpMenuDatacv.Controls.Add(this.lblKMenuDatacv);
            this.grpMenuDatacv.Controls.Add(this.lblHMenuDatacv);
            this.grpMenuDatacv.Controls.Add(this.btnMenuDatacv);
            this.grpMenuDatacv.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpMenuDatacv.ForeColor = System.Drawing.Color.Navy;
            this.grpMenuDatacv.Location = new System.Drawing.Point(53, 194);
            this.grpMenuDatacv.Name = "grpMenuDatacv";
            this.grpMenuDatacv.Size = new System.Drawing.Size(686, 80);
            this.grpMenuDatacv.TabIndex = 143;
            this.grpMenuDatacv.TabStop = false;
            this.grpMenuDatacv.Text = " 2. データコンバート ";
            // 
            // lblKMenuDatacv
            // 
            this.lblKMenuDatacv.BackColor = System.Drawing.SystemColors.Menu;
            this.lblKMenuDatacv.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblKMenuDatacv.ForeColor = System.Drawing.Color.Black;
            this.lblKMenuDatacv.Location = new System.Drawing.Point(148, 19);
            this.lblKMenuDatacv.Name = "lblKMenuDatacv";
            this.lblKMenuDatacv.Size = new System.Drawing.Size(513, 55);
            this.lblKMenuDatacv.TabIndex = 12;
            this.lblKMenuDatacv.Text = "データコンバートを行います。\r\n※\"事前調整\"を先に完了させておく必要があります。";
            this.lblKMenuDatacv.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblKMenuDatacv.UseCompatibleTextRendering = true;
            // 
            // lblHMenuDatacv
            // 
            this.lblHMenuDatacv.BackColor = System.Drawing.SystemColors.Menu;
            this.lblHMenuDatacv.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblHMenuDatacv.ForeColor = System.Drawing.Color.Black;
            this.lblHMenuDatacv.Location = new System.Drawing.Point(148, 19);
            this.lblHMenuDatacv.Name = "lblHMenuDatacv";
            this.lblHMenuDatacv.Size = new System.Drawing.Size(513, 55);
            this.lblHMenuDatacv.TabIndex = 14;
            this.lblHMenuDatacv.Text = "中間ファイルのチェックとデータコンバートを行います。";
            this.lblHMenuDatacv.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblHMenuDatacv.UseCompatibleTextRendering = true;
            // 
            // btnMenuDatacv
            // 
            this.btnMenuDatacv.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMenuDatacv.ForeColor = System.Drawing.Color.Black;
            this.btnMenuDatacv.Location = new System.Drawing.Point(32, 30);
            this.btnMenuDatacv.Name = "btnMenuDatacv";
            this.btnMenuDatacv.Size = new System.Drawing.Size(101, 33);
            this.btnMenuDatacv.TabIndex = 13;
            this.btnMenuDatacv.Text = "移行実行";
            this.btnMenuDatacv.UseVisualStyleBackColor = true;
            this.btnMenuDatacv.Click += new System.EventHandler(this.btnMenuDatacv_Click);
            // 
            // grpMenuJigo
            // 
            this.grpMenuJigo.Controls.Add(this.btnMenuJigo);
            this.grpMenuJigo.Controls.Add(this.lblMenuJigo);
            this.grpMenuJigo.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpMenuJigo.ForeColor = System.Drawing.Color.Navy;
            this.grpMenuJigo.Location = new System.Drawing.Point(53, 280);
            this.grpMenuJigo.Name = "grpMenuJigo";
            this.grpMenuJigo.Size = new System.Drawing.Size(686, 80);
            this.grpMenuJigo.TabIndex = 141;
            this.grpMenuJigo.TabStop = false;
            this.grpMenuJigo.Text = " 3. 事後作業 ";
            // 
            // btnMenuJigo
            // 
            this.btnMenuJigo.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMenuJigo.ForeColor = System.Drawing.Color.Black;
            this.btnMenuJigo.Location = new System.Drawing.Point(31, 30);
            this.btnMenuJigo.Name = "btnMenuJigo";
            this.btnMenuJigo.Size = new System.Drawing.Size(101, 33);
            this.btnMenuJigo.TabIndex = 19;
            this.btnMenuJigo.Text = "事後調整";
            this.btnMenuJigo.UseVisualStyleBackColor = true;
            this.btnMenuJigo.Click += new System.EventHandler(this.btnMenuJigo_Click);
            // 
            // lblMenuJigo
            // 
            this.lblMenuJigo.BackColor = System.Drawing.SystemColors.Menu;
            this.lblMenuJigo.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblMenuJigo.ForeColor = System.Drawing.Color.Black;
            this.lblMenuJigo.Location = new System.Drawing.Point(148, 19);
            this.lblMenuJigo.Name = "lblMenuJigo";
            this.lblMenuJigo.Size = new System.Drawing.Size(513, 55);
            this.lblMenuJigo.TabIndex = 18;
            this.lblMenuJigo.Text = "コンバートを実行した後の賃貸革命10データの調整を行います。";
            this.lblMenuJigo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblMenuJigo.UseCompatibleTextRendering = true;
            // 
            // lblMenuDescription1
            // 
            this.lblMenuDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblMenuDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblMenuDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblMenuDescription1.Location = new System.Drawing.Point(30, 20);
            this.lblMenuDescription1.Name = "lblMenuDescription1";
            this.lblMenuDescription1.Size = new System.Drawing.Size(820, 35);
            this.lblMenuDescription1.TabIndex = 140;
            this.lblMenuDescription1.Text = "以下の内容を確認し、必要な作業を行って下さい。\r\n(ボタンを押して下さい)";
            this.lblMenuDescription1.UseCompatibleTextRendering = true;
            // 
            // tabPageJizen
            // 
            this.tabPageJizen.BackColor = System.Drawing.SystemColors.MenuBar;
            this.tabPageJizen.Controls.Add(this.pnlJizenListPath);
            this.tabPageJizen.Controls.Add(this.lblLine1);
            this.tabPageJizen.Controls.Add(this.lblHidden3);
            this.tabPageJizen.Controls.Add(this.tabCtrlJizen);
            this.tabPageJizen.Controls.Add(this.lblCautionDescription);
            this.tabPageJizen.Controls.Add(this.lblJizenDescription1);
            this.tabPageJizen.Location = new System.Drawing.Point(4, 27);
            this.tabPageJizen.Name = "tabPageJizen";
            this.tabPageJizen.Size = new System.Drawing.Size(912, 549);
            this.tabPageJizen.TabIndex = 7;
            this.tabPageJizen.Text = " 事前調整作業";
            // 
            // pnlJizenListPath
            // 
            this.pnlJizenListPath.Controls.Add(this.Label254);
            this.pnlJizenListPath.Controls.Add(this.btnJizenListDirSeach);
            this.pnlJizenListPath.Controls.Add(this.txtJizenListPath);
            this.pnlJizenListPath.Location = new System.Drawing.Point(31, 86);
            this.pnlJizenListPath.Name = "pnlJizenListPath";
            this.pnlJizenListPath.Size = new System.Drawing.Size(420, 60);
            this.pnlJizenListPath.TabIndex = 47;
            // 
            // Label254
            // 
            this.Label254.AutoSize = true;
            this.Label254.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label254.Location = new System.Drawing.Point(5, 5);
            this.Label254.Name = "Label254";
            this.Label254.Size = new System.Drawing.Size(116, 18);
            this.Label254.TabIndex = 3;
            this.Label254.Text = "出力ファイル格納先";
            // 
            // btnJizenListDirSeach
            // 
            this.btnJizenListDirSeach.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnJizenListDirSeach.Location = new System.Drawing.Point(383, 27);
            this.btnJizenListDirSeach.Name = "btnJizenListDirSeach";
            this.btnJizenListDirSeach.Size = new System.Drawing.Size(30, 25);
            this.btnJizenListDirSeach.TabIndex = 2;
            this.btnJizenListDirSeach.Text = "...";
            this.btnJizenListDirSeach.UseVisualStyleBackColor = true;
            this.btnJizenListDirSeach.Click += new System.EventHandler(this.btnDirSeach_Click);
            // 
            // txtJizenListPath
            // 
            this.txtJizenListPath.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtJizenListPath.Location = new System.Drawing.Point(23, 28);
            this.txtJizenListPath.Name = "txtJizenListPath";
            this.txtJizenListPath.Size = new System.Drawing.Size(354, 24);
            this.txtJizenListPath.TabIndex = 1;
            // 
            // lblLine1
            // 
            this.lblLine1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine1.Location = new System.Drawing.Point(31, 175);
            this.lblLine1.Name = "lblLine1";
            this.lblLine1.Size = new System.Drawing.Size(853, 2);
            this.lblLine1.TabIndex = 3;
            // 
            // lblHidden3
            // 
            this.lblHidden3.Location = new System.Drawing.Point(6, 120);
            this.lblHidden3.Name = "lblHidden3";
            this.lblHidden3.Size = new System.Drawing.Size(18, 55);
            this.lblHidden3.TabIndex = 2;
            this.lblHidden3.Text = "　";
            // 
            // tabCtrlJizen
            // 
            this.tabCtrlJizen.Controls.Add(this.tabPageHJizen1);
            this.tabCtrlJizen.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.tabCtrlJizen.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.tabCtrlJizen.Location = new System.Drawing.Point(30, 150);
            this.tabCtrlJizen.Name = "tabCtrlJizen";
            this.tabCtrlJizen.SelectedIndex = 0;
            this.tabCtrlJizen.Size = new System.Drawing.Size(855, 394);
            this.tabCtrlJizen.TabIndex = 4;
            this.tabCtrlJizen.TabStop = false;
            // 
            // tabPageHJizen1
            // 
            this.tabPageHJizen1.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHJizen1.Controls.Add(this.GroupBox5);
            this.tabPageHJizen1.Location = new System.Drawing.Point(4, 27);
            this.tabPageHJizen1.Name = "tabPageHJizen1";
            this.tabPageHJizen1.Size = new System.Drawing.Size(847, 363);
            this.tabPageHJizen1.TabIndex = 3;
            this.tabPageHJizen1.Text = " 事前作業1(汎用) ";
            // 
            // GroupBox5
            // 
            this.GroupBox5.BackColor = System.Drawing.SystemColors.Menu;
            this.GroupBox5.Controls.Add(this.pnlHJizenTyukan);
            this.GroupBox5.Controls.Add(this.Label355);
            this.GroupBox5.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox5.Location = new System.Drawing.Point(6, 15);
            this.GroupBox5.Name = "GroupBox5";
            this.GroupBox5.Size = new System.Drawing.Size(835, 342);
            this.GroupBox5.TabIndex = 0;
            this.GroupBox5.TabStop = false;
            this.GroupBox5.Text = "【ユーザによる事前確認・調整】";
            // 
            // pnlHJizenTyukan
            // 
            this.pnlHJizenTyukan.Controls.Add(this.txtMidDirPath);
            this.pnlHJizenTyukan.Controls.Add(this.chkHJizenTyukan);
            this.pnlHJizenTyukan.Location = new System.Drawing.Point(6, 18);
            this.pnlHJizenTyukan.Name = "pnlHJizenTyukan";
            this.pnlHJizenTyukan.Size = new System.Drawing.Size(400, 305);
            this.pnlHJizenTyukan.TabIndex = 0;
            // 
            // txtMidDirPath
            // 
            this.txtMidDirPath.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMidDirPath.Location = new System.Drawing.Point(48, 189);
            this.txtMidDirPath.Name = "txtMidDirPath";
            this.txtMidDirPath.Size = new System.Drawing.Size(256, 24);
            this.txtMidDirPath.TabIndex = 2;
            this.txtMidDirPath.Visible = false;
            // 
            // chkHJizenTyukan
            // 
            this.chkHJizenTyukan.AutoSize = true;
            this.chkHJizenTyukan.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkHJizenTyukan.Location = new System.Drawing.Point(16, 4);
            this.chkHJizenTyukan.Name = "chkHJizenTyukan";
            this.chkHJizenTyukan.Size = new System.Drawing.Size(217, 22);
            this.chkHJizenTyukan.TabIndex = 0;
            this.chkHJizenTyukan.Text = "pre_XXXXテーブルからの登録作業";
            this.chkHJizenTyukan.UseVisualStyleBackColor = true;
            this.chkHJizenTyukan.CheckedChanged += new System.EventHandler(this.chkHJizenTyukan_CheckedChanged);
            // 
            // Label355
            // 
            this.Label355.AutoSize = true;
            this.Label355.Location = new System.Drawing.Point(794, 321);
            this.Label355.Name = "Label355";
            this.Label355.Size = new System.Drawing.Size(35, 18);
            this.Label355.TabIndex = 10;
            this.Label355.Text = "1 / 3";
            // 
            // lblCautionDescription
            // 
            this.lblCautionDescription.BackColor = System.Drawing.SystemColors.Menu;
            this.lblCautionDescription.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblCautionDescription.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.lblCautionDescription.Location = new System.Drawing.Point(30, 63);
            this.lblCautionDescription.Name = "lblCautionDescription";
            this.lblCautionDescription.Size = new System.Drawing.Size(820, 20);
            this.lblCautionDescription.TabIndex = 1;
            this.lblCautionDescription.Text = "※データコンバートを実行する前に必ず対応下さい。対応しなかった場合、異なる期待結果となる場合があります。";
            this.lblCautionDescription.UseCompatibleTextRendering = true;
            // 
            // lblJizenDescription1
            // 
            this.lblJizenDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblJizenDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblJizenDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblJizenDescription1.Location = new System.Drawing.Point(30, 20);
            this.lblJizenDescription1.Name = "lblJizenDescription1";
            this.lblJizenDescription1.Size = new System.Drawing.Size(820, 80);
            this.lblJizenDescription1.TabIndex = 0;
            this.lblJizenDescription1.Text = "対象となる賃貸革命V7のデータを、事前に調整する必要があります。以下の内容を確認し、必要な作業を行って下さい。\r\n(完了したら各チェックボックスをONにして下さい" +
    "。全てONにした場合のみデータコンバートを行うことができます)\r\n";
            this.lblJizenDescription1.UseCompatibleTextRendering = true;
            // 
            // tabPageJigo
            // 
            this.tabPageJigo.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageJigo.Controls.Add(this.Label245);
            this.tabPageJigo.Controls.Add(this.lblLine3);
            this.tabPageJigo.Controls.Add(this.lblHidden4);
            this.tabPageJigo.Controls.Add(this.tabCtrlJigo);
            this.tabPageJigo.Controls.Add(this.lblJigoDescription1);
            this.tabPageJigo.Location = new System.Drawing.Point(4, 27);
            this.tabPageJigo.Name = "tabPageJigo";
            this.tabPageJigo.Size = new System.Drawing.Size(912, 549);
            this.tabPageJigo.TabIndex = 14;
            this.tabPageJigo.Text = "事後作業";
            // 
            // Label245
            // 
            this.Label245.BackColor = System.Drawing.SystemColors.Menu;
            this.Label245.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label245.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Label245.Location = new System.Drawing.Point(30, 63);
            this.Label245.Name = "Label245";
            this.Label245.Size = new System.Drawing.Size(820, 20);
            this.Label245.TabIndex = 12;
            this.Label245.Text = "※データコンバートを実行後に必ず確認をして下さい。必要な作業を行わなかった場合、異なる期待結果となる場合があります。";
            this.Label245.UseCompatibleTextRendering = true;
            // 
            // lblLine3
            // 
            this.lblLine3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine3.Location = new System.Drawing.Point(31, 175);
            this.lblLine3.Name = "lblLine3";
            this.lblLine3.Size = new System.Drawing.Size(853, 2);
            this.lblLine3.TabIndex = 11;
            // 
            // lblHidden4
            // 
            this.lblHidden4.Location = new System.Drawing.Point(6, 150);
            this.lblHidden4.Name = "lblHidden4";
            this.lblHidden4.Size = new System.Drawing.Size(10, 26);
            this.lblHidden4.TabIndex = 10;
            this.lblHidden4.Text = "　";
            // 
            // tabCtrlJigo
            // 
            this.tabCtrlJigo.Controls.Add(this.tabPageJigo1);
            this.tabCtrlJigo.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabCtrlJigo.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.tabCtrlJigo.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.tabCtrlJigo.Location = new System.Drawing.Point(30, 150);
            this.tabCtrlJigo.Name = "tabCtrlJigo";
            this.tabCtrlJigo.SelectedIndex = 0;
            this.tabCtrlJigo.Size = new System.Drawing.Size(855, 394);
            this.tabCtrlJigo.TabIndex = 8;
            this.tabCtrlJigo.TabStop = false;
            // 
            // tabPageJigo1
            // 
            this.tabPageJigo1.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageJigo1.Controls.Add(this.grpJigoDonyuji);
            this.tabPageJigo1.Controls.Add(this.grpJigoUserSagyo);
            this.tabPageJigo1.Location = new System.Drawing.Point(4, 27);
            this.tabPageJigo1.Name = "tabPageJigo1";
            this.tabPageJigo1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageJigo1.Size = new System.Drawing.Size(847, 363);
            this.tabPageJigo1.TabIndex = 4;
            this.tabPageJigo1.Text = " 事後作業1 ";
            // 
            // grpJigoDonyuji
            // 
            this.grpJigoDonyuji.BackColor = System.Drawing.SystemColors.Menu;
            this.grpJigoDonyuji.Controls.Add(this.pnlJigoCmtSoKotiku);
            this.grpJigoDonyuji.Controls.Add(this.pnlJigoCmtNkNyuryoku);
            this.grpJigoDonyuji.Controls.Add(this.pnlJigoCmtSqKotiku);
            this.grpJigoDonyuji.Cursor = System.Windows.Forms.Cursors.Default;
            this.grpJigoDonyuji.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpJigoDonyuji.Location = new System.Drawing.Point(436, 15);
            this.grpJigoDonyuji.Name = "grpJigoDonyuji";
            this.grpJigoDonyuji.Size = new System.Drawing.Size(405, 342);
            this.grpJigoDonyuji.TabIndex = 141;
            this.grpJigoDonyuji.TabStop = false;
            this.grpJigoDonyuji.Text = "【賃貸革命導入時通常作業】";
            // 
            // pnlJigoCmtSoKotiku
            // 
            this.pnlJigoCmtSoKotiku.Controls.Add(this.Label109);
            this.pnlJigoCmtSoKotiku.Controls.Add(this.Label108);
            this.pnlJigoCmtSoKotiku.Controls.Add(this.Label142);
            this.pnlJigoCmtSoKotiku.Location = new System.Drawing.Point(20, 173);
            this.pnlJigoCmtSoKotiku.Name = "pnlJigoCmtSoKotiku";
            this.pnlJigoCmtSoKotiku.Size = new System.Drawing.Size(375, 65);
            this.pnlJigoCmtSoKotiku.TabIndex = 140;
            // 
            // Label109
            // 
            this.Label109.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label109.ForeColor = System.Drawing.Color.Red;
            this.Label109.Location = new System.Drawing.Point(27, 4);
            this.Label109.Name = "Label109";
            this.Label109.Size = new System.Drawing.Size(340, 20);
            this.Label109.TabIndex = 123;
            this.Label109.Text = "送金データ一括構築 (送金処理)";
            this.Label109.UseCompatibleTextRendering = true;
            // 
            // Label108
            // 
            this.Label108.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label108.Location = new System.Drawing.Point(23, 28);
            this.Label108.Name = "Label108";
            this.Label108.Size = new System.Drawing.Size(340, 36);
            this.Label108.TabIndex = 124;
            this.Label108.Text = "入金処理後、送金データ一括構築を実施して下さい。データ構築を行わないと送金データが作成されません。";
            this.Label108.UseCompatibleTextRendering = true;
            // 
            // Label142
            // 
            this.Label142.ForeColor = System.Drawing.Color.Red;
            this.Label142.Location = new System.Drawing.Point(7, 4);
            this.Label142.Name = "Label142";
            this.Label142.Size = new System.Drawing.Size(20, 20);
            this.Label142.TabIndex = 135;
            this.Label142.Text = "・";
            this.Label142.UseCompatibleTextRendering = true;
            // 
            // pnlJigoCmtNkNyuryoku
            // 
            this.pnlJigoCmtNkNyuryoku.Controls.Add(this.Label99);
            this.pnlJigoCmtNkNyuryoku.Controls.Add(this.Label111);
            this.pnlJigoCmtNkNyuryoku.Controls.Add(this.Label110);
            this.pnlJigoCmtNkNyuryoku.Location = new System.Drawing.Point(20, 104);
            this.pnlJigoCmtNkNyuryoku.Name = "pnlJigoCmtNkNyuryoku";
            this.pnlJigoCmtNkNyuryoku.Size = new System.Drawing.Size(375, 55);
            this.pnlJigoCmtNkNyuryoku.TabIndex = 139;
            // 
            // Label99
            // 
            this.Label99.ForeColor = System.Drawing.Color.Red;
            this.Label99.Location = new System.Drawing.Point(3, 4);
            this.Label99.Name = "Label99";
            this.Label99.Size = new System.Drawing.Size(20, 20);
            this.Label99.TabIndex = 134;
            this.Label99.Text = "・";
            this.Label99.UseCompatibleTextRendering = true;
            // 
            // Label111
            // 
            this.Label111.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label111.ForeColor = System.Drawing.Color.Red;
            this.Label111.Location = new System.Drawing.Point(23, 4);
            this.Label111.Name = "Label111";
            this.Label111.Size = new System.Drawing.Size(340, 20);
            this.Label111.TabIndex = 120;
            this.Label111.Text = "入金データ入力作業 (入金処理)";
            this.Label111.UseCompatibleTextRendering = true;
            // 
            // Label110
            // 
            this.Label110.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label110.Location = new System.Drawing.Point(23, 28);
            this.Label110.Name = "Label110";
            this.Label110.Size = new System.Drawing.Size(340, 20);
            this.Label110.TabIndex = 121;
            this.Label110.Text = "必要な入金処理を行って下さい。";
            this.Label110.UseCompatibleTextRendering = true;
            // 
            // pnlJigoCmtSqKotiku
            // 
            this.pnlJigoCmtSqKotiku.Controls.Add(this.Label113);
            this.pnlJigoCmtSqKotiku.Controls.Add(this.Label112);
            this.pnlJigoCmtSqKotiku.Controls.Add(this.Label24);
            this.pnlJigoCmtSqKotiku.Location = new System.Drawing.Point(20, 25);
            this.pnlJigoCmtSqKotiku.Name = "pnlJigoCmtSqKotiku";
            this.pnlJigoCmtSqKotiku.Size = new System.Drawing.Size(375, 65);
            this.pnlJigoCmtSqKotiku.TabIndex = 136;
            // 
            // Label113
            // 
            this.Label113.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label113.ForeColor = System.Drawing.Color.Red;
            this.Label113.Location = new System.Drawing.Point(23, 4);
            this.Label113.Name = "Label113";
            this.Label113.Size = new System.Drawing.Size(330, 20);
            this.Label113.TabIndex = 117;
            this.Label113.Text = "請求データ一括構築 (請求データ作成)";
            this.Label113.UseCompatibleTextRendering = true;
            // 
            // Label112
            // 
            this.Label112.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label112.Location = new System.Drawing.Point(23, 28);
            this.Label112.Name = "Label112";
            this.Label112.Size = new System.Drawing.Size(340, 36);
            this.Label112.TabIndex = 118;
            this.Label112.Text = "コンバート後、請求データ一括構築を実行して下さい。データ構築を行わないと請求データが作成されません。";
            this.Label112.UseCompatibleTextRendering = true;
            // 
            // Label24
            // 
            this.Label24.ForeColor = System.Drawing.Color.Red;
            this.Label24.Location = new System.Drawing.Point(3, 4);
            this.Label24.Name = "Label24";
            this.Label24.Size = new System.Drawing.Size(20, 20);
            this.Label24.TabIndex = 133;
            this.Label24.Text = "・";
            this.Label24.UseCompatibleTextRendering = true;
            // 
            // grpJigoUserSagyo
            // 
            this.grpJigoUserSagyo.BackColor = System.Drawing.SystemColors.Menu;
            this.grpJigoUserSagyo.Controls.Add(this.pnlHJigoCmtSyudo);
            this.grpJigoUserSagyo.Controls.Add(this.lblJizenPageNum1);
            this.grpJigoUserSagyo.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpJigoUserSagyo.Location = new System.Drawing.Point(6, 15);
            this.grpJigoUserSagyo.Name = "grpJigoUserSagyo";
            this.grpJigoUserSagyo.Size = new System.Drawing.Size(405, 342);
            this.grpJigoUserSagyo.TabIndex = 137;
            this.grpJigoUserSagyo.TabStop = false;
            this.grpJigoUserSagyo.Text = "【データ移行可否によるユーザ作業】";
            // 
            // pnlHJigoCmtSyudo
            // 
            this.pnlHJigoCmtSyudo.Controls.Add(this.Label136);
            this.pnlHJigoCmtSyudo.Controls.Add(this.Label317);
            this.pnlHJigoCmtSyudo.Controls.Add(this.Label318);
            this.pnlHJigoCmtSyudo.Location = new System.Drawing.Point(6, 18);
            this.pnlHJigoCmtSyudo.Name = "pnlHJigoCmtSyudo";
            this.pnlHJigoCmtSyudo.Size = new System.Drawing.Size(375, 80);
            this.pnlHJigoCmtSyudo.TabIndex = 140;
            // 
            // Label136
            // 
            this.Label136.Location = new System.Drawing.Point(11, 4);
            this.Label136.Name = "Label136";
            this.Label136.Size = new System.Drawing.Size(20, 20);
            this.Label136.TabIndex = 137;
            this.Label136.Text = "・";
            this.Label136.UseCompatibleTextRendering = true;
            // 
            // Label317
            // 
            this.Label317.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label317.Location = new System.Drawing.Point(32, 4);
            this.Label317.Name = "Label317";
            this.Label317.Size = new System.Drawing.Size(340, 20);
            this.Label317.TabIndex = 130;
            this.Label317.Text = "コンバート後に手動設定が必要な項目について";
            this.Label317.UseCompatibleTextRendering = true;
            // 
            // Label318
            // 
            this.Label318.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label318.Location = new System.Drawing.Point(36, 27);
            this.Label318.Name = "Label318";
            this.Label318.Size = new System.Drawing.Size(340, 50);
            this.Label318.TabIndex = 131;
            this.Label318.Text = "クライアントやセキュリティ、初期設定などにつていてはユーザにて手動設定頂く必要があります。設定が必要な項目については、仕様書をご確認下さい。";
            this.Label318.UseCompatibleTextRendering = true;
            // 
            // lblJizenPageNum1
            // 
            this.lblJizenPageNum1.AutoSize = true;
            this.lblJizenPageNum1.Location = new System.Drawing.Point(794, 321);
            this.lblJizenPageNum1.Name = "lblJizenPageNum1";
            this.lblJizenPageNum1.Size = new System.Drawing.Size(35, 18);
            this.lblJizenPageNum1.TabIndex = 138;
            this.lblJizenPageNum1.Text = "1 / 2";
            // 
            // lblJigoDescription1
            // 
            this.lblJigoDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblJigoDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblJigoDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblJigoDescription1.Location = new System.Drawing.Point(30, 20);
            this.lblJigoDescription1.Name = "lblJigoDescription1";
            this.lblJigoDescription1.Size = new System.Drawing.Size(820, 80);
            this.lblJigoDescription1.TabIndex = 6;
            this.lblJigoDescription1.Text = "コンバート後の作業を行います。\r\n以下の内容を確認し、必要な作業を行って下さい。";
            this.lblJigoDescription1.UseCompatibleTextRendering = true;
            // 
            // tabPageHojyo
            // 
            this.tabPageHojyo.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHojyo.Location = new System.Drawing.Point(4, 27);
            this.tabPageHojyo.Name = "tabPageHojyo";
            this.tabPageHojyo.Size = new System.Drawing.Size(912, 549);
            this.tabPageHojyo.TabIndex = 15;
            this.tabPageHojyo.Text = "検証用";
            // 
            // tabPageHajimeni
            // 
            this.tabPageHajimeni.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHajimeni.Controls.Add(this.lblDatacvHajimeniDescription1);
            this.tabPageHajimeni.Controls.Add(this.lblHDatacvHajimeniLabel);
            this.tabPageHajimeni.Controls.Add(this.btnDoui);
            this.tabPageHajimeni.Controls.Add(this.GroupBox10);
            this.tabPageHajimeni.Controls.Add(this.PictureBox14);
            this.tabPageHajimeni.Location = new System.Drawing.Point(4, 27);
            this.tabPageHajimeni.Name = "tabPageHajimeni";
            this.tabPageHajimeni.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageHajimeni.Size = new System.Drawing.Size(912, 549);
            this.tabPageHajimeni.TabIndex = 6;
            this.tabPageHajimeni.Text = " はじめに";
            // 
            // lblDatacvHajimeniDescription1
            // 
            this.lblDatacvHajimeniDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblDatacvHajimeniDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F);
            this.lblDatacvHajimeniDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblDatacvHajimeniDescription1.Location = new System.Drawing.Point(35, 49);
            this.lblDatacvHajimeniDescription1.Name = "lblDatacvHajimeniDescription1";
            this.lblDatacvHajimeniDescription1.Size = new System.Drawing.Size(540, 35);
            this.lblDatacvHajimeniDescription1.TabIndex = 1;
            this.lblDatacvHajimeniDescription1.Text = "以下の留意事項を確認の上、同意ボタンを押して下さい。\r\n(同意後「次へ」ボタンを押して下さい。)";
            this.lblDatacvHajimeniDescription1.UseCompatibleTextRendering = true;
            // 
            // lblHDatacvHajimeniLabel
            // 
            this.lblHDatacvHajimeniLabel.BackColor = System.Drawing.SystemColors.Menu;
            this.lblHDatacvHajimeniLabel.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHDatacvHajimeniLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lblHDatacvHajimeniLabel.Location = new System.Drawing.Point(30, 20);
            this.lblHDatacvHajimeniLabel.Name = "lblHDatacvHajimeniLabel";
            this.lblHDatacvHajimeniLabel.Size = new System.Drawing.Size(587, 80);
            this.lblHDatacvHajimeniLabel.TabIndex = 139;
            this.lblHDatacvHajimeniLabel.Text = "「中間テーブル」から「賃貸革命10」へデータコンバートを行います。";
            this.lblHDatacvHajimeniLabel.UseCompatibleTextRendering = true;
            // 
            // btnDoui
            // 
            this.btnDoui.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDoui.ForeColor = System.Drawing.Color.Red;
            this.btnDoui.Location = new System.Drawing.Point(632, 53);
            this.btnDoui.Name = "btnDoui";
            this.btnDoui.Size = new System.Drawing.Size(251, 43);
            this.btnDoui.TabIndex = 0;
            this.btnDoui.Text = "留意事項に同意します";
            this.btnDoui.UseVisualStyleBackColor = true;
            this.btnDoui.Click += new System.EventHandler(this.btnDoui_Click);
            // 
            // GroupBox10
            // 
            this.GroupBox10.BackColor = System.Drawing.SystemColors.Menu;
            this.GroupBox10.Controls.Add(this.pnlDcFstCmtH99);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmtH01);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmtH02);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt11);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt06);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt05);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt07);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt10);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt02);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt08);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt04);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt03);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt01);
            this.GroupBox10.Controls.Add(this.pnlDcFstCmt09);
            this.GroupBox10.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox10.Location = new System.Drawing.Point(24, 115);
            this.GroupBox10.Name = "GroupBox10";
            this.GroupBox10.Size = new System.Drawing.Size(865, 423);
            this.GroupBox10.TabIndex = 2;
            this.GroupBox10.TabStop = false;
            this.GroupBox10.Text = "【留意事項】";
            // 
            // pnlDcFstCmtH99
            // 
            this.pnlDcFstCmtH99.Controls.Add(this.Label301);
            this.pnlDcFstCmtH99.Controls.Add(this.Label302);
            this.pnlDcFstCmtH99.Location = new System.Drawing.Point(455, 300);
            this.pnlDcFstCmtH99.Name = "pnlDcFstCmtH99";
            this.pnlDcFstCmtH99.Size = new System.Drawing.Size(395, 115);
            this.pnlDcFstCmtH99.TabIndex = 26;
            // 
            // Label301
            // 
            this.Label301.Location = new System.Drawing.Point(3, 3);
            this.Label301.Name = "Label301";
            this.Label301.Size = new System.Drawing.Size(200, 16);
            this.Label301.TabIndex = 14;
            this.Label301.Text = "<コンバート対象>";
            this.Label301.UseCompatibleTextRendering = true;
            // 
            // Label302
            // 
            this.Label302.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label302.Location = new System.Drawing.Point(23, 4);
            this.Label302.Name = "Label302";
            this.Label302.Size = new System.Drawing.Size(370, 105);
            this.Label302.TabIndex = 15;
            this.Label302.Text = "\r\n物件情報、部屋情報\r\n契約情報(最新契約のみ対象。契約・更新歴、解約は対象外)\r\n家主情報、契約者情報、業者情報、口座情報、自社情報\r\n※請求・入金・送金、画" +
    "像は対応していません。\r\n※コンバート後に請求構築などの作業を行う必要があります。";
            this.Label302.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmtH01
            // 
            this.pnlDcFstCmtH01.Controls.Add(this.Label158);
            this.pnlDcFstCmtH01.Controls.Add(this.Label159);
            this.pnlDcFstCmtH01.Location = new System.Drawing.Point(455, 175);
            this.pnlDcFstCmtH01.Name = "pnlDcFstCmtH01";
            this.pnlDcFstCmtH01.Size = new System.Drawing.Size(395, 28);
            this.pnlDcFstCmtH01.TabIndex = 24;
            // 
            // Label158
            // 
            this.Label158.Location = new System.Drawing.Point(3, 3);
            this.Label158.Name = "Label158";
            this.Label158.Size = new System.Drawing.Size(14, 16);
            this.Label158.TabIndex = 4;
            this.Label158.Text = "・";
            this.Label158.UseCompatibleTextRendering = true;
            // 
            // Label159
            // 
            this.Label159.Location = new System.Drawing.Point(23, 4);
            this.Label159.Name = "Label159";
            this.Label159.Size = new System.Drawing.Size(370, 18);
            this.Label159.TabIndex = 5;
            this.Label159.Text = "中間ファイルへ事前にデータを登録する必要があります。";
            this.Label159.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmtH02
            // 
            this.pnlDcFstCmtH02.Controls.Add(this.Label27);
            this.pnlDcFstCmtH02.Controls.Add(this.Label28);
            this.pnlDcFstCmtH02.Location = new System.Drawing.Point(455, 202);
            this.pnlDcFstCmtH02.Name = "pnlDcFstCmtH02";
            this.pnlDcFstCmtH02.Size = new System.Drawing.Size(395, 63);
            this.pnlDcFstCmtH02.TabIndex = 23;
            // 
            // Label27
            // 
            this.Label27.Location = new System.Drawing.Point(3, 3);
            this.Label27.Name = "Label27";
            this.Label27.Size = new System.Drawing.Size(14, 16);
            this.Label27.TabIndex = 4;
            this.Label27.Text = "・";
            this.Label27.UseCompatibleTextRendering = true;
            // 
            // Label28
            // 
            this.Label28.Location = new System.Drawing.Point(23, 4);
            this.Label28.Name = "Label28";
            this.Label28.Size = new System.Drawing.Size(370, 55);
            this.Label28.TabIndex = 5;
            this.Label28.Text = "弊社にて用意した中間ファイルを必ずご使用下さい。中間ファイルの加工・それ以外(お客様で用意されたファイルなど)については対応していません。";
            this.Label28.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt11
            // 
            this.pnlDcFstCmt11.Controls.Add(this.Label21);
            this.pnlDcFstCmt11.Controls.Add(this.Label25);
            this.pnlDcFstCmt11.Location = new System.Drawing.Point(455, 131);
            this.pnlDcFstCmt11.Name = "pnlDcFstCmt11";
            this.pnlDcFstCmt11.Size = new System.Drawing.Size(395, 45);
            this.pnlDcFstCmt11.TabIndex = 22;
            // 
            // Label21
            // 
            this.Label21.Location = new System.Drawing.Point(3, 3);
            this.Label21.Name = "Label21";
            this.Label21.Size = new System.Drawing.Size(14, 16);
            this.Label21.TabIndex = 4;
            this.Label21.Text = "・";
            this.Label21.UseCompatibleTextRendering = true;
            // 
            // Label25
            // 
            this.Label25.Location = new System.Drawing.Point(23, 4);
            this.Label25.Name = "Label25";
            this.Label25.Size = new System.Drawing.Size(370, 38);
            this.Label25.TabIndex = 5;
            this.Label25.Text = "運用開始(本稼働)後のコンバートは、整合性が取れなくなる場合がある為、原則行わないで下さい。";
            this.Label25.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt06
            // 
            this.pnlDcFstCmt06.Controls.Add(this.Label364);
            this.pnlDcFstCmt06.Controls.Add(this.Label369);
            this.pnlDcFstCmt06.Location = new System.Drawing.Point(15, 264);
            this.pnlDcFstCmt06.Name = "pnlDcFstCmt06";
            this.pnlDcFstCmt06.Size = new System.Drawing.Size(395, 28);
            this.pnlDcFstCmt06.TabIndex = 20;
            // 
            // Label364
            // 
            this.Label364.Location = new System.Drawing.Point(3, 3);
            this.Label364.Name = "Label364";
            this.Label364.Size = new System.Drawing.Size(14, 16);
            this.Label364.TabIndex = 6;
            this.Label364.Text = "・";
            this.Label364.UseCompatibleTextRendering = true;
            // 
            // Label369
            // 
            this.Label369.Location = new System.Drawing.Point(23, 4);
            this.Label369.Name = "Label369";
            this.Label369.Size = new System.Drawing.Size(370, 18);
            this.Label369.TabIndex = 7;
            this.Label369.Text = "暦(こよみ)に存在しない不適切な日付はコンバートできません。";
            this.Label369.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt05
            // 
            this.pnlDcFstCmt05.Controls.Add(this.Label362);
            this.pnlDcFstCmt05.Controls.Add(this.Label363);
            this.pnlDcFstCmt05.Location = new System.Drawing.Point(15, 202);
            this.pnlDcFstCmt05.Name = "pnlDcFstCmt05";
            this.pnlDcFstCmt05.Size = new System.Drawing.Size(395, 63);
            this.pnlDcFstCmt05.TabIndex = 15;
            // 
            // Label362
            // 
            this.Label362.Location = new System.Drawing.Point(3, 3);
            this.Label362.Name = "Label362";
            this.Label362.Size = new System.Drawing.Size(14, 16);
            this.Label362.TabIndex = 10;
            this.Label362.Text = "・";
            this.Label362.UseCompatibleTextRendering = true;
            // 
            // Label363
            // 
            this.Label363.Location = new System.Drawing.Point(23, 4);
            this.Label363.Name = "Label363";
            this.Label363.Size = new System.Drawing.Size(370, 55);
            this.Label363.TabIndex = 11;
            this.Label363.Text = "ある情報を一意に識別するコードが重複しているデータの場合、最初に読み込まれるデータのみコンバート致します。(例：複数の同一物件が存在する場合、1 物件のみコンバー" +
    "トします)";
            this.Label363.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt07
            // 
            this.pnlDcFstCmt07.Controls.Add(this.Label122);
            this.pnlDcFstCmt07.Controls.Add(this.Label104);
            this.pnlDcFstCmt07.Location = new System.Drawing.Point(15, 291);
            this.pnlDcFstCmt07.Name = "pnlDcFstCmt07";
            this.pnlDcFstCmt07.Size = new System.Drawing.Size(395, 45);
            this.pnlDcFstCmt07.TabIndex = 6;
            // 
            // Label122
            // 
            this.Label122.Location = new System.Drawing.Point(3, 3);
            this.Label122.Name = "Label122";
            this.Label122.Size = new System.Drawing.Size(14, 16);
            this.Label122.TabIndex = 2;
            this.Label122.Text = "・";
            this.Label122.UseCompatibleTextRendering = true;
            // 
            // Label104
            // 
            this.Label104.Location = new System.Drawing.Point(23, 4);
            this.Label104.Name = "Label104";
            this.Label104.Size = new System.Drawing.Size(370, 38);
            this.Label104.TabIndex = 3;
            this.Label104.Text = "機種依存文字・旧字体・制御文字(改行・タブ文字など)文字は正しく変換できない場合があります。";
            this.Label104.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt10
            // 
            this.pnlDcFstCmt10.Controls.Add(this.Label123);
            this.pnlDcFstCmt10.Controls.Add(this.Label105);
            this.pnlDcFstCmt10.Location = new System.Drawing.Point(455, 87);
            this.pnlDcFstCmt10.Name = "pnlDcFstCmt10";
            this.pnlDcFstCmt10.Size = new System.Drawing.Size(395, 45);
            this.pnlDcFstCmt10.TabIndex = 8;
            // 
            // Label123
            // 
            this.Label123.Location = new System.Drawing.Point(3, 3);
            this.Label123.Name = "Label123";
            this.Label123.Size = new System.Drawing.Size(14, 16);
            this.Label123.TabIndex = 4;
            this.Label123.Text = "・";
            this.Label123.UseCompatibleTextRendering = true;
            // 
            // Label105
            // 
            this.Label105.Location = new System.Drawing.Point(23, 4);
            this.Label105.Name = "Label105";
            this.Label105.Size = new System.Drawing.Size(370, 38);
            this.Label105.TabIndex = 5;
            this.Label105.Text = "本プログラム処理中は、他の作業(賃貸革命やEXCELの使用)を行わないで下さい。";
            this.Label105.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt02
            // 
            this.pnlDcFstCmt02.Controls.Add(this.Label127);
            this.pnlDcFstCmt02.Controls.Add(this.Label133);
            this.pnlDcFstCmt02.Location = new System.Drawing.Point(15, 69);
            this.pnlDcFstCmt02.Name = "pnlDcFstCmt02";
            this.pnlDcFstCmt02.Size = new System.Drawing.Size(395, 63);
            this.pnlDcFstCmt02.TabIndex = 10;
            // 
            // Label127
            // 
            this.Label127.Location = new System.Drawing.Point(3, 3);
            this.Label127.Name = "Label127";
            this.Label127.Size = new System.Drawing.Size(14, 16);
            this.Label127.TabIndex = 6;
            this.Label127.Text = "・";
            this.Label127.UseCompatibleTextRendering = true;
            // 
            // Label133
            // 
            this.Label133.Location = new System.Drawing.Point(23, 4);
            this.Label133.Name = "Label133";
            this.Label133.Size = new System.Drawing.Size(370, 55);
            this.Label133.TabIndex = 7;
            this.Label133.Text = "複数のデータベースから1 つのデータベースへ、または1つのデータベースを複数のデータベースに分けるようなコンバートは対応しておりません。";
            this.Label133.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt08
            // 
            this.pnlDcFstCmt08.Controls.Add(this.Label126);
            this.pnlDcFstCmt08.Controls.Add(this.Label132);
            this.pnlDcFstCmt08.Location = new System.Drawing.Point(15, 335);
            this.pnlDcFstCmt08.Name = "pnlDcFstCmt08";
            this.pnlDcFstCmt08.Size = new System.Drawing.Size(395, 45);
            this.pnlDcFstCmt08.TabIndex = 12;
            // 
            // Label126
            // 
            this.Label126.Location = new System.Drawing.Point(3, 3);
            this.Label126.Name = "Label126";
            this.Label126.Size = new System.Drawing.Size(14, 16);
            this.Label126.TabIndex = 8;
            this.Label126.Text = "・";
            this.Label126.UseCompatibleTextRendering = true;
            // 
            // Label132
            // 
            this.Label132.Location = new System.Drawing.Point(23, 4);
            this.Label132.Name = "Label132";
            this.Label132.Size = new System.Drawing.Size(370, 33);
            this.Label132.TabIndex = 9;
            this.Label132.Text = "親情報がない子情報のコンバートはできません。(例：部屋情報がコンバートされていない契約情報は移行できません)";
            this.Label132.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt04
            // 
            this.pnlDcFstCmt04.Controls.Add(this.Label119);
            this.pnlDcFstCmt04.Controls.Add(this.Label116);
            this.pnlDcFstCmt04.Location = new System.Drawing.Point(15, 175);
            this.pnlDcFstCmt04.Name = "pnlDcFstCmt04";
            this.pnlDcFstCmt04.Size = new System.Drawing.Size(395, 28);
            this.pnlDcFstCmt04.TabIndex = 14;
            // 
            // Label119
            // 
            this.Label119.Location = new System.Drawing.Point(3, 3);
            this.Label119.Name = "Label119";
            this.Label119.Size = new System.Drawing.Size(14, 16);
            this.Label119.TabIndex = 10;
            this.Label119.Text = "・";
            this.Label119.UseCompatibleTextRendering = true;
            // 
            // Label116
            // 
            this.Label116.Location = new System.Drawing.Point(23, 4);
            this.Label116.Name = "Label116";
            this.Label116.Size = new System.Drawing.Size(370, 18);
            this.Label116.TabIndex = 11;
            this.Label116.Text = "必須項目が不足しているデータはコンバートできません。";
            this.Label116.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt03
            // 
            this.pnlDcFstCmt03.Controls.Add(this.Label97);
            this.pnlDcFstCmt03.Controls.Add(this.Label101);
            this.pnlDcFstCmt03.Location = new System.Drawing.Point(15, 131);
            this.pnlDcFstCmt03.Name = "pnlDcFstCmt03";
            this.pnlDcFstCmt03.Size = new System.Drawing.Size(395, 45);
            this.pnlDcFstCmt03.TabIndex = 19;
            // 
            // Label97
            // 
            this.Label97.Location = new System.Drawing.Point(23, 4);
            this.Label97.Name = "Label97";
            this.Label97.Size = new System.Drawing.Size(370, 38);
            this.Label97.TabIndex = 13;
            this.Label97.Text = "移行仕様書に記載していない、または従っていない登録データはコンバートできません。";
            this.Label97.UseCompatibleTextRendering = true;
            // 
            // Label101
            // 
            this.Label101.Location = new System.Drawing.Point(3, 3);
            this.Label101.Name = "Label101";
            this.Label101.Size = new System.Drawing.Size(14, 16);
            this.Label101.TabIndex = 12;
            this.Label101.Text = "・";
            this.Label101.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt01
            // 
            this.pnlDcFstCmt01.Controls.Add(this.Label129);
            this.pnlDcFstCmt01.Controls.Add(this.Label135);
            this.pnlDcFstCmt01.Location = new System.Drawing.Point(15, 25);
            this.pnlDcFstCmt01.Name = "pnlDcFstCmt01";
            this.pnlDcFstCmt01.Size = new System.Drawing.Size(395, 45);
            this.pnlDcFstCmt01.TabIndex = 18;
            // 
            // Label129
            // 
            this.Label129.Location = new System.Drawing.Point(3, 3);
            this.Label129.Name = "Label129";
            this.Label129.Size = new System.Drawing.Size(14, 16);
            this.Label129.TabIndex = 0;
            this.Label129.Text = "・";
            this.Label129.UseCompatibleTextRendering = true;
            // 
            // Label135
            // 
            this.Label135.Location = new System.Drawing.Point(23, 4);
            this.Label135.Name = "Label135";
            this.Label135.Size = new System.Drawing.Size(370, 38);
            this.Label135.TabIndex = 1;
            this.Label135.Text = "本プログラムは、賃貸革命10 がインストールされている環境のみで動作します。";
            this.Label135.UseCompatibleTextRendering = true;
            // 
            // pnlDcFstCmt09
            // 
            this.pnlDcFstCmt09.Controls.Add(this.Label326);
            this.pnlDcFstCmt09.Controls.Add(this.Label325);
            this.pnlDcFstCmt09.Location = new System.Drawing.Point(455, 25);
            this.pnlDcFstCmt09.Name = "pnlDcFstCmt09";
            this.pnlDcFstCmt09.Size = new System.Drawing.Size(395, 63);
            this.pnlDcFstCmt09.TabIndex = 16;
            // 
            // Label326
            // 
            this.Label326.Location = new System.Drawing.Point(3, 3);
            this.Label326.Name = "Label326";
            this.Label326.Size = new System.Drawing.Size(14, 16);
            this.Label326.TabIndex = 16;
            this.Label326.Text = "・";
            this.Label326.UseCompatibleTextRendering = true;
            // 
            // Label325
            // 
            this.Label325.Location = new System.Drawing.Point(23, 4);
            this.Label325.Name = "Label325";
            this.Label325.Size = new System.Drawing.Size(370, 55);
            this.Label325.TabIndex = 17;
            this.Label325.Text = "金融機関・支店・沿線・駅・住所などのマスタ値を参照するデータにおいて、賃貸革命10 へ登録されている設定値以外はコンバートできません。";
            this.Label325.UseCompatibleTextRendering = true;
            // 
            // PictureBox14
            // 
            this.PictureBox14.Image = ((System.Drawing.Image)(resources.GetObject("PictureBox14.Image")));
            this.PictureBox14.Location = new System.Drawing.Point(562, 342);
            this.PictureBox14.Name = "PictureBox14";
            this.PictureBox14.Size = new System.Drawing.Size(254, 185);
            this.PictureBox14.TabIndex = 138;
            this.PictureBox14.TabStop = false;
            // 
            // tabPageSelect
            // 
            this.tabPageSelect.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageSelect.Controls.Add(this.grpExistMidToBaseMid);
            this.tabPageSelect.Controls.Add(this.btnAllChk);
            this.tabPageSelect.Controls.Add(this.lblDatacvSelectCaution);
            this.tabPageSelect.Controls.Add(this.lblLine2);
            this.tabPageSelect.Controls.Add(this.lblHidden2);
            this.tabPageSelect.Controls.Add(this.grpMiddleFile);
            this.tabPageSelect.Controls.Add(this.tabCtrlCVItem);
            this.tabPageSelect.Controls.Add(this.pnlRekiClear);
            this.tabPageSelect.Controls.Add(this.lblDatacvSelectDescription1);
            this.tabPageSelect.Location = new System.Drawing.Point(4, 27);
            this.tabPageSelect.Name = "tabPageSelect";
            this.tabPageSelect.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSelect.Size = new System.Drawing.Size(912, 549);
            this.tabPageSelect.TabIndex = 2;
            this.tabPageSelect.Text = " 対象項目選択";
            // 
            // grpExistMidToBaseMid
            // 
            this.grpExistMidToBaseMid.Font = new System.Drawing.Font("メイリオ", 8.25F);
            this.grpExistMidToBaseMid.Location = new System.Drawing.Point(544, -14);
            this.grpExistMidToBaseMid.Name = "grpExistMidToBaseMid";
            this.grpExistMidToBaseMid.Size = new System.Drawing.Size(368, 48);
            this.grpExistMidToBaseMid.TabIndex = 140;
            this.grpExistMidToBaseMid.TabStop = false;
            // 
            // btnAllChk
            // 
            this.btnAllChk.Image = ((System.Drawing.Image)(resources.GetObject("btnAllChk.Image")));
            this.btnAllChk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAllChk.Location = new System.Drawing.Point(689, 87);
            this.btnAllChk.Name = "btnAllChk";
            this.btnAllChk.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnAllChk.Size = new System.Drawing.Size(200, 30);
            this.btnAllChk.TabIndex = 3;
            this.btnAllChk.Text = " 画面毎全チェックON/OFF";
            this.btnAllChk.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnAllChk.UseVisualStyleBackColor = true;
            this.btnAllChk.Click += new System.EventHandler(this.btnAllChk_Click);
            // 
            // lblDatacvSelectCaution
            // 
            this.lblDatacvSelectCaution.BackColor = System.Drawing.SystemColors.Menu;
            this.lblDatacvSelectCaution.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDatacvSelectCaution.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.lblDatacvSelectCaution.Location = new System.Drawing.Point(30, 59);
            this.lblDatacvSelectCaution.Name = "lblDatacvSelectCaution";
            this.lblDatacvSelectCaution.Size = new System.Drawing.Size(490, 58);
            this.lblDatacvSelectCaution.TabIndex = 1;
            this.lblDatacvSelectCaution.Text = "※選択項目の組み合わせ(項目の親子関係)によっては、正しく移行されない場合\r\n　があります。項目の親子関係を考慮した上で選択下さい。\r\n(例：物件情報が未移行の場" +
    "合、部屋情報や契約情報の移行はできません)";
            this.lblDatacvSelectCaution.UseCompatibleTextRendering = true;
            // 
            // lblLine2
            // 
            this.lblLine2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine2.Location = new System.Drawing.Point(31, 216);
            this.lblLine2.Name = "lblLine2";
            this.lblLine2.Size = new System.Drawing.Size(853, 2);
            this.lblLine2.TabIndex = 6;
            // 
            // lblHidden2
            // 
            this.lblHidden2.Location = new System.Drawing.Point(6, 190);
            this.lblHidden2.Name = "lblHidden2";
            this.lblHidden2.Size = new System.Drawing.Size(18, 28);
            this.lblHidden2.TabIndex = 6;
            this.lblHidden2.Text = "　";
            // 
            // grpMiddleFile
            // 
            this.grpMiddleFile.Controls.Add(this.Label349);
            this.grpMiddleFile.Controls.Add(this.Label328);
            this.grpMiddleFile.Controls.Add(this.Label124);
            this.grpMiddleFile.Controls.Add(this.btnMidFileCheck);
            this.grpMiddleFile.Controls.Add(this.btnMidDirLogSeach);
            this.grpMiddleFile.Controls.Add(this.txtMidDirLogPath);
            this.grpMiddleFile.Controls.Add(this.lblMiddleFileLog);
            this.grpMiddleFile.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpMiddleFile.Location = new System.Drawing.Point(30, 114);
            this.grpMiddleFile.Name = "grpMiddleFile";
            this.grpMiddleFile.Size = new System.Drawing.Size(854, 73);
            this.grpMiddleFile.TabIndex = 4;
            this.grpMiddleFile.TabStop = false;
            // 
            // Label349
            // 
            this.Label349.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label349.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Label349.Location = new System.Drawing.Point(158, 16);
            this.Label349.Name = "Label349";
            this.Label349.Size = new System.Drawing.Size(46, 20);
            this.Label349.TabIndex = 143;
            this.Label349.Text = "※必須";
            this.Label349.UseCompatibleTextRendering = true;
            // 
            // Label328
            // 
            this.Label328.BackColor = System.Drawing.SystemColors.Menu;
            this.Label328.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label328.ForeColor = System.Drawing.Color.Black;
            this.Label328.Location = new System.Drawing.Point(29, 35);
            this.Label328.Name = "Label328";
            this.Label328.Size = new System.Drawing.Size(315, 32);
            this.Label328.TabIndex = 138;
            this.Label328.Text = "選択された項目を対象に、Preテーブルへ登録したデータの\r\n型やサイズなどの正当性をチェックします。";
            this.Label328.UseCompatibleTextRendering = true;
            // 
            // Label124
            // 
            this.Label124.BackColor = System.Drawing.SystemColors.Menu;
            this.Label124.Font = new System.Drawing.Font("メイリオ", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label124.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Label124.Location = new System.Drawing.Point(15, 15);
            this.Label124.Name = "Label124";
            this.Label124.Size = new System.Drawing.Size(173, 19);
            this.Label124.TabIndex = 7;
            this.Label124.Text = "Preテーブルチェック";
            this.Label124.UseCompatibleTextRendering = true;
            // 
            // btnMidFileCheck
            // 
            this.btnMidFileCheck.BackColor = System.Drawing.SystemColors.Menu;
            this.btnMidFileCheck.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMidFileCheck.Image = ((System.Drawing.Image)(resources.GetObject("btnMidFileCheck.Image")));
            this.btnMidFileCheck.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnMidFileCheck.Location = new System.Drawing.Point(748, 43);
            this.btnMidFileCheck.Name = "btnMidFileCheck";
            this.btnMidFileCheck.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnMidFileCheck.Size = new System.Drawing.Size(100, 25);
            this.btnMidFileCheck.TabIndex = 6;
            this.btnMidFileCheck.Text = " チェック";
            this.btnMidFileCheck.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnMidFileCheck.UseVisualStyleBackColor = true;
            this.btnMidFileCheck.Click += new System.EventHandler(this.btnMidFileCheck_Click);
            // 
            // btnMidDirLogSeach
            // 
            this.btnMidDirLogSeach.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMidDirLogSeach.Location = new System.Drawing.Point(712, 43);
            this.btnMidDirLogSeach.Name = "btnMidDirLogSeach";
            this.btnMidDirLogSeach.Size = new System.Drawing.Size(30, 25);
            this.btnMidDirLogSeach.TabIndex = 5;
            this.btnMidDirLogSeach.Text = "...";
            this.btnMidDirLogSeach.UseVisualStyleBackColor = true;
            this.btnMidDirLogSeach.Click += new System.EventHandler(this.btnDirSeach_Click);
            // 
            // txtMidDirLogPath
            // 
            this.txtMidDirLogPath.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMidDirLogPath.Location = new System.Drawing.Point(520, 44);
            this.txtMidDirLogPath.Name = "txtMidDirLogPath";
            this.txtMidDirLogPath.Size = new System.Drawing.Size(186, 24);
            this.txtMidDirLogPath.TabIndex = 4;
            // 
            // lblMiddleFileLog
            // 
            this.lblMiddleFileLog.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblMiddleFileLog.Location = new System.Drawing.Point(397, 47);
            this.lblMiddleFileLog.Name = "lblMiddleFileLog";
            this.lblMiddleFileLog.Size = new System.Drawing.Size(117, 20);
            this.lblMiddleFileLog.TabIndex = 3;
            this.lblMiddleFileLog.Text = "結果ファイル格納先";
            this.lblMiddleFileLog.UseCompatibleTextRendering = true;
            // 
            // tabCtrlCVItem
            // 
            this.tabCtrlCVItem.Controls.Add(this.tabPageKizon110);
            this.tabCtrlCVItem.Controls.Add(this.tabPageKizon120);
            this.tabCtrlCVItem.Controls.Add(this.tabPageKizon130);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase110);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase120);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase130);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase140);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase150);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase160);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase170);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase180);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase190);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase200);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase210);
            this.tabCtrlCVItem.Controls.Add(this.tabPageBase900);
            this.tabCtrlCVItem.Controls.Add(this.tabPageHanyo110);
            this.tabCtrlCVItem.Controls.Add(this.tabPageHanyo120);
            this.tabCtrlCVItem.Controls.Add(this.tabPageHanyo130);
            this.tabCtrlCVItem.Controls.Add(this.tabPageHanyo140);
            this.tabCtrlCVItem.Controls.Add(this.tabPageHanyo150);
            this.tabCtrlCVItem.Controls.Add(this.TabPage1);
            this.tabCtrlCVItem.Location = new System.Drawing.Point(30, 190);
            this.tabCtrlCVItem.Name = "tabCtrlCVItem";
            this.tabCtrlCVItem.SelectedIndex = 0;
            this.tabCtrlCVItem.Size = new System.Drawing.Size(855, 358);
            this.tabCtrlCVItem.TabIndex = 5;
            this.tabCtrlCVItem.TabStop = false;
            // 
            // tabPageKizon110
            // 
            this.tabPageKizon110.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageKizon110.Controls.Add(this.grpKizon1);
            this.tabPageKizon110.Controls.Add(this.lblSelectPageCnt1);
            this.tabPageKizon110.Location = new System.Drawing.Point(4, 27);
            this.tabPageKizon110.Name = "tabPageKizon110";
            this.tabPageKizon110.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageKizon110.Size = new System.Drawing.Size(847, 327);
            this.tabPageKizon110.TabIndex = 9;
            this.tabPageKizon110.Text = " 項目選択1";
            // 
            // grpKizon1
            // 
            this.grpKizon1.Controls.Add(this.pnlKiMstHendo);
            this.grpKizon1.Controls.Add(this.pnlKiMstTokuyaku);
            this.grpKizon1.Controls.Add(this.pnlKiMstKasyoClaimrui);
            this.grpKizon1.Controls.Add(this.pnlKiMstTitle);
            this.grpKizon1.Controls.Add(this.pnlKiMstArea);
            this.grpKizon1.Controls.Add(this.pnlKiMstSchool);
            this.grpKizon1.Controls.Add(this.pnlKiMstHokenrui);
            this.grpKizon1.Controls.Add(this.pnlKiMstBus);
            this.grpKizon1.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpKizon1.Location = new System.Drawing.Point(6, 5);
            this.grpKizon1.Name = "grpKizon1";
            this.grpKizon1.Size = new System.Drawing.Size(835, 298);
            this.grpKizon1.TabIndex = 0;
            this.grpKizon1.TabStop = false;
            this.grpKizon1.Text = "【各マスタ情報】";
            // 
            // pnlKiMstHendo
            // 
            this.pnlKiMstHendo.Controls.Add(this.chkKiMstHendo);
            this.pnlKiMstHendo.Controls.Add(this.lblKiMstHendoCnt);
            this.pnlKiMstHendo.Controls.Add(this.Label230);
            this.pnlKiMstHendo.Controls.Add(this.Label82);
            this.pnlKiMstHendo.Controls.Add(this.Label81);
            this.pnlKiMstHendo.Location = new System.Drawing.Point(557, 20);
            this.pnlKiMstHendo.Name = "pnlKiMstHendo";
            this.pnlKiMstHendo.Size = new System.Drawing.Size(265, 80);
            this.pnlKiMstHendo.TabIndex = 50;
            // 
            // chkKiMstHendo
            // 
            this.chkKiMstHendo.AutoSize = true;
            this.chkKiMstHendo.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiMstHendo.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkKiMstHendo.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstHendo.Name = "chkKiMstHendo";
            this.chkKiMstHendo.Size = new System.Drawing.Size(134, 27);
            this.chkKiMstHendo.TabIndex = 30;
            this.chkKiMstHendo.Text = "変動費設定内容";
            this.chkKiMstHendo.UseVisualStyleBackColor = true;
            this.chkKiMstHendo.CheckedChanged += new System.EventHandler(this.chkKiMstHendo_CheckedChanged);
            // 
            // lblKiMstHendoCnt
            // 
            this.lblKiMstHendoCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstHendoCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiMstHendoCnt.Name = "lblKiMstHendoCnt";
            this.lblKiMstHendoCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstHendoCnt.TabIndex = 33;
            this.lblKiMstHendoCnt.Text = "9,999,999";
            this.lblKiMstHendoCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label230
            // 
            this.Label230.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label230.Location = new System.Drawing.Point(19, 54);
            this.Label230.Name = "Label230";
            this.Label230.Size = new System.Drawing.Size(92, 16);
            this.Label230.TabIndex = 32;
            this.Label230.Text = "( 参考件数 = ";
            this.Label230.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label82
            // 
            this.Label82.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label82.Location = new System.Drawing.Point(195, 54);
            this.Label82.Name = "Label82";
            this.Label82.Size = new System.Drawing.Size(22, 16);
            this.Label82.TabIndex = 34;
            this.Label82.Text = " )";
            this.Label82.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label81
            // 
            this.Label81.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label81.Location = new System.Drawing.Point(19, 33);
            this.Label81.Name = "Label81";
            this.Label81.Size = new System.Drawing.Size(198, 16);
            this.Label81.TabIndex = 31;
            this.Label81.Text = "口径や料金表などの設定";
            this.Label81.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiMstTokuyaku
            // 
            this.pnlKiMstTokuyaku.Controls.Add(this.chkKiMstTokuyaku);
            this.pnlKiMstTokuyaku.Controls.Add(this.lblKiMstTokuyakuCnt);
            this.pnlKiMstTokuyaku.Controls.Add(this.Label223);
            this.pnlKiMstTokuyaku.Controls.Add(this.Label222);
            this.pnlKiMstTokuyaku.Controls.Add(this.Label78);
            this.pnlKiMstTokuyaku.Location = new System.Drawing.Point(286, 112);
            this.pnlKiMstTokuyaku.Name = "pnlKiMstTokuyaku";
            this.pnlKiMstTokuyaku.Size = new System.Drawing.Size(265, 80);
            this.pnlKiMstTokuyaku.TabIndex = 49;
            // 
            // chkKiMstTokuyaku
            // 
            this.chkKiMstTokuyaku.AutoSize = true;
            this.chkKiMstTokuyaku.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiMstTokuyaku.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstTokuyaku.Name = "chkKiMstTokuyaku";
            this.chkKiMstTokuyaku.Size = new System.Drawing.Size(104, 27);
            this.chkKiMstTokuyaku.TabIndex = 20;
            this.chkKiMstTokuyaku.Text = "特約マスタ";
            this.chkKiMstTokuyaku.UseVisualStyleBackColor = true;
            this.chkKiMstTokuyaku.CheckedChanged += new System.EventHandler(this.chkKiMstTokuyaku_CheckedChanged);
            // 
            // lblKiMstTokuyakuCnt
            // 
            this.lblKiMstTokuyakuCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstTokuyakuCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiMstTokuyakuCnt.Name = "lblKiMstTokuyakuCnt";
            this.lblKiMstTokuyakuCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstTokuyakuCnt.TabIndex = 23;
            this.lblKiMstTokuyakuCnt.Text = "9,999,999";
            this.lblKiMstTokuyakuCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label223
            // 
            this.Label223.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label223.Location = new System.Drawing.Point(19, 54);
            this.Label223.Name = "Label223";
            this.Label223.Size = new System.Drawing.Size(92, 16);
            this.Label223.TabIndex = 22;
            this.Label223.Text = "( 参考件数 = ";
            this.Label223.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label222
            // 
            this.Label222.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label222.Location = new System.Drawing.Point(195, 54);
            this.Label222.Name = "Label222";
            this.Label222.Size = new System.Drawing.Size(22, 16);
            this.Label222.TabIndex = 24;
            this.Label222.Text = " )";
            this.Label222.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label78
            // 
            this.Label78.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label78.Location = new System.Drawing.Point(19, 33);
            this.Label78.Name = "Label78";
            this.Label78.Size = new System.Drawing.Size(198, 16);
            this.Label78.TabIndex = 21;
            this.Label78.Text = "原状回復特約, 修繕特約, 特約事項";
            this.Label78.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiMstKasyoClaimrui
            // 
            this.pnlKiMstKasyoClaimrui.Controls.Add(this.chkKiMstKasyoClaimrui);
            this.pnlKiMstKasyoClaimrui.Controls.Add(this.lblKiMstKasyoClaimruiCnt);
            this.pnlKiMstKasyoClaimrui.Controls.Add(this.Label228);
            this.pnlKiMstKasyoClaimrui.Controls.Add(this.Label227);
            this.pnlKiMstKasyoClaimrui.Controls.Add(this.Label226);
            this.pnlKiMstKasyoClaimrui.Location = new System.Drawing.Point(286, 204);
            this.pnlKiMstKasyoClaimrui.Name = "pnlKiMstKasyoClaimrui";
            this.pnlKiMstKasyoClaimrui.Size = new System.Drawing.Size(265, 80);
            this.pnlKiMstKasyoClaimrui.TabIndex = 49;
            // 
            // chkKiMstKasyoClaimrui
            // 
            this.chkKiMstKasyoClaimrui.AutoSize = true;
            this.chkKiMstKasyoClaimrui.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiMstKasyoClaimrui.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstKasyoClaimrui.Name = "chkKiMstKasyoClaimrui";
            this.chkKiMstKasyoClaimrui.Size = new System.Drawing.Size(179, 27);
            this.chkKiMstKasyoClaimrui.TabIndex = 25;
            this.chkKiMstKasyoClaimrui.Text = "クレーム分類設定内容";
            this.chkKiMstKasyoClaimrui.UseVisualStyleBackColor = true;
            this.chkKiMstKasyoClaimrui.CheckedChanged += new System.EventHandler(this.chkKiMstKasyoClaimrui_CheckedChanged);
            // 
            // lblKiMstKasyoClaimruiCnt
            // 
            this.lblKiMstKasyoClaimruiCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstKasyoClaimruiCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiMstKasyoClaimruiCnt.Name = "lblKiMstKasyoClaimruiCnt";
            this.lblKiMstKasyoClaimruiCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstKasyoClaimruiCnt.TabIndex = 28;
            this.lblKiMstKasyoClaimruiCnt.Text = "9,999,999";
            this.lblKiMstKasyoClaimruiCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label228
            // 
            this.Label228.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label228.Location = new System.Drawing.Point(19, 54);
            this.Label228.Name = "Label228";
            this.Label228.Size = new System.Drawing.Size(92, 16);
            this.Label228.TabIndex = 27;
            this.Label228.Text = "( 参考件数 = ";
            this.Label228.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label227
            // 
            this.Label227.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label227.Location = new System.Drawing.Point(195, 54);
            this.Label227.Name = "Label227";
            this.Label227.Size = new System.Drawing.Size(22, 16);
            this.Label227.TabIndex = 29;
            this.Label227.Text = " )";
            this.Label227.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label226
            // 
            this.Label226.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label226.Location = new System.Drawing.Point(19, 33);
            this.Label226.Name = "Label226";
            this.Label226.Size = new System.Drawing.Size(198, 16);
            this.Label226.TabIndex = 26;
            this.Label226.Text = "箇所分類(例：外溝), クレーム分類";
            this.Label226.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiMstTitle
            // 
            this.pnlKiMstTitle.Controls.Add(this.Label346);
            this.pnlKiMstTitle.Controls.Add(this.Label236);
            this.pnlKiMstTitle.Controls.Add(this.Label233);
            this.pnlKiMstTitle.Controls.Add(this.lblKiMstGazotitleCnt);
            this.pnlKiMstTitle.Controls.Add(this.Label347);
            this.pnlKiMstTitle.Controls.Add(this.lblKiMstBikotitleCnt);
            this.pnlKiMstTitle.Controls.Add(this.Label235);
            this.pnlKiMstTitle.Controls.Add(this.chkKiMstTitle);
            this.pnlKiMstTitle.Controls.Add(this.lblKiMstKagititleCnt);
            this.pnlKiMstTitle.Controls.Add(this.Label232);
            this.pnlKiMstTitle.Controls.Add(this.Label80);
            this.pnlKiMstTitle.Location = new System.Drawing.Point(557, 112);
            this.pnlKiMstTitle.Name = "pnlKiMstTitle";
            this.pnlKiMstTitle.Size = new System.Drawing.Size(265, 120);
            this.pnlKiMstTitle.TabIndex = 49;
            // 
            // Label346
            // 
            this.Label346.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label346.Location = new System.Drawing.Point(19, 93);
            this.Label346.Name = "Label346";
            this.Label346.Size = new System.Drawing.Size(120, 16);
            this.Label346.TabIndex = 48;
            this.Label346.Text = "( 画像 参考件数 = ";
            this.Label346.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label236
            // 
            this.Label236.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label236.Location = new System.Drawing.Point(19, 73);
            this.Label236.Name = "Label236";
            this.Label236.Size = new System.Drawing.Size(120, 16);
            this.Label236.TabIndex = 45;
            this.Label236.Text = "( 備考 参考件数 = ";
            this.Label236.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label233
            // 
            this.Label233.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label233.Location = new System.Drawing.Point(19, 54);
            this.Label233.Name = "Label233";
            this.Label233.Size = new System.Drawing.Size(120, 16);
            this.Label233.TabIndex = 37;
            this.Label233.Text = "( 鍵    参考件数 = ";
            this.Label233.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblKiMstGazotitleCnt
            // 
            this.lblKiMstGazotitleCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstGazotitleCnt.Location = new System.Drawing.Point(124, 93);
            this.lblKiMstGazotitleCnt.Name = "lblKiMstGazotitleCnt";
            this.lblKiMstGazotitleCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstGazotitleCnt.TabIndex = 49;
            this.lblKiMstGazotitleCnt.Text = "9,999,999";
            this.lblKiMstGazotitleCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label347
            // 
            this.Label347.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label347.Location = new System.Drawing.Point(213, 93);
            this.Label347.Name = "Label347";
            this.Label347.Size = new System.Drawing.Size(22, 16);
            this.Label347.TabIndex = 50;
            this.Label347.Text = " )";
            this.Label347.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblKiMstBikotitleCnt
            // 
            this.lblKiMstBikotitleCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstBikotitleCnt.Location = new System.Drawing.Point(124, 73);
            this.lblKiMstBikotitleCnt.Name = "lblKiMstBikotitleCnt";
            this.lblKiMstBikotitleCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstBikotitleCnt.TabIndex = 46;
            this.lblKiMstBikotitleCnt.Text = "9,999,999";
            this.lblKiMstBikotitleCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label235
            // 
            this.Label235.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label235.Location = new System.Drawing.Point(213, 73);
            this.Label235.Name = "Label235";
            this.Label235.Size = new System.Drawing.Size(22, 16);
            this.Label235.TabIndex = 47;
            this.Label235.Text = " )";
            this.Label235.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // chkKiMstTitle
            // 
            this.chkKiMstTitle.AutoSize = true;
            this.chkKiMstTitle.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiMstTitle.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstTitle.Name = "chkKiMstTitle";
            this.chkKiMstTitle.Size = new System.Drawing.Size(134, 27);
            this.chkKiMstTitle.TabIndex = 35;
            this.chkKiMstTitle.Text = "タイトルマスタ";
            this.chkKiMstTitle.UseVisualStyleBackColor = true;
            this.chkKiMstTitle.CheckedChanged += new System.EventHandler(this.chkKiMstKagititle_CheckedChanged);
            // 
            // lblKiMstKagititleCnt
            // 
            this.lblKiMstKagititleCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstKagititleCnt.Location = new System.Drawing.Point(124, 54);
            this.lblKiMstKagititleCnt.Name = "lblKiMstKagititleCnt";
            this.lblKiMstKagititleCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstKagititleCnt.TabIndex = 38;
            this.lblKiMstKagititleCnt.Text = "9,999,999";
            this.lblKiMstKagititleCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label232
            // 
            this.Label232.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label232.Location = new System.Drawing.Point(213, 54);
            this.Label232.Name = "Label232";
            this.Label232.Size = new System.Drawing.Size(22, 16);
            this.Label232.TabIndex = 39;
            this.Label232.Text = " )";
            this.Label232.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label80
            // 
            this.Label80.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label80.Location = new System.Drawing.Point(19, 33);
            this.Label80.Name = "Label80";
            this.Label80.Size = new System.Drawing.Size(198, 16);
            this.Label80.TabIndex = 36;
            this.Label80.Text = "鍵、備考、画像タイトル";
            this.Label80.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiMstArea
            // 
            this.pnlKiMstArea.Controls.Add(this.chkKiMstArea);
            this.pnlKiMstArea.Controls.Add(this.lblKiMstAreaCnt);
            this.pnlKiMstArea.Controls.Add(this.Label216);
            this.pnlKiMstArea.Controls.Add(this.Label77);
            this.pnlKiMstArea.Controls.Add(this.Label76);
            this.pnlKiMstArea.Location = new System.Drawing.Point(15, 204);
            this.pnlKiMstArea.Name = "pnlKiMstArea";
            this.pnlKiMstArea.Size = new System.Drawing.Size(265, 80);
            this.pnlKiMstArea.TabIndex = 49;
            // 
            // chkKiMstArea
            // 
            this.chkKiMstArea.AutoSize = true;
            this.chkKiMstArea.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiMstArea.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstArea.Name = "chkKiMstArea";
            this.chkKiMstArea.Size = new System.Drawing.Size(119, 27);
            this.chkKiMstArea.TabIndex = 10;
            this.chkKiMstArea.Text = "エリアマスタ";
            this.chkKiMstArea.UseVisualStyleBackColor = true;
            this.chkKiMstArea.CheckedChanged += new System.EventHandler(this.chkKiMstArea_CheckedChanged);
            // 
            // lblKiMstAreaCnt
            // 
            this.lblKiMstAreaCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstAreaCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiMstAreaCnt.Name = "lblKiMstAreaCnt";
            this.lblKiMstAreaCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstAreaCnt.TabIndex = 13;
            this.lblKiMstAreaCnt.Text = "9,999,999";
            this.lblKiMstAreaCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label216
            // 
            this.Label216.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label216.Location = new System.Drawing.Point(19, 54);
            this.Label216.Name = "Label216";
            this.Label216.Size = new System.Drawing.Size(92, 16);
            this.Label216.TabIndex = 12;
            this.Label216.Text = "( 参考件数 = ";
            this.Label216.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label77
            // 
            this.Label77.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label77.Location = new System.Drawing.Point(195, 54);
            this.Label77.Name = "Label77";
            this.Label77.Size = new System.Drawing.Size(22, 16);
            this.Label77.TabIndex = 14;
            this.Label77.Text = " )";
            this.Label77.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label76
            // 
            this.Label76.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label76.Location = new System.Drawing.Point(19, 33);
            this.Label76.Name = "Label76";
            this.Label76.Size = new System.Drawing.Size(198, 16);
            this.Label76.TabIndex = 11;
            this.Label76.Text = "例：東部, 北部";
            this.Label76.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiMstSchool
            // 
            this.pnlKiMstSchool.Controls.Add(this.chkKiMstSchool);
            this.pnlKiMstSchool.Controls.Add(this.lblKiMstSchoolCnt);
            this.pnlKiMstSchool.Controls.Add(this.Label214);
            this.pnlKiMstSchool.Controls.Add(this.Label213);
            this.pnlKiMstSchool.Controls.Add(this.Label75);
            this.pnlKiMstSchool.Location = new System.Drawing.Point(15, 112);
            this.pnlKiMstSchool.Name = "pnlKiMstSchool";
            this.pnlKiMstSchool.Size = new System.Drawing.Size(265, 80);
            this.pnlKiMstSchool.TabIndex = 49;
            // 
            // chkKiMstSchool
            // 
            this.chkKiMstSchool.AutoSize = true;
            this.chkKiMstSchool.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiMstSchool.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkKiMstSchool.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstSchool.Name = "chkKiMstSchool";
            this.chkKiMstSchool.Size = new System.Drawing.Size(119, 27);
            this.chkKiMstSchool.TabIndex = 5;
            this.chkKiMstSchool.Text = "学校区マスタ";
            this.chkKiMstSchool.UseVisualStyleBackColor = true;
            this.chkKiMstSchool.CheckedChanged += new System.EventHandler(this.chkKiMstSchool_CheckedChanged);
            // 
            // lblKiMstSchoolCnt
            // 
            this.lblKiMstSchoolCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstSchoolCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiMstSchoolCnt.Name = "lblKiMstSchoolCnt";
            this.lblKiMstSchoolCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstSchoolCnt.TabIndex = 8;
            this.lblKiMstSchoolCnt.Text = "9,999,999";
            this.lblKiMstSchoolCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label214
            // 
            this.Label214.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label214.Location = new System.Drawing.Point(19, 54);
            this.Label214.Name = "Label214";
            this.Label214.Size = new System.Drawing.Size(92, 16);
            this.Label214.TabIndex = 7;
            this.Label214.Text = "( 参考件数 = ";
            this.Label214.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label213
            // 
            this.Label213.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label213.Location = new System.Drawing.Point(195, 54);
            this.Label213.Name = "Label213";
            this.Label213.Size = new System.Drawing.Size(22, 16);
            this.Label213.TabIndex = 9;
            this.Label213.Text = " )";
            this.Label213.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label75
            // 
            this.Label75.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label75.Location = new System.Drawing.Point(19, 33);
            this.Label75.Name = "Label75";
            this.Label75.Size = new System.Drawing.Size(198, 16);
            this.Label75.TabIndex = 6;
            this.Label75.Text = "小学校, 中学校";
            this.Label75.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiMstHokenrui
            // 
            this.pnlKiMstHokenrui.Controls.Add(this.chkKiMstHokenrui);
            this.pnlKiMstHokenrui.Controls.Add(this.lblKiMstHokenruiCnt);
            this.pnlKiMstHokenrui.Controls.Add(this.Label220);
            this.pnlKiMstHokenrui.Controls.Add(this.Label219);
            this.pnlKiMstHokenrui.Controls.Add(this.Label218);
            this.pnlKiMstHokenrui.Location = new System.Drawing.Point(286, 20);
            this.pnlKiMstHokenrui.Name = "pnlKiMstHokenrui";
            this.pnlKiMstHokenrui.Size = new System.Drawing.Size(265, 80);
            this.pnlKiMstHokenrui.TabIndex = 48;
            // 
            // chkKiMstHokenrui
            // 
            this.chkKiMstHokenrui.AutoSize = true;
            this.chkKiMstHokenrui.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiMstHokenrui.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstHokenrui.Name = "chkKiMstHokenrui";
            this.chkKiMstHokenrui.Size = new System.Drawing.Size(134, 27);
            this.chkKiMstHokenrui.TabIndex = 15;
            this.chkKiMstHokenrui.Text = "保険種類マスタ";
            this.chkKiMstHokenrui.UseVisualStyleBackColor = true;
            this.chkKiMstHokenrui.CheckedChanged += new System.EventHandler(this.chkKiMstHokenrui_CheckedChanged);
            // 
            // lblKiMstHokenruiCnt
            // 
            this.lblKiMstHokenruiCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiMstHokenruiCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiMstHokenruiCnt.Name = "lblKiMstHokenruiCnt";
            this.lblKiMstHokenruiCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstHokenruiCnt.TabIndex = 18;
            this.lblKiMstHokenruiCnt.Text = "9,999,999";
            this.lblKiMstHokenruiCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label220
            // 
            this.Label220.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label220.Location = new System.Drawing.Point(19, 54);
            this.Label220.Name = "Label220";
            this.Label220.Size = new System.Drawing.Size(92, 16);
            this.Label220.TabIndex = 17;
            this.Label220.Text = "( 参考件数 = ";
            this.Label220.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label219
            // 
            this.Label219.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label219.Location = new System.Drawing.Point(195, 54);
            this.Label219.Name = "Label219";
            this.Label219.Size = new System.Drawing.Size(22, 16);
            this.Label219.TabIndex = 19;
            this.Label219.Text = " )";
            this.Label219.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label218
            // 
            this.Label218.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label218.Location = new System.Drawing.Point(19, 33);
            this.Label218.Name = "Label218";
            this.Label218.Size = new System.Drawing.Size(198, 16);
            this.Label218.TabIndex = 16;
            this.Label218.Text = "例：住宅総合保険";
            this.Label218.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiMstBus
            // 
            this.pnlKiMstBus.Controls.Add(this.chkKiMstBus);
            this.pnlKiMstBus.Controls.Add(this.lblKiMstBusCnt);
            this.pnlKiMstBus.Controls.Add(this.Label211);
            this.pnlKiMstBus.Controls.Add(this.Label181);
            this.pnlKiMstBus.Controls.Add(this.Label210);
            this.pnlKiMstBus.Location = new System.Drawing.Point(15, 20);
            this.pnlKiMstBus.Name = "pnlKiMstBus";
            this.pnlKiMstBus.Size = new System.Drawing.Size(265, 80);
            this.pnlKiMstBus.TabIndex = 47;
            // 
            // chkKiMstBus
            // 
            this.chkKiMstBus.AutoSize = true;
            this.chkKiMstBus.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiMstBus.Location = new System.Drawing.Point(3, 3);
            this.chkKiMstBus.Name = "chkKiMstBus";
            this.chkKiMstBus.Size = new System.Drawing.Size(134, 27);
            this.chkKiMstBus.TabIndex = 0;
            this.chkKiMstBus.Text = "バス交通マスタ";
            this.chkKiMstBus.UseVisualStyleBackColor = true;
            this.chkKiMstBus.CheckedChanged += new System.EventHandler(this.chkKiMstBus_CheckedChanged);
            // 
            // lblKiMstBusCnt
            // 
            this.lblKiMstBusCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiMstBusCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiMstBusCnt.Name = "lblKiMstBusCnt";
            this.lblKiMstBusCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiMstBusCnt.TabIndex = 3;
            this.lblKiMstBusCnt.Text = "9,999,999";
            this.lblKiMstBusCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label211
            // 
            this.Label211.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label211.Location = new System.Drawing.Point(19, 54);
            this.Label211.Name = "Label211";
            this.Label211.Size = new System.Drawing.Size(92, 16);
            this.Label211.TabIndex = 2;
            this.Label211.Text = "( 参考件数 = ";
            this.Label211.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label181
            // 
            this.Label181.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label181.Location = new System.Drawing.Point(195, 54);
            this.Label181.Name = "Label181";
            this.Label181.Size = new System.Drawing.Size(22, 16);
            this.Label181.TabIndex = 4;
            this.Label181.Text = " )";
            this.Label181.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label210
            // 
            this.Label210.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label210.Location = new System.Drawing.Point(19, 33);
            this.Label210.Name = "Label210";
            this.Label210.Size = new System.Drawing.Size(198, 16);
            this.Label210.TabIndex = 1;
            this.Label210.Text = "バス会社, 系統, バス停";
            this.Label210.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblSelectPageCnt1
            // 
            this.lblSelectPageCnt1.AutoSize = true;
            this.lblSelectPageCnt1.Location = new System.Drawing.Point(806, 306);
            this.lblSelectPageCnt1.Name = "lblSelectPageCnt1";
            this.lblSelectPageCnt1.Size = new System.Drawing.Size(35, 18);
            this.lblSelectPageCnt1.TabIndex = 46;
            this.lblSelectPageCnt1.Text = "1 / 3";
            // 
            // tabPageKizon120
            // 
            this.tabPageKizon120.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageKizon120.Controls.Add(this.grpKizon3);
            this.tabPageKizon120.Controls.Add(this.grpKizon5);
            this.tabPageKizon120.Controls.Add(this.lblSelectPageCnt2);
            this.tabPageKizon120.Font = new System.Drawing.Font("メイリオ", 9F);
            this.tabPageKizon120.Location = new System.Drawing.Point(4, 27);
            this.tabPageKizon120.Name = "tabPageKizon120";
            this.tabPageKizon120.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageKizon120.Size = new System.Drawing.Size(847, 327);
            this.tabPageKizon120.TabIndex = 10;
            this.tabPageKizon120.Text = " 項目選択2";
            // 
            // grpKizon3
            // 
            this.grpKizon3.Controls.Add(this.pnlKiGySyuzenBase);
            this.grpKizon3.Controls.Add(this.pnlKiGyYatinhosyoBase);
            this.grpKizon3.Controls.Add(this.pnlKiGyLifelineBase);
            this.grpKizon3.Controls.Add(this.pnlKiGyHokenBase);
            this.grpKizon3.Controls.Add(this.pnlKiGySisetuBase);
            this.grpKizon3.Controls.Add(this.pnlKiGyCyukaiBase);
            this.grpKizon3.Controls.Add(this.pnlKiGySekoBase);
            this.grpKizon3.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpKizon3.Location = new System.Drawing.Point(304, 5);
            this.grpKizon3.Name = "grpKizon3";
            this.grpKizon3.Size = new System.Drawing.Size(537, 298);
            this.grpKizon3.TabIndex = 1;
            this.grpKizon3.TabStop = false;
            this.grpKizon3.Text = "【業者情報】";
            // 
            // pnlKiGySyuzenBase
            // 
            this.pnlKiGySyuzenBase.Controls.Add(this.chkKiGySyuzenBase);
            this.pnlKiGySyuzenBase.Controls.Add(this.lblKiGySyuzenBaseCnt);
            this.pnlKiGySyuzenBase.Controls.Add(this.Label241);
            this.pnlKiGySyuzenBase.Controls.Add(this.Label240);
            this.pnlKiGySyuzenBase.Location = new System.Drawing.Point(15, 89);
            this.pnlKiGySyuzenBase.Name = "pnlKiGySyuzenBase";
            this.pnlKiGySyuzenBase.Size = new System.Drawing.Size(255, 60);
            this.pnlKiGySyuzenBase.TabIndex = 19;
            // 
            // chkKiGySyuzenBase
            // 
            this.chkKiGySyuzenBase.AutoSize = true;
            this.chkKiGySyuzenBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiGySyuzenBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiGySyuzenBase.Name = "chkKiGySyuzenBase";
            this.chkKiGySyuzenBase.Size = new System.Drawing.Size(119, 27);
            this.chkKiGySyuzenBase.TabIndex = 4;
            this.chkKiGySyuzenBase.Text = "修繕業者情報";
            this.chkKiGySyuzenBase.UseVisualStyleBackColor = true;
            this.chkKiGySyuzenBase.CheckedChanged += new System.EventHandler(this.chkKiGySyuzenBase_CheckedChanged);
            // 
            // lblKiGySyuzenBaseCnt
            // 
            this.lblKiGySyuzenBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiGySyuzenBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiGySyuzenBaseCnt.Name = "lblKiGySyuzenBaseCnt";
            this.lblKiGySyuzenBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiGySyuzenBaseCnt.TabIndex = 6;
            this.lblKiGySyuzenBaseCnt.Text = "9,999,999";
            this.lblKiGySyuzenBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label241
            // 
            this.Label241.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label241.Location = new System.Drawing.Point(19, 33);
            this.Label241.Name = "Label241";
            this.Label241.Size = new System.Drawing.Size(92, 16);
            this.Label241.TabIndex = 5;
            this.Label241.Text = "( 参考件数 = ";
            this.Label241.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label240
            // 
            this.Label240.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label240.Location = new System.Drawing.Point(195, 33);
            this.Label240.Name = "Label240";
            this.Label240.Size = new System.Drawing.Size(22, 16);
            this.Label240.TabIndex = 7;
            this.Label240.Text = " )";
            this.Label240.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiGyYatinhosyoBase
            // 
            this.pnlKiGyYatinhosyoBase.Controls.Add(this.chkKiGyYatinhosyoBase);
            this.pnlKiGyYatinhosyoBase.Controls.Add(this.lblKiGyYatinhosyoBaseCnt);
            this.pnlKiGyYatinhosyoBase.Controls.Add(this.Label250);
            this.pnlKiGyYatinhosyoBase.Controls.Add(this.Label249);
            this.pnlKiGyYatinhosyoBase.Location = new System.Drawing.Point(276, 20);
            this.pnlKiGyYatinhosyoBase.Name = "pnlKiGyYatinhosyoBase";
            this.pnlKiGyYatinhosyoBase.Size = new System.Drawing.Size(255, 60);
            this.pnlKiGyYatinhosyoBase.TabIndex = 19;
            // 
            // chkKiGyYatinhosyoBase
            // 
            this.chkKiGyYatinhosyoBase.AutoSize = true;
            this.chkKiGyYatinhosyoBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiGyYatinhosyoBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiGyYatinhosyoBase.Name = "chkKiGyYatinhosyoBase";
            this.chkKiGyYatinhosyoBase.Size = new System.Drawing.Size(149, 27);
            this.chkKiGyYatinhosyoBase.TabIndex = 16;
            this.chkKiGyYatinhosyoBase.Text = "家賃保証業者情報";
            this.chkKiGyYatinhosyoBase.UseVisualStyleBackColor = true;
            this.chkKiGyYatinhosyoBase.CheckedChanged += new System.EventHandler(this.chkKiGyYatinhosyoBase_CheckedChanged);
            // 
            // lblKiGyYatinhosyoBaseCnt
            // 
            this.lblKiGyYatinhosyoBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiGyYatinhosyoBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiGyYatinhosyoBaseCnt.Name = "lblKiGyYatinhosyoBaseCnt";
            this.lblKiGyYatinhosyoBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiGyYatinhosyoBaseCnt.TabIndex = 18;
            this.lblKiGyYatinhosyoBaseCnt.Text = "9,999,999";
            this.lblKiGyYatinhosyoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label250
            // 
            this.Label250.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label250.Location = new System.Drawing.Point(19, 33);
            this.Label250.Name = "Label250";
            this.Label250.Size = new System.Drawing.Size(92, 16);
            this.Label250.TabIndex = 17;
            this.Label250.Text = "( 参考件数 = ";
            this.Label250.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label249
            // 
            this.Label249.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label249.Location = new System.Drawing.Point(195, 33);
            this.Label249.Name = "Label249";
            this.Label249.Size = new System.Drawing.Size(22, 16);
            this.Label249.TabIndex = 19;
            this.Label249.Text = " )";
            this.Label249.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiGyLifelineBase
            // 
            this.pnlKiGyLifelineBase.Controls.Add(this.chkKiGyLifelineBase);
            this.pnlKiGyLifelineBase.Controls.Add(this.lblKiGyLifelineBaseCnt);
            this.pnlKiGyLifelineBase.Controls.Add(this.Label244);
            this.pnlKiGyLifelineBase.Controls.Add(this.Label243);
            this.pnlKiGyLifelineBase.Location = new System.Drawing.Point(15, 158);
            this.pnlKiGyLifelineBase.Name = "pnlKiGyLifelineBase";
            this.pnlKiGyLifelineBase.Size = new System.Drawing.Size(255, 60);
            this.pnlKiGyLifelineBase.TabIndex = 19;
            // 
            // chkKiGyLifelineBase
            // 
            this.chkKiGyLifelineBase.AutoSize = true;
            this.chkKiGyLifelineBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiGyLifelineBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiGyLifelineBase.Name = "chkKiGyLifelineBase";
            this.chkKiGyLifelineBase.Size = new System.Drawing.Size(179, 27);
            this.chkKiGyLifelineBase.TabIndex = 8;
            this.chkKiGyLifelineBase.Text = "ライフライン業者情報";
            this.chkKiGyLifelineBase.UseVisualStyleBackColor = true;
            this.chkKiGyLifelineBase.CheckedChanged += new System.EventHandler(this.chkKiGyLifelineBase_CheckedChanged);
            // 
            // lblKiGyLifelineBaseCnt
            // 
            this.lblKiGyLifelineBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiGyLifelineBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiGyLifelineBaseCnt.Name = "lblKiGyLifelineBaseCnt";
            this.lblKiGyLifelineBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiGyLifelineBaseCnt.TabIndex = 10;
            this.lblKiGyLifelineBaseCnt.Text = "9,999,999";
            this.lblKiGyLifelineBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label244
            // 
            this.Label244.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label244.Location = new System.Drawing.Point(19, 33);
            this.Label244.Name = "Label244";
            this.Label244.Size = new System.Drawing.Size(92, 16);
            this.Label244.TabIndex = 9;
            this.Label244.Text = "( 参考件数 = ";
            this.Label244.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label243
            // 
            this.Label243.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label243.Location = new System.Drawing.Point(195, 33);
            this.Label243.Name = "Label243";
            this.Label243.Size = new System.Drawing.Size(22, 16);
            this.Label243.TabIndex = 11;
            this.Label243.Text = " )";
            this.Label243.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiGyHokenBase
            // 
            this.pnlKiGyHokenBase.Controls.Add(this.chkKiGyHokenBase);
            this.pnlKiGyHokenBase.Controls.Add(this.lblKiGyHokenBaseCnt);
            this.pnlKiGyHokenBase.Controls.Add(this.Label247);
            this.pnlKiGyHokenBase.Controls.Add(this.Label246);
            this.pnlKiGyHokenBase.Location = new System.Drawing.Point(15, 227);
            this.pnlKiGyHokenBase.Name = "pnlKiGyHokenBase";
            this.pnlKiGyHokenBase.Size = new System.Drawing.Size(255, 60);
            this.pnlKiGyHokenBase.TabIndex = 19;
            // 
            // chkKiGyHokenBase
            // 
            this.chkKiGyHokenBase.AutoSize = true;
            this.chkKiGyHokenBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiGyHokenBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiGyHokenBase.Name = "chkKiGyHokenBase";
            this.chkKiGyHokenBase.Size = new System.Drawing.Size(119, 27);
            this.chkKiGyHokenBase.TabIndex = 12;
            this.chkKiGyHokenBase.Text = "保険業者情報";
            this.chkKiGyHokenBase.UseVisualStyleBackColor = true;
            this.chkKiGyHokenBase.CheckedChanged += new System.EventHandler(this.chkKiGyHokenBase_CheckedChanged);
            // 
            // lblKiGyHokenBaseCnt
            // 
            this.lblKiGyHokenBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiGyHokenBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiGyHokenBaseCnt.Name = "lblKiGyHokenBaseCnt";
            this.lblKiGyHokenBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiGyHokenBaseCnt.TabIndex = 14;
            this.lblKiGyHokenBaseCnt.Text = "9,999,999";
            this.lblKiGyHokenBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label247
            // 
            this.Label247.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label247.Location = new System.Drawing.Point(19, 33);
            this.Label247.Name = "Label247";
            this.Label247.Size = new System.Drawing.Size(92, 16);
            this.Label247.TabIndex = 13;
            this.Label247.Text = "( 参考件数 = ";
            this.Label247.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label246
            // 
            this.Label246.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label246.Location = new System.Drawing.Point(195, 33);
            this.Label246.Name = "Label246";
            this.Label246.Size = new System.Drawing.Size(22, 16);
            this.Label246.TabIndex = 15;
            this.Label246.Text = " )";
            this.Label246.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiGySisetuBase
            // 
            this.pnlKiGySisetuBase.Controls.Add(this.chkKiGySisetuBase);
            this.pnlKiGySisetuBase.Controls.Add(this.lblKiGySisetuBaseCnt);
            this.pnlKiGySisetuBase.Controls.Add(this.Label253);
            this.pnlKiGySisetuBase.Controls.Add(this.Label252);
            this.pnlKiGySisetuBase.Location = new System.Drawing.Point(276, 89);
            this.pnlKiGySisetuBase.Name = "pnlKiGySisetuBase";
            this.pnlKiGySisetuBase.Size = new System.Drawing.Size(255, 60);
            this.pnlKiGySisetuBase.TabIndex = 19;
            // 
            // chkKiGySisetuBase
            // 
            this.chkKiGySisetuBase.AutoSize = true;
            this.chkKiGySisetuBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiGySisetuBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiGySisetuBase.Name = "chkKiGySisetuBase";
            this.chkKiGySisetuBase.Size = new System.Drawing.Size(149, 27);
            this.chkKiGySisetuBase.TabIndex = 20;
            this.chkKiGySisetuBase.Text = "施設保守業者情報";
            this.chkKiGySisetuBase.UseVisualStyleBackColor = true;
            this.chkKiGySisetuBase.CheckedChanged += new System.EventHandler(this.chkKiGySisetuBase_CheckedChanged);
            // 
            // lblKiGySisetuBaseCnt
            // 
            this.lblKiGySisetuBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiGySisetuBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiGySisetuBaseCnt.Name = "lblKiGySisetuBaseCnt";
            this.lblKiGySisetuBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiGySisetuBaseCnt.TabIndex = 22;
            this.lblKiGySisetuBaseCnt.Text = "9,999,999";
            this.lblKiGySisetuBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label253
            // 
            this.Label253.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label253.Location = new System.Drawing.Point(19, 33);
            this.Label253.Name = "Label253";
            this.Label253.Size = new System.Drawing.Size(92, 16);
            this.Label253.TabIndex = 21;
            this.Label253.Text = "( 参考件数 = ";
            this.Label253.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label252
            // 
            this.Label252.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label252.Location = new System.Drawing.Point(195, 33);
            this.Label252.Name = "Label252";
            this.Label252.Size = new System.Drawing.Size(22, 16);
            this.Label252.TabIndex = 23;
            this.Label252.Text = " )";
            this.Label252.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiGyCyukaiBase
            // 
            this.pnlKiGyCyukaiBase.Controls.Add(this.chkKiGyCyukaiBase);
            this.pnlKiGyCyukaiBase.Controls.Add(this.lblKiGyCyukaiBaseCnt);
            this.pnlKiGyCyukaiBase.Controls.Add(this.Label238);
            this.pnlKiGyCyukaiBase.Controls.Add(this.Label87);
            this.pnlKiGyCyukaiBase.Location = new System.Drawing.Point(15, 20);
            this.pnlKiGyCyukaiBase.Name = "pnlKiGyCyukaiBase";
            this.pnlKiGyCyukaiBase.Size = new System.Drawing.Size(255, 60);
            this.pnlKiGyCyukaiBase.TabIndex = 19;
            // 
            // chkKiGyCyukaiBase
            // 
            this.chkKiGyCyukaiBase.AutoSize = true;
            this.chkKiGyCyukaiBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiGyCyukaiBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiGyCyukaiBase.Name = "chkKiGyCyukaiBase";
            this.chkKiGyCyukaiBase.Size = new System.Drawing.Size(164, 27);
            this.chkKiGyCyukaiBase.TabIndex = 0;
            this.chkKiGyCyukaiBase.Text = "仲介・管理業者情報";
            this.chkKiGyCyukaiBase.UseVisualStyleBackColor = true;
            this.chkKiGyCyukaiBase.CheckedChanged += new System.EventHandler(this.chkKiGyCyukaiBase_CheckedChanged);
            // 
            // lblKiGyCyukaiBaseCnt
            // 
            this.lblKiGyCyukaiBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiGyCyukaiBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiGyCyukaiBaseCnt.Name = "lblKiGyCyukaiBaseCnt";
            this.lblKiGyCyukaiBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiGyCyukaiBaseCnt.TabIndex = 2;
            this.lblKiGyCyukaiBaseCnt.Text = "9,999,999";
            this.lblKiGyCyukaiBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label238
            // 
            this.Label238.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label238.Location = new System.Drawing.Point(19, 33);
            this.Label238.Name = "Label238";
            this.Label238.Size = new System.Drawing.Size(92, 16);
            this.Label238.TabIndex = 1;
            this.Label238.Text = "( 参考件数 = ";
            this.Label238.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label87
            // 
            this.Label87.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label87.Location = new System.Drawing.Point(195, 33);
            this.Label87.Name = "Label87";
            this.Label87.Size = new System.Drawing.Size(22, 16);
            this.Label87.TabIndex = 3;
            this.Label87.Text = " )";
            this.Label87.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiGySekoBase
            // 
            this.pnlKiGySekoBase.Controls.Add(this.chkKiGySekoBase);
            this.pnlKiGySekoBase.Controls.Add(this.lblKiGySekoBaseCnt);
            this.pnlKiGySekoBase.Controls.Add(this.Label256);
            this.pnlKiGySekoBase.Controls.Add(this.Label255);
            this.pnlKiGySekoBase.Location = new System.Drawing.Point(276, 158);
            this.pnlKiGySekoBase.Name = "pnlKiGySekoBase";
            this.pnlKiGySekoBase.Size = new System.Drawing.Size(255, 60);
            this.pnlKiGySekoBase.TabIndex = 19;
            // 
            // chkKiGySekoBase
            // 
            this.chkKiGySekoBase.AutoSize = true;
            this.chkKiGySekoBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiGySekoBase.Location = new System.Drawing.Point(6, 3);
            this.chkKiGySekoBase.Name = "chkKiGySekoBase";
            this.chkKiGySekoBase.Size = new System.Drawing.Size(119, 27);
            this.chkKiGySekoBase.TabIndex = 24;
            this.chkKiGySekoBase.Text = "施工業者情報";
            this.chkKiGySekoBase.UseVisualStyleBackColor = true;
            this.chkKiGySekoBase.CheckedChanged += new System.EventHandler(this.chkKiGySekoBase_CheckedChanged);
            // 
            // lblKiGySekoBaseCnt
            // 
            this.lblKiGySekoBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiGySekoBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiGySekoBaseCnt.Name = "lblKiGySekoBaseCnt";
            this.lblKiGySekoBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiGySekoBaseCnt.TabIndex = 26;
            this.lblKiGySekoBaseCnt.Text = "9,999,999";
            this.lblKiGySekoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label256
            // 
            this.Label256.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label256.Location = new System.Drawing.Point(19, 33);
            this.Label256.Name = "Label256";
            this.Label256.Size = new System.Drawing.Size(92, 16);
            this.Label256.TabIndex = 25;
            this.Label256.Text = "( 参考件数 = ";
            this.Label256.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label255
            // 
            this.Label255.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label255.Location = new System.Drawing.Point(195, 33);
            this.Label255.Name = "Label255";
            this.Label255.Size = new System.Drawing.Size(22, 16);
            this.Label255.TabIndex = 27;
            this.Label255.Text = " )";
            this.Label255.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // grpKizon5
            // 
            this.grpKizon5.Controls.Add(this.pnlKiOw);
            this.grpKizon5.Controls.Add(this.pnlKiJisya);
            this.grpKizon5.Controls.Add(this.pnlKiSyskanriBase);
            this.grpKizon5.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpKizon5.Location = new System.Drawing.Point(6, 5);
            this.grpKizon5.Name = "grpKizon5";
            this.grpKizon5.Size = new System.Drawing.Size(292, 298);
            this.grpKizon5.TabIndex = 0;
            this.grpKizon5.TabStop = false;
            this.grpKizon5.Text = "【基本情報1】";
            // 
            // pnlKiOw
            // 
            this.pnlKiOw.Controls.Add(this.chkKiOwBase);
            this.pnlKiOw.Controls.Add(this.lblKiOwBaseCnt);
            this.pnlKiOw.Controls.Add(this.Label261);
            this.pnlKiOw.Controls.Add(this.Label268);
            this.pnlKiOw.Controls.Add(this.Label267);
            this.pnlKiOw.Location = new System.Drawing.Point(15, 204);
            this.pnlKiOw.Name = "pnlKiOw";
            this.pnlKiOw.Size = new System.Drawing.Size(265, 80);
            this.pnlKiOw.TabIndex = 17;
            // 
            // chkKiOwBase
            // 
            this.chkKiOwBase.AutoSize = true;
            this.chkKiOwBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiOwBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiOwBase.Name = "chkKiOwBase";
            this.chkKiOwBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiOwBase.TabIndex = 10;
            this.chkKiOwBase.Text = "家主情報";
            this.chkKiOwBase.UseVisualStyleBackColor = true;
            this.chkKiOwBase.CheckedChanged += new System.EventHandler(this.chkKiOwBase_CheckedChanged);
            // 
            // lblKiOwBaseCnt
            // 
            this.lblKiOwBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiOwBaseCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiOwBaseCnt.Name = "lblKiOwBaseCnt";
            this.lblKiOwBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiOwBaseCnt.TabIndex = 13;
            this.lblKiOwBaseCnt.Text = "9,999,999";
            this.lblKiOwBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label261
            // 
            this.Label261.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label261.Location = new System.Drawing.Point(19, 33);
            this.Label261.Name = "Label261";
            this.Label261.Size = new System.Drawing.Size(198, 16);
            this.Label261.TabIndex = 11;
            this.Label261.Text = "家主情報 口座";
            this.Label261.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label268
            // 
            this.Label268.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label268.Location = new System.Drawing.Point(19, 54);
            this.Label268.Name = "Label268";
            this.Label268.Size = new System.Drawing.Size(92, 16);
            this.Label268.TabIndex = 12;
            this.Label268.Text = "( 参考件数 = ";
            this.Label268.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label267
            // 
            this.Label267.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label267.Location = new System.Drawing.Point(195, 54);
            this.Label267.Name = "Label267";
            this.Label267.Size = new System.Drawing.Size(22, 16);
            this.Label267.TabIndex = 14;
            this.Label267.Text = " )";
            this.Label267.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiJisya
            // 
            this.pnlKiJisya.Controls.Add(this.chkKiJisyaBase);
            this.pnlKiJisya.Controls.Add(this.lblKiJisyaBaseCnt);
            this.pnlKiJisya.Controls.Add(this.Label265);
            this.pnlKiJisya.Controls.Add(this.Label264);
            this.pnlKiJisya.Controls.Add(this.Label263);
            this.pnlKiJisya.Location = new System.Drawing.Point(15, 112);
            this.pnlKiJisya.Name = "pnlKiJisya";
            this.pnlKiJisya.Size = new System.Drawing.Size(265, 80);
            this.pnlKiJisya.TabIndex = 16;
            // 
            // chkKiJisyaBase
            // 
            this.chkKiJisyaBase.AutoSize = true;
            this.chkKiJisyaBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiJisyaBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiJisyaBase.Name = "chkKiJisyaBase";
            this.chkKiJisyaBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiJisyaBase.TabIndex = 5;
            this.chkKiJisyaBase.Text = "自社情報";
            this.chkKiJisyaBase.UseVisualStyleBackColor = true;
            this.chkKiJisyaBase.CheckedChanged += new System.EventHandler(this.chkKiJisyaBase_CheckedChanged);
            // 
            // lblKiJisyaBaseCnt
            // 
            this.lblKiJisyaBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiJisyaBaseCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiJisyaBaseCnt.Name = "lblKiJisyaBaseCnt";
            this.lblKiJisyaBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiJisyaBaseCnt.TabIndex = 8;
            this.lblKiJisyaBaseCnt.Text = "9,999,999";
            this.lblKiJisyaBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label265
            // 
            this.Label265.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label265.Location = new System.Drawing.Point(19, 54);
            this.Label265.Name = "Label265";
            this.Label265.Size = new System.Drawing.Size(92, 16);
            this.Label265.TabIndex = 7;
            this.Label265.Text = "( 参考件数 = ";
            this.Label265.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label264
            // 
            this.Label264.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label264.Location = new System.Drawing.Point(195, 54);
            this.Label264.Name = "Label264";
            this.Label264.Size = new System.Drawing.Size(22, 16);
            this.Label264.TabIndex = 9;
            this.Label264.Text = " )";
            this.Label264.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label263
            // 
            this.Label263.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label263.Location = new System.Drawing.Point(19, 33);
            this.Label263.Name = "Label263";
            this.Label263.Size = new System.Drawing.Size(198, 16);
            this.Label263.TabIndex = 6;
            this.Label263.Text = "自社支店, 口座, 担当者";
            this.Label263.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiSyskanriBase
            // 
            this.pnlKiSyskanriBase.Controls.Add(this.chkKiSyskanriBase);
            this.pnlKiSyskanriBase.Controls.Add(this.lblKiSyskanriBaseCnt);
            this.pnlKiSyskanriBase.Controls.Add(this.Label260);
            this.pnlKiSyskanriBase.Controls.Add(this.Label259);
            this.pnlKiSyskanriBase.Controls.Add(this.Label258);
            this.pnlKiSyskanriBase.Location = new System.Drawing.Point(15, 20);
            this.pnlKiSyskanriBase.Name = "pnlKiSyskanriBase";
            this.pnlKiSyskanriBase.Size = new System.Drawing.Size(265, 80);
            this.pnlKiSyskanriBase.TabIndex = 2;
            // 
            // chkKiSyskanriBase
            // 
            this.chkKiSyskanriBase.AutoSize = true;
            this.chkKiSyskanriBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiSyskanriBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiSyskanriBase.Name = "chkKiSyskanriBase";
            this.chkKiSyskanriBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiSyskanriBase.TabIndex = 0;
            this.chkKiSyskanriBase.Text = "初期設定";
            this.chkKiSyskanriBase.UseVisualStyleBackColor = true;
            this.chkKiSyskanriBase.CheckedChanged += new System.EventHandler(this.chkKiSyskanriBase_CheckedChanged);
            // 
            // lblKiSyskanriBaseCnt
            // 
            this.lblKiSyskanriBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiSyskanriBaseCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiSyskanriBaseCnt.Name = "lblKiSyskanriBaseCnt";
            this.lblKiSyskanriBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiSyskanriBaseCnt.TabIndex = 3;
            this.lblKiSyskanriBaseCnt.Text = "9,999,999";
            this.lblKiSyskanriBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label260
            // 
            this.Label260.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label260.Location = new System.Drawing.Point(19, 54);
            this.Label260.Name = "Label260";
            this.Label260.Size = new System.Drawing.Size(92, 16);
            this.Label260.TabIndex = 2;
            this.Label260.Text = "( 参考件数 = ";
            this.Label260.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label259
            // 
            this.Label259.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label259.Location = new System.Drawing.Point(195, 54);
            this.Label259.Name = "Label259";
            this.Label259.Size = new System.Drawing.Size(22, 16);
            this.Label259.TabIndex = 4;
            this.Label259.Text = " )";
            this.Label259.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label258
            // 
            this.Label258.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label258.Location = new System.Drawing.Point(19, 33);
            this.Label258.Name = "Label258";
            this.Label258.Size = new System.Drawing.Size(198, 16);
            this.Label258.TabIndex = 1;
            this.Label258.Text = "管理情報設定, タイトル";
            this.Label258.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblSelectPageCnt2
            // 
            this.lblSelectPageCnt2.AutoSize = true;
            this.lblSelectPageCnt2.Location = new System.Drawing.Point(806, 306);
            this.lblSelectPageCnt2.Name = "lblSelectPageCnt2";
            this.lblSelectPageCnt2.Size = new System.Drawing.Size(35, 18);
            this.lblSelectPageCnt2.TabIndex = 28;
            this.lblSelectPageCnt2.Text = "2 / 3";
            // 
            // tabPageKizon130
            // 
            this.tabPageKizon130.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageKizon130.Controls.Add(this.grpKizon2);
            this.tabPageKizon130.Controls.Add(this.Label272);
            this.tabPageKizon130.Location = new System.Drawing.Point(4, 27);
            this.tabPageKizon130.Name = "tabPageKizon130";
            this.tabPageKizon130.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageKizon130.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.tabPageKizon130.Size = new System.Drawing.Size(847, 327);
            this.tabPageKizon130.TabIndex = 11;
            this.tabPageKizon130.Text = " 項目選択3";
            // 
            // grpKizon2
            // 
            this.grpKizon2.Controls.Add(this.Label359);
            this.grpKizon2.Controls.Add(this.pnlRendo);
            this.grpKizon2.Controls.Add(this.pnlKiSq);
            this.grpKizon2.Controls.Add(this.pnlSzen);
            this.grpKizon2.Controls.Add(this.pnlKiClaim);
            this.grpKizon2.Controls.Add(this.pnlKiKy);
            this.grpKizon2.Controls.Add(this.pnlKiKys);
            this.grpKizon2.Controls.Add(this.pnlKiHy);
            this.grpKizon2.Controls.Add(this.pnlKiBk);
            this.grpKizon2.Controls.Add(this.grpKizonKagi);
            this.grpKizon2.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpKizon2.Location = new System.Drawing.Point(6, 5);
            this.grpKizon2.Name = "grpKizon2";
            this.grpKizon2.Size = new System.Drawing.Size(835, 298);
            this.grpKizon2.TabIndex = 0;
            this.grpKizon2.TabStop = false;
            this.grpKizon2.Text = "【基本情報2】";
            // 
            // Label359
            // 
            this.Label359.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label359.Location = new System.Drawing.Point(733, 262);
            this.Label359.Name = "Label359";
            this.Label359.Size = new System.Drawing.Size(85, 16);
            this.Label359.TabIndex = 38;
            this.Label359.Text = "(指定IDを設定)";
            this.Label359.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlRendo
            // 
            this.pnlRendo.Controls.Add(this.Label83);
            this.pnlRendo.Controls.Add(this.txtRendoID);
            this.pnlRendo.Controls.Add(this.Label121);
            this.pnlRendo.Controls.Add(this.chkKiRendoBase);
            this.pnlRendo.Controls.Add(this.lblKiRendoBaseCnt);
            this.pnlRendo.Controls.Add(this.Label314);
            this.pnlRendo.Controls.Add(this.Label316);
            this.pnlRendo.Location = new System.Drawing.Point(557, 189);
            this.pnlRendo.Name = "pnlRendo";
            this.pnlRendo.Size = new System.Drawing.Size(265, 99);
            this.pnlRendo.TabIndex = 45;
            // 
            // Label83
            // 
            this.Label83.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label83.Location = new System.Drawing.Point(32, 73);
            this.Label83.Name = "Label83";
            this.Label83.Size = new System.Drawing.Size(60, 16);
            this.Label83.TabIndex = 37;
            this.Label83.Text = "連動ID：";
            this.Label83.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtRendoID
            // 
            this.txtRendoID.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtRendoID.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.txtRendoID.Location = new System.Drawing.Point(94, 70);
            this.txtRendoID.Name = "txtRendoID";
            this.txtRendoID.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.txtRendoID.Size = new System.Drawing.Size(71, 24);
            this.txtRendoID.TabIndex = 36;
            this.txtRendoID.Text = "830013";
            // 
            // Label121
            // 
            this.Label121.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label121.Location = new System.Drawing.Point(19, 33);
            this.Label121.Name = "Label121";
            this.Label121.Size = new System.Drawing.Size(198, 16);
            this.Label121.TabIndex = 35;
            this.Label121.Text = "送信設定, 広告補足, 周辺環境";
            this.Label121.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkKiRendoBase
            // 
            this.chkKiRendoBase.AutoSize = true;
            this.chkKiRendoBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline);
            this.chkKiRendoBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiRendoBase.Name = "chkKiRendoBase";
            this.chkKiRendoBase.Size = new System.Drawing.Size(149, 27);
            this.chkKiRendoBase.TabIndex = 29;
            this.chkKiRendoBase.Text = "ポータル連動情報";
            this.chkKiRendoBase.UseVisualStyleBackColor = true;
            this.chkKiRendoBase.CheckedChanged += new System.EventHandler(this.chkKiRendoBase_CheckedChanged);
            // 
            // lblKiRendoBaseCnt
            // 
            this.lblKiRendoBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F);
            this.lblKiRendoBaseCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiRendoBaseCnt.Name = "lblKiRendoBaseCnt";
            this.lblKiRendoBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiRendoBaseCnt.TabIndex = 31;
            this.lblKiRendoBaseCnt.Text = "9,999,999";
            this.lblKiRendoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label314
            // 
            this.Label314.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label314.Location = new System.Drawing.Point(19, 54);
            this.Label314.Name = "Label314";
            this.Label314.Size = new System.Drawing.Size(92, 16);
            this.Label314.TabIndex = 30;
            this.Label314.Text = "( 参考件数 = ";
            this.Label314.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label316
            // 
            this.Label316.Font = new System.Drawing.Font("メイリオ", 9F);
            this.Label316.Location = new System.Drawing.Point(195, 54);
            this.Label316.Name = "Label316";
            this.Label316.Size = new System.Drawing.Size(22, 16);
            this.Label316.TabIndex = 32;
            this.Label316.Text = " )";
            this.Label316.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiSq
            // 
            this.pnlKiSq.Controls.Add(this.Label308);
            this.pnlKiSq.Controls.Add(this.lblKiSqOwKojoBaseCnt);
            this.pnlKiSq.Controls.Add(this.Label311);
            this.pnlKiSq.Controls.Add(this.Label290);
            this.pnlKiSq.Controls.Add(this.Label162);
            this.pnlKiSq.Controls.Add(this.chkKiSqBase);
            this.pnlKiSq.Controls.Add(this.lblKiSqMiBaseCnt);
            this.pnlKiSq.Controls.Add(this.Label289);
            this.pnlKiSq.Controls.Add(this.Label288);
            this.pnlKiSq.Controls.Add(this.lblKiSqAzBaseCnt);
            this.pnlKiSq.Controls.Add(this.Label160);
            this.pnlKiSq.Location = new System.Drawing.Point(286, 188);
            this.pnlKiSq.Name = "pnlKiSq";
            this.pnlKiSq.Size = new System.Drawing.Size(265, 100);
            this.pnlKiSq.TabIndex = 33;
            // 
            // Label308
            // 
            this.Label308.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label308.Location = new System.Drawing.Point(19, 24);
            this.Label308.Name = "Label308";
            this.Label308.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label308.Size = new System.Drawing.Size(92, 16);
            this.Label308.TabIndex = 42;
            this.Label308.Text = "( 参考件数 = ";
            this.Label308.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label308.Visible = false;
            // 
            // lblKiSqOwKojoBaseCnt
            // 
            this.lblKiSqOwKojoBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiSqOwKojoBaseCnt.Location = new System.Drawing.Point(106, 24);
            this.lblKiSqOwKojoBaseCnt.Name = "lblKiSqOwKojoBaseCnt";
            this.lblKiSqOwKojoBaseCnt.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblKiSqOwKojoBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiSqOwKojoBaseCnt.TabIndex = 43;
            this.lblKiSqOwKojoBaseCnt.Text = "9,999,999";
            this.lblKiSqOwKojoBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblKiSqOwKojoBaseCnt.Visible = false;
            // 
            // Label311
            // 
            this.Label311.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label311.Location = new System.Drawing.Point(195, 24);
            this.Label311.Name = "Label311";
            this.Label311.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label311.Size = new System.Drawing.Size(22, 16);
            this.Label311.TabIndex = 44;
            this.Label311.Text = " )";
            this.Label311.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Label311.Visible = false;
            // 
            // Label290
            // 
            this.Label290.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label290.Location = new System.Drawing.Point(19, 54);
            this.Label290.Name = "Label290";
            this.Label290.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label290.Size = new System.Drawing.Size(135, 16);
            this.Label290.TabIndex = 26;
            this.Label290.Text = "( 未収滞納 参考件数 = ";
            this.Label290.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label162
            // 
            this.Label162.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label162.Location = new System.Drawing.Point(19, 74);
            this.Label162.Name = "Label162";
            this.Label162.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label162.Size = new System.Drawing.Size(135, 16);
            this.Label162.TabIndex = 39;
            this.Label162.Text = "( 預り金 参考件数 = ";
            this.Label162.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkKiSqBase
            // 
            this.chkKiSqBase.AutoSize = true;
            this.chkKiSqBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiSqBase.Location = new System.Drawing.Point(3, 4);
            this.chkKiSqBase.Name = "chkKiSqBase";
            this.chkKiSqBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiSqBase.TabIndex = 24;
            this.chkKiSqBase.Text = "請求情報";
            this.chkKiSqBase.UseVisualStyleBackColor = true;
            this.chkKiSqBase.CheckedChanged += new System.EventHandler(this.chkKiSqBase_CheckedChanged);
            // 
            // lblKiSqMiBaseCnt
            // 
            this.lblKiSqMiBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiSqMiBaseCnt.Location = new System.Drawing.Point(142, 54);
            this.lblKiSqMiBaseCnt.Name = "lblKiSqMiBaseCnt";
            this.lblKiSqMiBaseCnt.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblKiSqMiBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiSqMiBaseCnt.TabIndex = 27;
            this.lblKiSqMiBaseCnt.Text = "9,999,999";
            this.lblKiSqMiBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label289
            // 
            this.Label289.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label289.Location = new System.Drawing.Point(231, 54);
            this.Label289.Name = "Label289";
            this.Label289.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label289.Size = new System.Drawing.Size(22, 16);
            this.Label289.TabIndex = 28;
            this.Label289.Text = " )";
            this.Label289.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label288
            // 
            this.Label288.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label288.Location = new System.Drawing.Point(19, 33);
            this.Label288.Name = "Label288";
            this.Label288.Size = new System.Drawing.Size(240, 20);
            this.Label288.TabIndex = 25;
            this.Label288.Text = "未収滞納, 預り金, その他請求, 変動, 控除";
            this.Label288.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblKiSqAzBaseCnt
            // 
            this.lblKiSqAzBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiSqAzBaseCnt.Location = new System.Drawing.Point(142, 74);
            this.lblKiSqAzBaseCnt.Name = "lblKiSqAzBaseCnt";
            this.lblKiSqAzBaseCnt.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblKiSqAzBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiSqAzBaseCnt.TabIndex = 40;
            this.lblKiSqAzBaseCnt.Text = "9,999,999";
            this.lblKiSqAzBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label160
            // 
            this.Label160.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label160.Location = new System.Drawing.Point(231, 74);
            this.Label160.Name = "Label160";
            this.Label160.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label160.Size = new System.Drawing.Size(22, 16);
            this.Label160.TabIndex = 41;
            this.Label160.Text = " )";
            this.Label160.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlSzen
            // 
            this.pnlSzen.Controls.Add(this.chkKiSzenBase);
            this.pnlSzen.Controls.Add(this.lblKiSzenBaseCnt);
            this.pnlSzen.Controls.Add(this.Label295);
            this.pnlSzen.Controls.Add(this.Label294);
            this.pnlSzen.Controls.Add(this.Label174);
            this.pnlSzen.Location = new System.Drawing.Point(557, 103);
            this.pnlSzen.Name = "pnlSzen";
            this.pnlSzen.Size = new System.Drawing.Size(265, 80);
            this.pnlSzen.TabIndex = 33;
            // 
            // chkKiSzenBase
            // 
            this.chkKiSzenBase.AutoSize = true;
            this.chkKiSzenBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiSzenBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiSzenBase.Name = "chkKiSzenBase";
            this.chkKiSzenBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiSzenBase.TabIndex = 33;
            this.chkKiSzenBase.Text = "修繕情報";
            this.chkKiSzenBase.UseVisualStyleBackColor = true;
            this.chkKiSzenBase.CheckedChanged += new System.EventHandler(this.chkKiSzenBase_CheckedChanged);
            // 
            // lblKiSzenBaseCnt
            // 
            this.lblKiSzenBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiSzenBaseCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiSzenBaseCnt.Name = "lblKiSzenBaseCnt";
            this.lblKiSzenBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiSzenBaseCnt.TabIndex = 36;
            this.lblKiSzenBaseCnt.Text = "9,999,999";
            this.lblKiSzenBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label295
            // 
            this.Label295.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label295.Location = new System.Drawing.Point(19, 54);
            this.Label295.Name = "Label295";
            this.Label295.Size = new System.Drawing.Size(92, 16);
            this.Label295.TabIndex = 35;
            this.Label295.Text = "( 参考件数 = ";
            this.Label295.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label294
            // 
            this.Label294.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label294.Location = new System.Drawing.Point(195, 54);
            this.Label294.Name = "Label294";
            this.Label294.Size = new System.Drawing.Size(22, 16);
            this.Label294.TabIndex = 37;
            this.Label294.Text = " )";
            this.Label294.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label174
            // 
            this.Label174.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label174.Location = new System.Drawing.Point(19, 33);
            this.Label174.Name = "Label174";
            this.Label174.Size = new System.Drawing.Size(198, 16);
            this.Label174.TabIndex = 34;
            this.Label174.Text = "原状回復, リフォーム";
            this.Label174.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiClaim
            // 
            this.pnlKiClaim.Controls.Add(this.chkKiClaimBase);
            this.pnlKiClaim.Controls.Add(this.lblKiClaimBaseCnt);
            this.pnlKiClaim.Controls.Add(this.Label292);
            this.pnlKiClaim.Controls.Add(this.Label180);
            this.pnlKiClaim.Location = new System.Drawing.Point(557, 20);
            this.pnlKiClaim.Name = "pnlKiClaim";
            this.pnlKiClaim.Size = new System.Drawing.Size(265, 60);
            this.pnlKiClaim.TabIndex = 24;
            // 
            // chkKiClaimBase
            // 
            this.chkKiClaimBase.AutoSize = true;
            this.chkKiClaimBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiClaimBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiClaimBase.Name = "chkKiClaimBase";
            this.chkKiClaimBase.Size = new System.Drawing.Size(119, 27);
            this.chkKiClaimBase.TabIndex = 29;
            this.chkKiClaimBase.Text = "クレーム情報";
            this.chkKiClaimBase.UseVisualStyleBackColor = true;
            this.chkKiClaimBase.CheckedChanged += new System.EventHandler(this.chkKiClaimBase_CheckedChanged);
            // 
            // lblKiClaimBaseCnt
            // 
            this.lblKiClaimBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiClaimBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiClaimBaseCnt.Name = "lblKiClaimBaseCnt";
            this.lblKiClaimBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiClaimBaseCnt.TabIndex = 31;
            this.lblKiClaimBaseCnt.Text = "9,999,999";
            this.lblKiClaimBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label292
            // 
            this.Label292.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label292.Location = new System.Drawing.Point(19, 33);
            this.Label292.Name = "Label292";
            this.Label292.Size = new System.Drawing.Size(92, 16);
            this.Label292.TabIndex = 30;
            this.Label292.Text = "( 参考件数 = ";
            this.Label292.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label180
            // 
            this.Label180.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label180.Location = new System.Drawing.Point(195, 33);
            this.Label180.Name = "Label180";
            this.Label180.Size = new System.Drawing.Size(22, 16);
            this.Label180.TabIndex = 32;
            this.Label180.Text = " )";
            this.Label180.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiKy
            // 
            this.pnlKiKy.Controls.Add(this.chkKiKyBase);
            this.pnlKiKy.Controls.Add(this.lblKiKyBaseCnt);
            this.pnlKiKy.Controls.Add(this.Label286);
            this.pnlKiKy.Controls.Add(this.Label285);
            this.pnlKiKy.Controls.Add(this.Label284);
            this.pnlKiKy.Location = new System.Drawing.Point(286, 104);
            this.pnlKiKy.Name = "pnlKiKy";
            this.pnlKiKy.Size = new System.Drawing.Size(265, 80);
            this.pnlKiKy.TabIndex = 42;
            // 
            // chkKiKyBase
            // 
            this.chkKiKyBase.AutoSize = true;
            this.chkKiKyBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiKyBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiKyBase.Name = "chkKiKyBase";
            this.chkKiKyBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiKyBase.TabIndex = 19;
            this.chkKiKyBase.Text = "契約情報";
            this.chkKiKyBase.UseVisualStyleBackColor = true;
            this.chkKiKyBase.CheckedChanged += new System.EventHandler(this.chkKiKyBase_CheckedChanged);
            // 
            // lblKiKyBaseCnt
            // 
            this.lblKiKyBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiKyBaseCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiKyBaseCnt.Name = "lblKiKyBaseCnt";
            this.lblKiKyBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiKyBaseCnt.TabIndex = 22;
            this.lblKiKyBaseCnt.Text = "9,999,999";
            this.lblKiKyBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label286
            // 
            this.Label286.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label286.Location = new System.Drawing.Point(19, 54);
            this.Label286.Name = "Label286";
            this.Label286.Size = new System.Drawing.Size(92, 16);
            this.Label286.TabIndex = 21;
            this.Label286.Text = "( 参考件数 = ";
            this.Label286.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label285
            // 
            this.Label285.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label285.Location = new System.Drawing.Point(195, 54);
            this.Label285.Name = "Label285";
            this.Label285.Size = new System.Drawing.Size(22, 16);
            this.Label285.TabIndex = 23;
            this.Label285.Text = " )";
            this.Label285.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label284
            // 
            this.Label284.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label284.Location = new System.Drawing.Point(19, 33);
            this.Label284.Name = "Label284";
            this.Label284.Size = new System.Drawing.Size(198, 16);
            this.Label284.TabIndex = 20;
            this.Label284.Text = "契約, 更新, 解約, 入居者";
            this.Label284.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiKys
            // 
            this.pnlKiKys.Controls.Add(this.chkKiKysBase);
            this.pnlKiKys.Controls.Add(this.lblKiKysBaseCnt);
            this.pnlKiKys.Controls.Add(this.Label282);
            this.pnlKiKys.Controls.Add(this.Label281);
            this.pnlKiKys.Controls.Add(this.Label280);
            this.pnlKiKys.Location = new System.Drawing.Point(286, 20);
            this.pnlKiKys.Name = "pnlKiKys";
            this.pnlKiKys.Size = new System.Drawing.Size(265, 80);
            this.pnlKiKys.TabIndex = 2;
            // 
            // chkKiKysBase
            // 
            this.chkKiKysBase.AutoSize = true;
            this.chkKiKysBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiKysBase.Location = new System.Drawing.Point(3, 5);
            this.chkKiKysBase.Name = "chkKiKysBase";
            this.chkKiKysBase.Size = new System.Drawing.Size(104, 27);
            this.chkKiKysBase.TabIndex = 13;
            this.chkKiKysBase.Text = "契約者情報";
            this.chkKiKysBase.UseVisualStyleBackColor = true;
            this.chkKiKysBase.CheckedChanged += new System.EventHandler(this.chkKiKysBase_CheckedChanged);
            // 
            // lblKiKysBaseCnt
            // 
            this.lblKiKysBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiKysBaseCnt.Location = new System.Drawing.Point(106, 54);
            this.lblKiKysBaseCnt.Name = "lblKiKysBaseCnt";
            this.lblKiKysBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiKysBaseCnt.TabIndex = 16;
            this.lblKiKysBaseCnt.Text = "9,999,999";
            this.lblKiKysBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label282
            // 
            this.Label282.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label282.Location = new System.Drawing.Point(19, 54);
            this.Label282.Name = "Label282";
            this.Label282.Size = new System.Drawing.Size(92, 16);
            this.Label282.TabIndex = 15;
            this.Label282.Text = "( 参考件数 = ";
            this.Label282.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label281
            // 
            this.Label281.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label281.Location = new System.Drawing.Point(195, 54);
            this.Label281.Name = "Label281";
            this.Label281.Size = new System.Drawing.Size(22, 16);
            this.Label281.TabIndex = 17;
            this.Label281.Text = " )";
            this.Label281.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label280
            // 
            this.Label280.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label280.Location = new System.Drawing.Point(19, 33);
            this.Label280.Name = "Label280";
            this.Label280.Size = new System.Drawing.Size(198, 16);
            this.Label280.TabIndex = 14;
            this.Label280.Text = "契約者・保証人, 口座";
            this.Label280.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlKiHy
            // 
            this.pnlKiHy.Controls.Add(this.chkKiHyBase);
            this.pnlKiHy.Controls.Add(this.lblKiHyBaseCnt);
            this.pnlKiHy.Controls.Add(this.Label275);
            this.pnlKiHy.Controls.Add(this.Label270);
            this.pnlKiHy.Controls.Add(this.chkKiHySetubi);
            this.pnlKiHy.Controls.Add(this.lblKiHySetubiCnt);
            this.pnlKiHy.Controls.Add(this.Label278);
            this.pnlKiHy.Controls.Add(this.Label277);
            this.pnlKiHy.Location = new System.Drawing.Point(15, 84);
            this.pnlKiHy.Name = "pnlKiHy";
            this.pnlKiHy.Size = new System.Drawing.Size(265, 115);
            this.pnlKiHy.TabIndex = 2;
            // 
            // chkKiHyBase
            // 
            this.chkKiHyBase.AutoSize = true;
            this.chkKiHyBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiHyBase.Location = new System.Drawing.Point(3, 4);
            this.chkKiHyBase.Name = "chkKiHyBase";
            this.chkKiHyBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiHyBase.TabIndex = 4;
            this.chkKiHyBase.Text = "部屋情報";
            this.chkKiHyBase.UseVisualStyleBackColor = true;
            this.chkKiHyBase.CheckedChanged += new System.EventHandler(this.chkKiHyBase_CheckedChanged);
            // 
            // lblKiHyBaseCnt
            // 
            this.lblKiHyBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiHyBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiHyBaseCnt.Name = "lblKiHyBaseCnt";
            this.lblKiHyBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiHyBaseCnt.TabIndex = 6;
            this.lblKiHyBaseCnt.Text = "9,999,999";
            this.lblKiHyBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label275
            // 
            this.Label275.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label275.Location = new System.Drawing.Point(19, 33);
            this.Label275.Name = "Label275";
            this.Label275.Size = new System.Drawing.Size(92, 16);
            this.Label275.TabIndex = 5;
            this.Label275.Text = "( 参考件数 = ";
            this.Label275.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label270
            // 
            this.Label270.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label270.Location = new System.Drawing.Point(195, 33);
            this.Label270.Name = "Label270";
            this.Label270.Size = new System.Drawing.Size(22, 16);
            this.Label270.TabIndex = 7;
            this.Label270.Text = " )";
            this.Label270.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // chkKiHySetubi
            // 
            this.chkKiHySetubi.AutoSize = true;
            this.chkKiHySetubi.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiHySetubi.Location = new System.Drawing.Point(18, 56);
            this.chkKiHySetubi.Name = "chkKiHySetubi";
            this.chkKiHySetubi.Size = new System.Drawing.Size(119, 27);
            this.chkKiHySetubi.TabIndex = 8;
            this.chkKiHySetubi.Text = "部屋設備情報";
            this.chkKiHySetubi.UseVisualStyleBackColor = true;
            this.chkKiHySetubi.CheckedChanged += new System.EventHandler(this.chkKiHySetubi_CheckedChanged);
            // 
            // lblKiHySetubiCnt
            // 
            this.lblKiHySetubiCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiHySetubiCnt.Location = new System.Drawing.Point(118, 86);
            this.lblKiHySetubiCnt.Name = "lblKiHySetubiCnt";
            this.lblKiHySetubiCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiHySetubiCnt.TabIndex = 10;
            this.lblKiHySetubiCnt.Text = "9,999,999";
            this.lblKiHySetubiCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label278
            // 
            this.Label278.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label278.Location = new System.Drawing.Point(34, 86);
            this.Label278.Name = "Label278";
            this.Label278.Size = new System.Drawing.Size(92, 16);
            this.Label278.TabIndex = 9;
            this.Label278.Text = "( 参考件数 = ";
            this.Label278.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label277
            // 
            this.Label277.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label277.Location = new System.Drawing.Point(207, 86);
            this.Label277.Name = "Label277";
            this.Label277.Size = new System.Drawing.Size(22, 16);
            this.Label277.TabIndex = 11;
            this.Label277.Text = " )";
            this.Label277.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlKiBk
            // 
            this.pnlKiBk.Controls.Add(this.chkKiBkBase);
            this.pnlKiBk.Controls.Add(this.lblKiBkBaseCnt);
            this.pnlKiBk.Controls.Add(this.Label273);
            this.pnlKiBk.Controls.Add(this.Label271);
            this.pnlKiBk.Location = new System.Drawing.Point(15, 20);
            this.pnlKiBk.Name = "pnlKiBk";
            this.pnlKiBk.Size = new System.Drawing.Size(265, 60);
            this.pnlKiBk.TabIndex = 1;
            // 
            // chkKiBkBase
            // 
            this.chkKiBkBase.AutoSize = true;
            this.chkKiBkBase.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKiBkBase.Location = new System.Drawing.Point(3, 3);
            this.chkKiBkBase.Name = "chkKiBkBase";
            this.chkKiBkBase.Size = new System.Drawing.Size(89, 27);
            this.chkKiBkBase.TabIndex = 0;
            this.chkKiBkBase.Text = "物件情報";
            this.chkKiBkBase.UseVisualStyleBackColor = true;
            this.chkKiBkBase.CheckedChanged += new System.EventHandler(this.chkKiBkBase_CheckedChanged);
            // 
            // lblKiBkBaseCnt
            // 
            this.lblKiBkBaseCnt.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblKiBkBaseCnt.Location = new System.Drawing.Point(106, 33);
            this.lblKiBkBaseCnt.Name = "lblKiBkBaseCnt";
            this.lblKiBkBaseCnt.Size = new System.Drawing.Size(83, 16);
            this.lblKiBkBaseCnt.TabIndex = 2;
            this.lblKiBkBaseCnt.Text = "9,999,999";
            this.lblKiBkBaseCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label273
            // 
            this.Label273.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label273.Location = new System.Drawing.Point(19, 33);
            this.Label273.Name = "Label273";
            this.Label273.Size = new System.Drawing.Size(92, 16);
            this.Label273.TabIndex = 1;
            this.Label273.Text = "( 参考件数 = ";
            this.Label273.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label271
            // 
            this.Label271.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label271.Location = new System.Drawing.Point(195, 33);
            this.Label271.Name = "Label271";
            this.Label271.Size = new System.Drawing.Size(22, 16);
            this.Label271.TabIndex = 3;
            this.Label271.Text = " )";
            this.Label271.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // grpKizonKagi
            // 
            this.grpKizonKagi.Controls.Add(this.optKiKyKagi);
            this.grpKizonKagi.Controls.Add(this.optKiHyKagi);
            this.grpKizonKagi.Controls.Add(this.Label86);
            this.grpKizonKagi.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpKizonKagi.Location = new System.Drawing.Point(18, 205);
            this.grpKizonKagi.Name = "grpKizonKagi";
            this.grpKizonKagi.Size = new System.Drawing.Size(262, 87);
            this.grpKizonKagi.TabIndex = 12;
            this.grpKizonKagi.TabStop = false;
            this.grpKizonKagi.Text = "鍵情報選択";
            // 
            // optKiKyKagi
            // 
            this.optKiKyKagi.AutoSize = true;
            this.optKiKyKagi.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optKiKyKagi.Location = new System.Drawing.Point(150, 54);
            this.optKiKyKagi.Name = "optKiKyKagi";
            this.optKiKyKagi.Size = new System.Drawing.Size(103, 27);
            this.optKiKyKagi.TabIndex = 2;
            this.optKiKyKagi.TabStop = true;
            this.optKiKyKagi.Text = "契約鍵情報";
            this.optKiKyKagi.UseVisualStyleBackColor = true;
            this.optKiKyKagi.CheckedChanged += new System.EventHandler(this.optKiKagi_CheckedChanged);
            // 
            // optKiHyKagi
            // 
            this.optKiHyKagi.AutoSize = true;
            this.optKiHyKagi.Checked = true;
            this.optKiHyKagi.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optKiHyKagi.Location = new System.Drawing.Point(20, 54);
            this.optKiHyKagi.Name = "optKiHyKagi";
            this.optKiHyKagi.Size = new System.Drawing.Size(103, 27);
            this.optKiHyKagi.TabIndex = 1;
            this.optKiHyKagi.TabStop = true;
            this.optKiHyKagi.Text = "部屋鍵情報";
            this.optKiHyKagi.UseVisualStyleBackColor = true;
            this.optKiHyKagi.CheckedChanged += new System.EventHandler(this.optKiKagi_CheckedChanged);
            // 
            // Label86
            // 
            this.Label86.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label86.Location = new System.Drawing.Point(12, 20);
            this.Label86.Name = "Label86";
            this.Label86.Size = new System.Drawing.Size(244, 35);
            this.Label86.TabIndex = 0;
            this.Label86.Text = "V7では部屋鍵と契約鍵が存在します。どちらを移行対象とするか選択して下さい。";
            // 
            // Label272
            // 
            this.Label272.AutoSize = true;
            this.Label272.Location = new System.Drawing.Point(806, 306);
            this.Label272.Name = "Label272";
            this.Label272.Size = new System.Drawing.Size(35, 18);
            this.Label272.TabIndex = 38;
            this.Label272.Text = "3 / 3";
            // 
            // tabPageBase110
            // 
            this.tabPageBase110.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase110.Controls.Add(this.lblRelItemInfo);
            this.tabPageBase110.Controls.Add(this.grpMst);
            this.tabPageBase110.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase110.Name = "tabPageBase110";
            this.tabPageBase110.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase110.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase110.TabIndex = 0;
            this.tabPageBase110.Text = "各マスタ情報";
            // 
            // lblRelItemInfo
            // 
            this.lblRelItemInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblRelItemInfo.Location = new System.Drawing.Point(439, 20);
            this.lblRelItemInfo.Name = "lblRelItemInfo";
            this.lblRelItemInfo.Size = new System.Drawing.Size(377, 235);
            this.lblRelItemInfo.TabIndex = 1;
            this.lblRelItemInfo.Text = "\r\n  以下は賃貸革命10のマスター情報と紐付けを行う項目です。\r\n  紐付設定画面にて、ユーザーデータと革命10のマスターデータの\r\n  紐付けを行って下さい。" +
    "\r\n\r\n    ・ 物件分類マスタ  ・ 取引態様マスタ  ・ 部屋分類マスタ\r\n    ・ 口座種別マスタ  ・ 構造マスタ        ・ 入金区分マスタ" +
    "\r\n    ・ 設備マスタ        ・ 入金項目マスタ\r\n";
            // 
            // grpMst
            // 
            this.grpMst.Controls.Add(this.chkMstGazotitle);
            this.grpMst.Controls.Add(this.chkMstBikolst);
            this.grpMst.Controls.Add(this.chkMstBikotitle);
            this.grpMst.Controls.Add(this.chkMstHendoitiran);
            this.grpMst.Controls.Add(this.chkMstHendo);
            this.grpMst.Controls.Add(this.chkMstKasyoClaimrui);
            this.grpMst.Controls.Add(this.chkMstBus);
            this.grpMst.Controls.Add(this.chkMstSchool);
            this.grpMst.Controls.Add(this.chkMstTokuyaku);
            this.grpMst.Controls.Add(this.chkMstHokenrui);
            this.grpMst.Controls.Add(this.chkMstBusKotu);
            this.grpMst.Controls.Add(this.chkMstKagititle);
            this.grpMst.Controls.Add(this.chkMstArea);
            this.grpMst.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpMst.Location = new System.Drawing.Point(10, 5);
            this.grpMst.Name = "grpMst";
            this.grpMst.Size = new System.Drawing.Size(397, 265);
            this.grpMst.TabIndex = 0;
            this.grpMst.TabStop = false;
            this.grpMst.Text = " 各マスタ情報 ";
            // 
            // chkMstGazotitle
            // 
            this.chkMstGazotitle.AutoSize = true;
            this.chkMstGazotitle.Location = new System.Drawing.Point(200, 192);
            this.chkMstGazotitle.Name = "chkMstGazotitle";
            this.chkMstGazotitle.Size = new System.Drawing.Size(135, 22);
            this.chkMstGazotitle.TabIndex = 13;
            this.chkMstGazotitle.Text = "画像タイトルマスタ";
            this.chkMstGazotitle.UseVisualStyleBackColor = true;
            // 
            // chkMstBikolst
            // 
            this.chkMstBikolst.AutoSize = true;
            this.chkMstBikolst.Location = new System.Drawing.Point(216, 165);
            this.chkMstBikolst.Name = "chkMstBikolst";
            this.chkMstBikolst.Size = new System.Drawing.Size(171, 22);
            this.chkMstBikolst.TabIndex = 12;
            this.chkMstBikolst.Text = "備考入力補助リストマスタ";
            this.chkMstBikolst.UseVisualStyleBackColor = true;
            // 
            // chkMstBikotitle
            // 
            this.chkMstBikotitle.AutoSize = true;
            this.chkMstBikotitle.Location = new System.Drawing.Point(200, 136);
            this.chkMstBikotitle.Name = "chkMstBikotitle";
            this.chkMstBikotitle.Size = new System.Drawing.Size(135, 22);
            this.chkMstBikotitle.TabIndex = 11;
            this.chkMstBikotitle.Text = "備考タイトルマスタ";
            this.chkMstBikotitle.UseVisualStyleBackColor = true;
            // 
            // chkMstHendoitiran
            // 
            this.chkMstHendoitiran.AutoSize = true;
            this.chkMstHendoitiran.Location = new System.Drawing.Point(215, 109);
            this.chkMstHendoitiran.Name = "chkMstHendoitiran";
            this.chkMstHendoitiran.Size = new System.Drawing.Size(123, 22);
            this.chkMstHendoitiran.TabIndex = 10;
            this.chkMstHendoitiran.Text = "変動費料金単価表";
            this.chkMstHendoitiran.UseVisualStyleBackColor = true;
            // 
            // chkMstHendo
            // 
            this.chkMstHendo.AutoSize = true;
            this.chkMstHendo.Location = new System.Drawing.Point(201, 81);
            this.chkMstHendo.Name = "chkMstHendo";
            this.chkMstHendo.Size = new System.Drawing.Size(111, 22);
            this.chkMstHendo.TabIndex = 9;
            this.chkMstHendo.Text = "変動費設定内容";
            this.chkMstHendo.UseVisualStyleBackColor = true;
            // 
            // chkMstKasyoClaimrui
            // 
            this.chkMstKasyoClaimrui.AutoSize = true;
            this.chkMstKasyoClaimrui.Location = new System.Drawing.Point(200, 53);
            this.chkMstKasyoClaimrui.Name = "chkMstKasyoClaimrui";
            this.chkMstKasyoClaimrui.Size = new System.Drawing.Size(147, 22);
            this.chkMstKasyoClaimrui.TabIndex = 8;
            this.chkMstKasyoClaimrui.Text = "クレーム分類設定内容";
            this.chkMstKasyoClaimrui.UseVisualStyleBackColor = true;
            // 
            // chkMstBus
            // 
            this.chkMstBus.AutoSize = true;
            this.chkMstBus.Location = new System.Drawing.Point(25, 25);
            this.chkMstBus.Name = "chkMstBus";
            this.chkMstBus.Size = new System.Drawing.Size(111, 22);
            this.chkMstBus.TabIndex = 0;
            this.chkMstBus.Text = "バス交通マスタ";
            this.chkMstBus.UseVisualStyleBackColor = true;
            // 
            // chkMstSchool
            // 
            this.chkMstSchool.AutoSize = true;
            this.chkMstSchool.Location = new System.Drawing.Point(25, 81);
            this.chkMstSchool.Name = "chkMstSchool";
            this.chkMstSchool.Size = new System.Drawing.Size(99, 22);
            this.chkMstSchool.TabIndex = 2;
            this.chkMstSchool.Text = "学校区マスタ";
            this.chkMstSchool.UseVisualStyleBackColor = true;
            // 
            // chkMstTokuyaku
            // 
            this.chkMstTokuyaku.AutoSize = true;
            this.chkMstTokuyaku.Location = new System.Drawing.Point(200, 25);
            this.chkMstTokuyaku.Name = "chkMstTokuyaku";
            this.chkMstTokuyaku.Size = new System.Drawing.Size(87, 22);
            this.chkMstTokuyaku.TabIndex = 7;
            this.chkMstTokuyaku.Text = "特約マスタ";
            this.chkMstTokuyaku.UseVisualStyleBackColor = true;
            // 
            // chkMstHokenrui
            // 
            this.chkMstHokenrui.AutoSize = true;
            this.chkMstHokenrui.Location = new System.Drawing.Point(25, 137);
            this.chkMstHokenrui.Name = "chkMstHokenrui";
            this.chkMstHokenrui.Size = new System.Drawing.Size(111, 22);
            this.chkMstHokenrui.TabIndex = 4;
            this.chkMstHokenrui.Text = "保険種類マスタ";
            this.chkMstHokenrui.UseVisualStyleBackColor = true;
            // 
            // chkMstBusKotu
            // 
            this.chkMstBusKotu.AutoSize = true;
            this.chkMstBusKotu.Location = new System.Drawing.Point(45, 53);
            this.chkMstBusKotu.Name = "chkMstBusKotu";
            this.chkMstBusKotu.Size = new System.Drawing.Size(99, 22);
            this.chkMstBusKotu.TabIndex = 1;
            this.chkMstBusKotu.Text = "バス停マスタ";
            this.chkMstBusKotu.UseVisualStyleBackColor = true;
            // 
            // chkMstKagititle
            // 
            this.chkMstKagititle.AutoSize = true;
            this.chkMstKagititle.Location = new System.Drawing.Point(25, 168);
            this.chkMstKagititle.Name = "chkMstKagititle";
            this.chkMstKagititle.Size = new System.Drawing.Size(123, 22);
            this.chkMstKagititle.TabIndex = 6;
            this.chkMstKagititle.Text = "鍵タイトルマスタ";
            this.chkMstKagititle.UseVisualStyleBackColor = true;
            // 
            // chkMstArea
            // 
            this.chkMstArea.AutoSize = true;
            this.chkMstArea.Location = new System.Drawing.Point(25, 109);
            this.chkMstArea.Name = "chkMstArea";
            this.chkMstArea.Size = new System.Drawing.Size(99, 22);
            this.chkMstArea.TabIndex = 3;
            this.chkMstArea.Text = "エリアマスタ";
            this.chkMstArea.UseVisualStyleBackColor = true;
            // 
            // tabPageBase120
            // 
            this.tabPageBase120.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase120.Controls.Add(this.grpGy);
            this.tabPageBase120.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase120.Name = "tabPageBase120";
            this.tabPageBase120.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase120.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase120.TabIndex = 1;
            this.tabPageBase120.Text = "業者情報";
            // 
            // grpGy
            // 
            this.grpGy.Controls.Add(this.chkGySyuzenMemo);
            this.grpGy.Controls.Add(this.chkGyHokenMemo);
            this.grpGy.Controls.Add(this.chkGySyuzenKoza);
            this.grpGy.Controls.Add(this.chkGySekoBase);
            this.grpGy.Controls.Add(this.chkGySisetuBase);
            this.grpGy.Controls.Add(this.chkGySyuzenBase);
            this.grpGy.Controls.Add(this.chkGyYatinhosyoMemo);
            this.grpGy.Controls.Add(this.chkGyYatinhosyoBase);
            this.grpGy.Controls.Add(this.chkGyCyukaiKoza);
            this.grpGy.Controls.Add(this.chkGyCyukaiMemo);
            this.grpGy.Controls.Add(this.chkGyHokenBase);
            this.grpGy.Controls.Add(this.chkGyLifelineBase);
            this.grpGy.Controls.Add(this.chkGyHokenKoza);
            this.grpGy.Controls.Add(this.chkGyCyukaiBase);
            this.grpGy.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpGy.Location = new System.Drawing.Point(10, 5);
            this.grpGy.Name = "grpGy";
            this.grpGy.Size = new System.Drawing.Size(588, 265);
            this.grpGy.TabIndex = 0;
            this.grpGy.TabStop = false;
            this.grpGy.Text = " 業者情報 ";
            // 
            // chkGySyuzenMemo
            // 
            this.chkGySyuzenMemo.AutoSize = true;
            this.chkGySyuzenMemo.Location = new System.Drawing.Point(245, 80);
            this.chkGySyuzenMemo.Name = "chkGySyuzenMemo";
            this.chkGySyuzenMemo.Size = new System.Drawing.Size(123, 22);
            this.chkGySyuzenMemo.TabIndex = 9;
            this.chkGySyuzenMemo.Text = "修繕業者メモ情報";
            this.chkGySyuzenMemo.UseVisualStyleBackColor = true;
            // 
            // chkGyHokenMemo
            // 
            this.chkGyHokenMemo.AutoSize = true;
            this.chkGyHokenMemo.Location = new System.Drawing.Point(45, 164);
            this.chkGyHokenMemo.Name = "chkGyHokenMemo";
            this.chkGyHokenMemo.Size = new System.Drawing.Size(123, 22);
            this.chkGyHokenMemo.TabIndex = 5;
            this.chkGyHokenMemo.Text = "保険業者メモ情報";
            this.chkGyHokenMemo.UseVisualStyleBackColor = true;
            // 
            // chkGySyuzenKoza
            // 
            this.chkGySyuzenKoza.AutoSize = true;
            this.chkGySyuzenKoza.Location = new System.Drawing.Point(245, 52);
            this.chkGySyuzenKoza.Name = "chkGySyuzenKoza";
            this.chkGySyuzenKoza.Size = new System.Drawing.Size(123, 22);
            this.chkGySyuzenKoza.TabIndex = 8;
            this.chkGySyuzenKoza.Text = "修繕業者口座情報";
            this.chkGySyuzenKoza.UseVisualStyleBackColor = true;
            // 
            // chkGySekoBase
            // 
            this.chkGySekoBase.AutoSize = true;
            this.chkGySekoBase.Location = new System.Drawing.Point(425, 52);
            this.chkGySekoBase.Name = "chkGySekoBase";
            this.chkGySekoBase.Size = new System.Drawing.Size(99, 22);
            this.chkGySekoBase.TabIndex = 13;
            this.chkGySekoBase.Text = "施工業者情報";
            this.chkGySekoBase.UseVisualStyleBackColor = true;
            // 
            // chkGySisetuBase
            // 
            this.chkGySisetuBase.AutoSize = true;
            this.chkGySisetuBase.Location = new System.Drawing.Point(425, 25);
            this.chkGySisetuBase.Name = "chkGySisetuBase";
            this.chkGySisetuBase.Size = new System.Drawing.Size(123, 22);
            this.chkGySisetuBase.TabIndex = 12;
            this.chkGySisetuBase.Text = "施設保守業者情報";
            this.chkGySisetuBase.UseVisualStyleBackColor = true;
            // 
            // chkGySyuzenBase
            // 
            this.chkGySyuzenBase.AutoSize = true;
            this.chkGySyuzenBase.Location = new System.Drawing.Point(225, 25);
            this.chkGySyuzenBase.Name = "chkGySyuzenBase";
            this.chkGySyuzenBase.Size = new System.Drawing.Size(123, 22);
            this.chkGySyuzenBase.TabIndex = 7;
            this.chkGySyuzenBase.Text = "修繕業者基本情報";
            this.chkGySyuzenBase.UseVisualStyleBackColor = true;
            // 
            // chkGyYatinhosyoMemo
            // 
            this.chkGyYatinhosyoMemo.AutoSize = true;
            this.chkGyYatinhosyoMemo.Location = new System.Drawing.Point(245, 136);
            this.chkGyYatinhosyoMemo.Name = "chkGyYatinhosyoMemo";
            this.chkGyYatinhosyoMemo.Size = new System.Drawing.Size(147, 22);
            this.chkGyYatinhosyoMemo.TabIndex = 11;
            this.chkGyYatinhosyoMemo.Text = "家賃保証業者メモ情報";
            this.chkGyYatinhosyoMemo.UseVisualStyleBackColor = true;
            // 
            // chkGyYatinhosyoBase
            // 
            this.chkGyYatinhosyoBase.AutoSize = true;
            this.chkGyYatinhosyoBase.Location = new System.Drawing.Point(225, 108);
            this.chkGyYatinhosyoBase.Name = "chkGyYatinhosyoBase";
            this.chkGyYatinhosyoBase.Size = new System.Drawing.Size(147, 22);
            this.chkGyYatinhosyoBase.TabIndex = 10;
            this.chkGyYatinhosyoBase.Text = "家賃保証業者基本情報";
            this.chkGyYatinhosyoBase.UseVisualStyleBackColor = true;
            // 
            // chkGyCyukaiKoza
            // 
            this.chkGyCyukaiKoza.AutoSize = true;
            this.chkGyCyukaiKoza.Location = new System.Drawing.Point(45, 52);
            this.chkGyCyukaiKoza.Name = "chkGyCyukaiKoza";
            this.chkGyCyukaiKoza.Size = new System.Drawing.Size(123, 22);
            this.chkGyCyukaiKoza.TabIndex = 1;
            this.chkGyCyukaiKoza.Text = "仲介業者口座情報";
            this.chkGyCyukaiKoza.UseVisualStyleBackColor = true;
            // 
            // chkGyCyukaiMemo
            // 
            this.chkGyCyukaiMemo.AutoSize = true;
            this.chkGyCyukaiMemo.Location = new System.Drawing.Point(45, 80);
            this.chkGyCyukaiMemo.Name = "chkGyCyukaiMemo";
            this.chkGyCyukaiMemo.Size = new System.Drawing.Size(123, 22);
            this.chkGyCyukaiMemo.TabIndex = 2;
            this.chkGyCyukaiMemo.Text = "仲介業者メモ情報";
            this.chkGyCyukaiMemo.UseVisualStyleBackColor = true;
            // 
            // chkGyHokenBase
            // 
            this.chkGyHokenBase.AutoSize = true;
            this.chkGyHokenBase.Location = new System.Drawing.Point(25, 108);
            this.chkGyHokenBase.Name = "chkGyHokenBase";
            this.chkGyHokenBase.Size = new System.Drawing.Size(123, 22);
            this.chkGyHokenBase.TabIndex = 3;
            this.chkGyHokenBase.Text = "保険業者基本情報";
            this.chkGyHokenBase.UseVisualStyleBackColor = true;
            // 
            // chkGyLifelineBase
            // 
            this.chkGyLifelineBase.AutoSize = true;
            this.chkGyLifelineBase.Location = new System.Drawing.Point(25, 192);
            this.chkGyLifelineBase.Name = "chkGyLifelineBase";
            this.chkGyLifelineBase.Size = new System.Drawing.Size(147, 22);
            this.chkGyLifelineBase.TabIndex = 6;
            this.chkGyLifelineBase.Text = "ライフライン業者情報";
            this.chkGyLifelineBase.UseVisualStyleBackColor = true;
            // 
            // chkGyHokenKoza
            // 
            this.chkGyHokenKoza.AutoSize = true;
            this.chkGyHokenKoza.Location = new System.Drawing.Point(45, 136);
            this.chkGyHokenKoza.Name = "chkGyHokenKoza";
            this.chkGyHokenKoza.Size = new System.Drawing.Size(123, 22);
            this.chkGyHokenKoza.TabIndex = 4;
            this.chkGyHokenKoza.Text = "保険業者口座情報";
            this.chkGyHokenKoza.UseVisualStyleBackColor = true;
            // 
            // chkGyCyukaiBase
            // 
            this.chkGyCyukaiBase.AutoSize = true;
            this.chkGyCyukaiBase.Location = new System.Drawing.Point(25, 24);
            this.chkGyCyukaiBase.Name = "chkGyCyukaiBase";
            this.chkGyCyukaiBase.Size = new System.Drawing.Size(123, 22);
            this.chkGyCyukaiBase.TabIndex = 0;
            this.chkGyCyukaiBase.Text = "仲介業者基本情報";
            this.chkGyCyukaiBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase130
            // 
            this.tabPageBase130.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase130.Controls.Add(this.grpKys);
            this.tabPageBase130.Controls.Add(this.grpOw);
            this.tabPageBase130.Controls.Add(this.grpJisya);
            this.tabPageBase130.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase130.Name = "tabPageBase130";
            this.tabPageBase130.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase130.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase130.TabIndex = 2;
            this.tabPageBase130.Text = "各基本情報";
            // 
            // grpKys
            // 
            this.grpKys.Controls.Add(this.chkKysHosyonin);
            this.grpKys.Controls.Add(this.chkKysSyogoKana);
            this.grpKys.Controls.Add(this.chkKysKoza);
            this.grpKys.Controls.Add(this.chkKysMemo);
            this.grpKys.Controls.Add(this.chkKysBase);
            this.grpKys.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpKys.Location = new System.Drawing.Point(638, 6);
            this.grpKys.Name = "grpKys";
            this.grpKys.Size = new System.Drawing.Size(200, 265);
            this.grpKys.TabIndex = 2;
            this.grpKys.TabStop = false;
            this.grpKys.Text = "契約者情報";
            // 
            // chkKysHosyonin
            // 
            this.chkKysHosyonin.AutoSize = true;
            this.chkKysHosyonin.Location = new System.Drawing.Point(45, 108);
            this.chkKysHosyonin.Name = "chkKysHosyonin";
            this.chkKysHosyonin.Size = new System.Drawing.Size(123, 22);
            this.chkKysHosyonin.TabIndex = 3;
            this.chkKysHosyonin.Text = "契約者保証人情報";
            this.chkKysHosyonin.UseVisualStyleBackColor = true;
            // 
            // chkKysSyogoKana
            // 
            this.chkKysSyogoKana.AutoSize = true;
            this.chkKysSyogoKana.Location = new System.Drawing.Point(45, 80);
            this.chkKysSyogoKana.Name = "chkKysSyogoKana";
            this.chkKysSyogoKana.Size = new System.Drawing.Size(147, 22);
            this.chkKysSyogoKana.TabIndex = 2;
            this.chkKysSyogoKana.Text = "契約者照合用カナ情報";
            this.chkKysSyogoKana.UseVisualStyleBackColor = true;
            // 
            // chkKysKoza
            // 
            this.chkKysKoza.AutoSize = true;
            this.chkKysKoza.Location = new System.Drawing.Point(45, 53);
            this.chkKysKoza.Name = "chkKysKoza";
            this.chkKysKoza.Size = new System.Drawing.Size(111, 22);
            this.chkKysKoza.TabIndex = 1;
            this.chkKysKoza.Text = "契約者口座情報";
            this.chkKysKoza.UseVisualStyleBackColor = true;
            // 
            // chkKysMemo
            // 
            this.chkKysMemo.AutoSize = true;
            this.chkKysMemo.Location = new System.Drawing.Point(45, 136);
            this.chkKysMemo.Name = "chkKysMemo";
            this.chkKysMemo.Size = new System.Drawing.Size(111, 22);
            this.chkKysMemo.TabIndex = 4;
            this.chkKysMemo.Text = "契約者メモ情報";
            this.chkKysMemo.UseVisualStyleBackColor = true;
            // 
            // chkKysBase
            // 
            this.chkKysBase.AutoSize = true;
            this.chkKysBase.Location = new System.Drawing.Point(25, 25);
            this.chkKysBase.Name = "chkKysBase";
            this.chkKysBase.Size = new System.Drawing.Size(111, 22);
            this.chkKysBase.TabIndex = 0;
            this.chkKysBase.Text = "契約者基本情報";
            this.chkKysBase.UseVisualStyleBackColor = true;
            // 
            // grpOw
            // 
            this.grpOw.Controls.Add(this.chkOwEvent);
            this.grpOw.Controls.Add(this.chkOwKoza);
            this.grpOw.Controls.Add(this.chkOwMemo);
            this.grpOw.Controls.Add(this.chkOwBase);
            this.grpOw.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpOw.Location = new System.Drawing.Point(432, 6);
            this.grpOw.Name = "grpOw";
            this.grpOw.Size = new System.Drawing.Size(200, 265);
            this.grpOw.TabIndex = 1;
            this.grpOw.TabStop = false;
            this.grpOw.Text = "家主情報";
            // 
            // chkOwEvent
            // 
            this.chkOwEvent.AutoSize = true;
            this.chkOwEvent.Location = new System.Drawing.Point(45, 80);
            this.chkOwEvent.Name = "chkOwEvent";
            this.chkOwEvent.Size = new System.Drawing.Size(123, 22);
            this.chkOwEvent.TabIndex = 2;
            this.chkOwEvent.Text = "家主イベント情報";
            this.chkOwEvent.UseVisualStyleBackColor = true;
            // 
            // chkOwKoza
            // 
            this.chkOwKoza.AutoSize = true;
            this.chkOwKoza.Location = new System.Drawing.Point(45, 53);
            this.chkOwKoza.Name = "chkOwKoza";
            this.chkOwKoza.Size = new System.Drawing.Size(99, 22);
            this.chkOwKoza.TabIndex = 1;
            this.chkOwKoza.Text = "家主口座情報";
            this.chkOwKoza.UseVisualStyleBackColor = true;
            // 
            // chkOwMemo
            // 
            this.chkOwMemo.AutoSize = true;
            this.chkOwMemo.Location = new System.Drawing.Point(45, 108);
            this.chkOwMemo.Name = "chkOwMemo";
            this.chkOwMemo.Size = new System.Drawing.Size(99, 22);
            this.chkOwMemo.TabIndex = 3;
            this.chkOwMemo.Text = "家主メモ情報";
            this.chkOwMemo.UseVisualStyleBackColor = true;
            // 
            // chkOwBase
            // 
            this.chkOwBase.AutoSize = true;
            this.chkOwBase.Location = new System.Drawing.Point(25, 25);
            this.chkOwBase.Name = "chkOwBase";
            this.chkOwBase.Size = new System.Drawing.Size(99, 22);
            this.chkOwBase.TabIndex = 0;
            this.chkOwBase.Text = "家主基本情報";
            this.chkOwBase.UseVisualStyleBackColor = true;
            // 
            // grpJisya
            // 
            this.grpJisya.Controls.Add(this.chkFBANSERSetuzoku);
            this.grpJisya.Controls.Add(this.chkMstANSERArea);
            this.grpJisya.Controls.Add(this.chkMstANSERAccpoint);
            this.grpJisya.Controls.Add(this.chkJisyaTanto);
            this.grpJisya.Controls.Add(this.chkMstYatinKoza);
            this.grpJisya.Controls.Add(this.chkJisyaKoza);
            this.grpJisya.Controls.Add(this.chkJisyaBase);
            this.grpJisya.Controls.Add(this.chkFBFuriirai);
            this.grpJisya.Controls.Add(this.chkFBNsSyutoku);
            this.grpJisya.Controls.Add(this.chkJisyaMemo);
            this.grpJisya.Controls.Add(this.chkFBKozafurikae);
            this.grpJisya.Controls.Add(this.chkFBFuritesuryo);
            this.grpJisya.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpJisya.Location = new System.Drawing.Point(10, 5);
            this.grpJisya.Name = "grpJisya";
            this.grpJisya.Size = new System.Drawing.Size(398, 266);
            this.grpJisya.TabIndex = 0;
            this.grpJisya.TabStop = false;
            this.grpJisya.Text = " 自社情報 ";
            // 
            // chkFBANSERSetuzoku
            // 
            this.chkFBANSERSetuzoku.AutoSize = true;
            this.chkFBANSERSetuzoku.Location = new System.Drawing.Point(176, 81);
            this.chkFBANSERSetuzoku.Name = "chkFBANSERSetuzoku";
            this.chkFBANSERSetuzoku.Size = new System.Drawing.Size(115, 22);
            this.chkFBANSERSetuzoku.TabIndex = 17;
            this.chkFBANSERSetuzoku.Text = "ANSER接続情報";
            this.chkFBANSERSetuzoku.UseVisualStyleBackColor = true;
            // 
            // chkMstANSERArea
            // 
            this.chkMstANSERArea.AutoSize = true;
            this.chkMstANSERArea.Location = new System.Drawing.Point(176, 53);
            this.chkMstANSERArea.Name = "chkMstANSERArea";
            this.chkMstANSERArea.Size = new System.Drawing.Size(127, 22);
            this.chkMstANSERArea.TabIndex = 16;
            this.chkMstANSERArea.Text = "ANSERエリア情報";
            this.chkMstANSERArea.UseVisualStyleBackColor = true;
            // 
            // chkMstANSERAccpoint
            // 
            this.chkMstANSERAccpoint.AutoSize = true;
            this.chkMstANSERAccpoint.Location = new System.Drawing.Point(176, 25);
            this.chkMstANSERAccpoint.Name = "chkMstANSERAccpoint";
            this.chkMstANSERAccpoint.Size = new System.Drawing.Size(187, 22);
            this.chkMstANSERAccpoint.TabIndex = 15;
            this.chkMstANSERAccpoint.Text = "ANSERアクセスポイント情報";
            this.chkMstANSERAccpoint.UseVisualStyleBackColor = true;
            // 
            // chkJisyaTanto
            // 
            this.chkJisyaTanto.AutoSize = true;
            this.chkJisyaTanto.Location = new System.Drawing.Point(25, 53);
            this.chkJisyaTanto.Name = "chkJisyaTanto";
            this.chkJisyaTanto.Size = new System.Drawing.Size(111, 22);
            this.chkJisyaTanto.TabIndex = 11;
            this.chkJisyaTanto.Text = "自社担当者情報";
            this.chkJisyaTanto.UseVisualStyleBackColor = true;
            // 
            // chkMstYatinKoza
            // 
            this.chkMstYatinKoza.AutoSize = true;
            this.chkMstYatinKoza.Location = new System.Drawing.Point(176, 137);
            this.chkMstYatinKoza.Name = "chkMstYatinKoza";
            this.chkMstYatinKoza.Size = new System.Drawing.Size(123, 22);
            this.chkMstYatinKoza.TabIndex = 10;
            this.chkMstYatinKoza.Text = "家賃入金口座情報";
            this.chkMstYatinKoza.UseVisualStyleBackColor = true;
            // 
            // chkJisyaKoza
            // 
            this.chkJisyaKoza.AutoSize = true;
            this.chkJisyaKoza.Location = new System.Drawing.Point(25, 109);
            this.chkJisyaKoza.Name = "chkJisyaKoza";
            this.chkJisyaKoza.Size = new System.Drawing.Size(99, 22);
            this.chkJisyaKoza.TabIndex = 1;
            this.chkJisyaKoza.Text = "自社口座情報";
            this.chkJisyaKoza.UseVisualStyleBackColor = true;
            // 
            // chkJisyaBase
            // 
            this.chkJisyaBase.AutoSize = true;
            this.chkJisyaBase.Location = new System.Drawing.Point(25, 25);
            this.chkJisyaBase.Name = "chkJisyaBase";
            this.chkJisyaBase.Size = new System.Drawing.Size(99, 22);
            this.chkJisyaBase.TabIndex = 0;
            this.chkJisyaBase.Text = "自社基本情報";
            this.chkJisyaBase.UseVisualStyleBackColor = true;
            // 
            // chkFBFuriirai
            // 
            this.chkFBFuriirai.AutoSize = true;
            this.chkFBFuriirai.Location = new System.Drawing.Point(56, 137);
            this.chkFBFuriirai.Name = "chkFBFuriirai";
            this.chkFBFuriirai.Size = new System.Drawing.Size(111, 22);
            this.chkFBFuriirai.TabIndex = 2;
            this.chkFBFuriirai.Text = "振込依頼人情報";
            this.chkFBFuriirai.UseVisualStyleBackColor = true;
            // 
            // chkFBNsSyutoku
            // 
            this.chkFBNsSyutoku.AutoSize = true;
            this.chkFBNsSyutoku.Location = new System.Drawing.Point(56, 221);
            this.chkFBNsSyutoku.Name = "chkFBNsSyutoku";
            this.chkFBNsSyutoku.Size = new System.Drawing.Size(111, 22);
            this.chkFBNsSyutoku.TabIndex = 9;
            this.chkFBNsSyutoku.Text = "入出金取得情報";
            this.chkFBNsSyutoku.UseVisualStyleBackColor = true;
            // 
            // chkJisyaMemo
            // 
            this.chkJisyaMemo.AutoSize = true;
            this.chkJisyaMemo.Location = new System.Drawing.Point(25, 81);
            this.chkJisyaMemo.Name = "chkJisyaMemo";
            this.chkJisyaMemo.Size = new System.Drawing.Size(99, 22);
            this.chkJisyaMemo.TabIndex = 4;
            this.chkJisyaMemo.Text = "自社メモ情報";
            this.chkJisyaMemo.UseVisualStyleBackColor = true;
            // 
            // chkFBKozafurikae
            // 
            this.chkFBKozafurikae.AutoSize = true;
            this.chkFBKozafurikae.Location = new System.Drawing.Point(56, 165);
            this.chkFBKozafurikae.Name = "chkFBKozafurikae";
            this.chkFBKozafurikae.Size = new System.Drawing.Size(99, 22);
            this.chkFBKozafurikae.TabIndex = 3;
            this.chkFBKozafurikae.Text = "口座振替情報";
            this.chkFBKozafurikae.UseVisualStyleBackColor = true;
            // 
            // chkFBFuritesuryo
            // 
            this.chkFBFuritesuryo.AutoSize = true;
            this.chkFBFuritesuryo.Location = new System.Drawing.Point(56, 193);
            this.chkFBFuritesuryo.Name = "chkFBFuritesuryo";
            this.chkFBFuritesuryo.Size = new System.Drawing.Size(111, 22);
            this.chkFBFuritesuryo.TabIndex = 8;
            this.chkFBFuritesuryo.Text = "振込手数料情報";
            this.chkFBFuritesuryo.UseVisualStyleBackColor = true;
            // 
            // tabPageBase140
            // 
            this.tabPageBase140.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase140.Controls.Add(this.grpBk);
            this.tabPageBase140.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase140.Name = "tabPageBase140";
            this.tabPageBase140.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase140.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase140.TabIndex = 4;
            this.tabPageBase140.Text = "物件情報";
            // 
            // grpBk
            // 
            this.grpBk.Controls.Add(this.chkOpKys);
            this.grpBk.Controls.Add(this.chkOpOw);
            this.grpBk.Controls.Add(this.grpKagiSelect);
            this.grpBk.Controls.Add(this.chkBkSyo);
            this.grpBk.Controls.Add(this.chkBkHendo);
            this.grpBk.Controls.Add(this.chkBkKinrincyusyajo);
            this.grpBk.Controls.Add(this.chkBkSansyofile);
            this.grpBk.Controls.Add(this.chkBkSzeniji);
            this.grpBk.Controls.Add(this.chkBkSyuhen);
            this.grpBk.Controls.Add(this.chkBkSetudo);
            this.grpBk.Controls.Add(this.chkBkKotu);
            this.grpBk.Controls.Add(this.chkBkKenri);
            this.grpBk.Controls.Add(this.chkBkGomi);
            this.grpBk.Controls.Add(this.chkBkSyosai);
            this.grpBk.Controls.Add(this.chkBkKagi);
            this.grpBk.Controls.Add(this.chkBkMemo);
            this.grpBk.Controls.Add(this.chkBkBase);
            this.grpBk.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpBk.Location = new System.Drawing.Point(10, 5);
            this.grpBk.Name = "grpBk";
            this.grpBk.Size = new System.Drawing.Size(817, 265);
            this.grpBk.TabIndex = 0;
            this.grpBk.TabStop = false;
            this.grpBk.Text = "物件情報";
            // 
            // chkOpKys
            // 
            this.chkOpKys.AutoSize = true;
            this.chkOpKys.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkOpKys.Location = new System.Drawing.Point(545, 179);
            this.chkOpKys.Name = "chkOpKys";
            this.chkOpKys.Size = new System.Drawing.Size(229, 22);
            this.chkOpKys.TabIndex = 16;
            this.chkOpKys.Text = "(未使用の契約者データは移行しない)";
            this.chkOpKys.UseVisualStyleBackColor = true;
            // 
            // chkOpOw
            // 
            this.chkOpOw.AutoSize = true;
            this.chkOpOw.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkOpOw.Location = new System.Drawing.Point(545, 151);
            this.chkOpOw.Name = "chkOpOw";
            this.chkOpOw.Size = new System.Drawing.Size(217, 22);
            this.chkOpOw.TabIndex = 15;
            this.chkOpOw.Text = "(未使用の家主データは移行しない)";
            this.chkOpOw.UseVisualStyleBackColor = true;
            // 
            // grpKagiSelect
            // 
            this.grpKagiSelect.Controls.Add(this.optKyKagi);
            this.grpKagiSelect.Controls.Add(this.optHyKagi);
            this.grpKagiSelect.Location = new System.Drawing.Point(557, 81);
            this.grpKagiSelect.Name = "grpKagiSelect";
            this.grpKagiSelect.Size = new System.Drawing.Size(240, 50);
            this.grpKagiSelect.TabIndex = 14;
            this.grpKagiSelect.TabStop = false;
            this.grpKagiSelect.Text = " 取得する鍵情報 ";
            // 
            // optKyKagi
            // 
            this.optKyKagi.AutoSize = true;
            this.optKyKagi.Location = new System.Drawing.Point(140, 20);
            this.optKyKagi.Name = "optKyKagi";
            this.optKyKagi.Size = new System.Drawing.Size(86, 22);
            this.optKyKagi.TabIndex = 1;
            this.optKyKagi.TabStop = true;
            this.optKyKagi.Text = "契約鍵情報";
            this.optKyKagi.UseVisualStyleBackColor = true;
            // 
            // optHyKagi
            // 
            this.optHyKagi.AutoSize = true;
            this.optHyKagi.Checked = true;
            this.optHyKagi.Location = new System.Drawing.Point(20, 20);
            this.optHyKagi.Name = "optHyKagi";
            this.optHyKagi.Size = new System.Drawing.Size(86, 22);
            this.optHyKagi.TabIndex = 0;
            this.optHyKagi.TabStop = true;
            this.optHyKagi.Text = "部屋鍵情報";
            this.optHyKagi.UseVisualStyleBackColor = true;
            // 
            // chkBkSyo
            // 
            this.chkBkSyo.AutoSize = true;
            this.chkBkSyo.Location = new System.Drawing.Point(295, 193);
            this.chkBkSyo.Name = "chkBkSyo";
            this.chkBkSyo.Size = new System.Drawing.Size(111, 22);
            this.chkBkSyo.TabIndex = 12;
            this.chkBkSyo.Text = "物件所有者情報";
            this.chkBkSyo.UseVisualStyleBackColor = true;
            // 
            // chkBkHendo
            // 
            this.chkBkHendo.AutoSize = true;
            this.chkBkHendo.Location = new System.Drawing.Point(295, 109);
            this.chkBkHendo.Name = "chkBkHendo";
            this.chkBkHendo.Size = new System.Drawing.Size(171, 22);
            this.chkBkHendo.TabIndex = 9;
            this.chkBkHendo.Text = "物件変動費親メーター情報";
            this.chkBkHendo.UseVisualStyleBackColor = true;
            // 
            // chkBkKinrincyusyajo
            // 
            this.chkBkKinrincyusyajo.AutoSize = true;
            this.chkBkKinrincyusyajo.Location = new System.Drawing.Point(295, 137);
            this.chkBkKinrincyusyajo.Name = "chkBkKinrincyusyajo";
            this.chkBkKinrincyusyajo.Size = new System.Drawing.Size(135, 22);
            this.chkBkKinrincyusyajo.TabIndex = 10;
            this.chkBkKinrincyusyajo.Text = "物件近隣駐車場情報";
            this.chkBkKinrincyusyajo.UseVisualStyleBackColor = true;
            // 
            // chkBkSansyofile
            // 
            this.chkBkSansyofile.AutoSize = true;
            this.chkBkSansyofile.Location = new System.Drawing.Point(295, 165);
            this.chkBkSansyofile.Name = "chkBkSansyofile";
            this.chkBkSansyofile.Size = new System.Drawing.Size(147, 22);
            this.chkBkSansyofile.TabIndex = 11;
            this.chkBkSansyofile.Text = "物件参照ファイル情報";
            this.chkBkSansyofile.UseVisualStyleBackColor = true;
            // 
            // chkBkSzeniji
            // 
            this.chkBkSzeniji.AutoSize = true;
            this.chkBkSzeniji.Location = new System.Drawing.Point(295, 81);
            this.chkBkSzeniji.Name = "chkBkSzeniji";
            this.chkBkSzeniji.Size = new System.Drawing.Size(183, 22);
            this.chkBkSzeniji.TabIndex = 8;
            this.chkBkSzeniji.Text = "物件修繕維持管理連絡先情報";
            this.chkBkSzeniji.UseVisualStyleBackColor = true;
            // 
            // chkBkSyuhen
            // 
            this.chkBkSyuhen.AutoSize = true;
            this.chkBkSyuhen.Location = new System.Drawing.Point(45, 81);
            this.chkBkSyuhen.Name = "chkBkSyuhen";
            this.chkBkSyuhen.Size = new System.Drawing.Size(99, 22);
            this.chkBkSyuhen.TabIndex = 2;
            this.chkBkSyuhen.Text = "物件周辺情報";
            this.chkBkSyuhen.UseVisualStyleBackColor = true;
            // 
            // chkBkSetudo
            // 
            this.chkBkSetudo.AutoSize = true;
            this.chkBkSetudo.Location = new System.Drawing.Point(45, 193);
            this.chkBkSetudo.Name = "chkBkSetudo";
            this.chkBkSetudo.Size = new System.Drawing.Size(99, 22);
            this.chkBkSetudo.TabIndex = 6;
            this.chkBkSetudo.Text = "物件接道情報";
            this.chkBkSetudo.UseVisualStyleBackColor = true;
            // 
            // chkBkKotu
            // 
            this.chkBkKotu.AutoSize = true;
            this.chkBkKotu.Location = new System.Drawing.Point(45, 165);
            this.chkBkKotu.Name = "chkBkKotu";
            this.chkBkKotu.Size = new System.Drawing.Size(99, 22);
            this.chkBkKotu.TabIndex = 5;
            this.chkBkKotu.Text = "物件交通情報";
            this.chkBkKotu.UseVisualStyleBackColor = true;
            // 
            // chkBkKenri
            // 
            this.chkBkKenri.AutoSize = true;
            this.chkBkKenri.Location = new System.Drawing.Point(45, 137);
            this.chkBkKenri.Name = "chkBkKenri";
            this.chkBkKenri.Size = new System.Drawing.Size(99, 22);
            this.chkBkKenri.TabIndex = 4;
            this.chkBkKenri.Text = "物件権利情報";
            this.chkBkKenri.UseVisualStyleBackColor = true;
            // 
            // chkBkGomi
            // 
            this.chkBkGomi.AutoSize = true;
            this.chkBkGomi.Location = new System.Drawing.Point(45, 109);
            this.chkBkGomi.Name = "chkBkGomi";
            this.chkBkGomi.Size = new System.Drawing.Size(99, 22);
            this.chkBkGomi.TabIndex = 3;
            this.chkBkGomi.Text = "物件ゴミ情報";
            this.chkBkGomi.UseVisualStyleBackColor = true;
            // 
            // chkBkSyosai
            // 
            this.chkBkSyosai.AutoSize = true;
            this.chkBkSyosai.Location = new System.Drawing.Point(45, 53);
            this.chkBkSyosai.Name = "chkBkSyosai";
            this.chkBkSyosai.Size = new System.Drawing.Size(99, 22);
            this.chkBkSyosai.TabIndex = 1;
            this.chkBkSyosai.Text = "物件詳細情報";
            this.chkBkSyosai.UseVisualStyleBackColor = true;
            // 
            // chkBkKagi
            // 
            this.chkBkKagi.AutoSize = true;
            this.chkBkKagi.Location = new System.Drawing.Point(545, 53);
            this.chkBkKagi.Name = "chkBkKagi";
            this.chkBkKagi.Size = new System.Drawing.Size(87, 22);
            this.chkBkKagi.TabIndex = 13;
            this.chkBkKagi.Text = "物件鍵情報";
            this.chkBkKagi.UseVisualStyleBackColor = true;
            // 
            // chkBkMemo
            // 
            this.chkBkMemo.AutoSize = true;
            this.chkBkMemo.Location = new System.Drawing.Point(295, 53);
            this.chkBkMemo.Name = "chkBkMemo";
            this.chkBkMemo.Size = new System.Drawing.Size(99, 22);
            this.chkBkMemo.TabIndex = 7;
            this.chkBkMemo.Text = "物件メモ情報";
            this.chkBkMemo.UseVisualStyleBackColor = true;
            // 
            // chkBkBase
            // 
            this.chkBkBase.AutoSize = true;
            this.chkBkBase.Location = new System.Drawing.Point(25, 25);
            this.chkBkBase.Name = "chkBkBase";
            this.chkBkBase.Size = new System.Drawing.Size(99, 22);
            this.chkBkBase.TabIndex = 0;
            this.chkBkBase.Text = "物件基本情報";
            this.chkBkBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase150
            // 
            this.tabPageBase150.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase150.Controls.Add(this.grpHy);
            this.tabPageBase150.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase150.Name = "tabPageBase150";
            this.tabPageBase150.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase150.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase150.TabIndex = 14;
            this.tabPageBase150.Text = "部屋情報";
            // 
            // grpHy
            // 
            this.grpHy.Controls.Add(this.chkHySyo);
            this.grpHy.Controls.Add(this.chkHySansyofile);
            this.grpHy.Controls.Add(this.chkHyGenjotanka);
            this.grpHy.Controls.Add(this.chkHyKenri);
            this.grpHy.Controls.Add(this.chkHyConfirm);
            this.grpHy.Controls.Add(this.chkHyHendo);
            this.grpHy.Controls.Add(this.chkHyCommonsalespoint);
            this.grpHy.Controls.Add(this.chkHyMenseki);
            this.grpHy.Controls.Add(this.chkHyNkinkomk);
            this.grpHy.Controls.Add(this.chkHyMadoriutiwake);
            this.grpHy.Controls.Add(this.chkHySzeniji);
            this.grpHy.Controls.Add(this.chkHyTokuyaku);
            this.grpHy.Controls.Add(this.chkHyParking);
            this.grpHy.Controls.Add(this.chkHySyosai);
            this.grpHy.Controls.Add(this.chkHyMemo);
            this.grpHy.Controls.Add(this.chkHyKagi);
            this.grpHy.Controls.Add(this.chkHySetubi);
            this.grpHy.Controls.Add(this.chkHyBase);
            this.grpHy.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpHy.Location = new System.Drawing.Point(10, 5);
            this.grpHy.Name = "grpHy";
            this.grpHy.Size = new System.Drawing.Size(817, 265);
            this.grpHy.TabIndex = 0;
            this.grpHy.TabStop = false;
            this.grpHy.Text = "部屋情報";
            // 
            // chkHySyo
            // 
            this.chkHySyo.AutoSize = true;
            this.chkHySyo.Location = new System.Drawing.Point(295, 164);
            this.chkHySyo.Name = "chkHySyo";
            this.chkHySyo.Size = new System.Drawing.Size(111, 22);
            this.chkHySyo.TabIndex = 12;
            this.chkHySyo.Text = "部屋所有者情報";
            this.chkHySyo.UseVisualStyleBackColor = true;
            // 
            // chkHySansyofile
            // 
            this.chkHySansyofile.AutoSize = true;
            this.chkHySansyofile.Location = new System.Drawing.Point(545, 81);
            this.chkHySansyofile.Name = "chkHySansyofile";
            this.chkHySansyofile.Size = new System.Drawing.Size(147, 22);
            this.chkHySansyofile.TabIndex = 16;
            this.chkHySansyofile.Text = "部屋参照ファイル情報";
            this.chkHySansyofile.UseVisualStyleBackColor = true;
            // 
            // chkHyGenjotanka
            // 
            this.chkHyGenjotanka.AutoSize = true;
            this.chkHyGenjotanka.Location = new System.Drawing.Point(545, 109);
            this.chkHyGenjotanka.Name = "chkHyGenjotanka";
            this.chkHyGenjotanka.Size = new System.Drawing.Size(171, 22);
            this.chkHyGenjotanka.TabIndex = 17;
            this.chkHyGenjotanka.Text = "部屋原状回復目安単価情報";
            this.chkHyGenjotanka.UseVisualStyleBackColor = true;
            // 
            // chkHyKenri
            // 
            this.chkHyKenri.AutoSize = true;
            this.chkHyKenri.Location = new System.Drawing.Point(545, 53);
            this.chkHyKenri.Name = "chkHyKenri";
            this.chkHyKenri.Size = new System.Drawing.Size(99, 22);
            this.chkHyKenri.TabIndex = 15;
            this.chkHyKenri.Text = "部屋権利情報";
            this.chkHyKenri.UseVisualStyleBackColor = true;
            // 
            // chkHyConfirm
            // 
            this.chkHyConfirm.AutoSize = true;
            this.chkHyConfirm.Location = new System.Drawing.Point(295, 221);
            this.chkHyConfirm.Name = "chkHyConfirm";
            this.chkHyConfirm.Size = new System.Drawing.Size(171, 22);
            this.chkHyConfirm.TabIndex = 14;
            this.chkHyConfirm.Text = "部屋契約解約確認事項情報";
            this.chkHyConfirm.UseVisualStyleBackColor = true;
            // 
            // chkHyHendo
            // 
            this.chkHyHendo.AutoSize = true;
            this.chkHyHendo.Location = new System.Drawing.Point(295, 137);
            this.chkHyHendo.Name = "chkHyHendo";
            this.chkHyHendo.Size = new System.Drawing.Size(183, 22);
            this.chkHyHendo.TabIndex = 11;
            this.chkHyHendo.Text = "部屋変動費各戸メーター情報";
            this.chkHyHendo.UseVisualStyleBackColor = true;
            // 
            // chkHyCommonsalespoint
            // 
            this.chkHyCommonsalespoint.AutoSize = true;
            this.chkHyCommonsalespoint.Location = new System.Drawing.Point(295, 193);
            this.chkHyCommonsalespoint.Name = "chkHyCommonsalespoint";
            this.chkHyCommonsalespoint.Size = new System.Drawing.Size(195, 22);
            this.chkHyCommonsalespoint.TabIndex = 13;
            this.chkHyCommonsalespoint.Text = "部屋共通セールスポイント情報";
            this.chkHyCommonsalespoint.UseVisualStyleBackColor = true;
            // 
            // chkHyMenseki
            // 
            this.chkHyMenseki.AutoSize = true;
            this.chkHyMenseki.Location = new System.Drawing.Point(45, 109);
            this.chkHyMenseki.Name = "chkHyMenseki";
            this.chkHyMenseki.Size = new System.Drawing.Size(99, 22);
            this.chkHyMenseki.TabIndex = 3;
            this.chkHyMenseki.Text = "部屋面積情報";
            this.chkHyMenseki.UseVisualStyleBackColor = true;
            // 
            // chkHyNkinkomk
            // 
            this.chkHyNkinkomk.AutoSize = true;
            this.chkHyNkinkomk.Location = new System.Drawing.Point(295, 81);
            this.chkHyNkinkomk.Name = "chkHyNkinkomk";
            this.chkHyNkinkomk.Size = new System.Drawing.Size(123, 22);
            this.chkHyNkinkomk.TabIndex = 9;
            this.chkHyNkinkomk.Text = "部屋入金項目情報";
            this.chkHyNkinkomk.UseVisualStyleBackColor = true;
            // 
            // chkHyMadoriutiwake
            // 
            this.chkHyMadoriutiwake.AutoSize = true;
            this.chkHyMadoriutiwake.Location = new System.Drawing.Point(45, 165);
            this.chkHyMadoriutiwake.Name = "chkHyMadoriutiwake";
            this.chkHyMadoriutiwake.Size = new System.Drawing.Size(123, 22);
            this.chkHyMadoriutiwake.TabIndex = 5;
            this.chkHyMadoriutiwake.Text = "部屋間取内訳情報";
            this.chkHyMadoriutiwake.UseVisualStyleBackColor = true;
            // 
            // chkHySzeniji
            // 
            this.chkHySzeniji.AutoSize = true;
            this.chkHySzeniji.Location = new System.Drawing.Point(295, 109);
            this.chkHySzeniji.Name = "chkHySzeniji";
            this.chkHySzeniji.Size = new System.Drawing.Size(183, 22);
            this.chkHySzeniji.TabIndex = 10;
            this.chkHySzeniji.Text = "部屋修繕維持管理連絡先情報";
            this.chkHySzeniji.UseVisualStyleBackColor = true;
            // 
            // chkHyTokuyaku
            // 
            this.chkHyTokuyaku.AutoSize = true;
            this.chkHyTokuyaku.Location = new System.Drawing.Point(45, 221);
            this.chkHyTokuyaku.Name = "chkHyTokuyaku";
            this.chkHyTokuyaku.Size = new System.Drawing.Size(99, 22);
            this.chkHyTokuyaku.TabIndex = 7;
            this.chkHyTokuyaku.Text = "部屋特約情報";
            this.chkHyTokuyaku.UseVisualStyleBackColor = true;
            // 
            // chkHyParking
            // 
            this.chkHyParking.AutoSize = true;
            this.chkHyParking.Location = new System.Drawing.Point(45, 193);
            this.chkHyParking.Name = "chkHyParking";
            this.chkHyParking.Size = new System.Drawing.Size(111, 22);
            this.chkHyParking.TabIndex = 6;
            this.chkHyParking.Text = "部屋駐車場情報";
            this.chkHyParking.UseVisualStyleBackColor = true;
            // 
            // chkHySyosai
            // 
            this.chkHySyosai.AutoSize = true;
            this.chkHySyosai.Location = new System.Drawing.Point(45, 53);
            this.chkHySyosai.Name = "chkHySyosai";
            this.chkHySyosai.Size = new System.Drawing.Size(99, 22);
            this.chkHySyosai.TabIndex = 1;
            this.chkHySyosai.Text = "部屋詳細情報";
            this.chkHySyosai.UseVisualStyleBackColor = true;
            // 
            // chkHyMemo
            // 
            this.chkHyMemo.AutoSize = true;
            this.chkHyMemo.Location = new System.Drawing.Point(45, 137);
            this.chkHyMemo.Name = "chkHyMemo";
            this.chkHyMemo.Size = new System.Drawing.Size(99, 22);
            this.chkHyMemo.TabIndex = 4;
            this.chkHyMemo.Text = "部屋メモ情報";
            this.chkHyMemo.UseVisualStyleBackColor = true;
            // 
            // chkHyKagi
            // 
            this.chkHyKagi.AutoSize = true;
            this.chkHyKagi.Location = new System.Drawing.Point(295, 53);
            this.chkHyKagi.Name = "chkHyKagi";
            this.chkHyKagi.Size = new System.Drawing.Size(87, 22);
            this.chkHyKagi.TabIndex = 8;
            this.chkHyKagi.Text = "部屋鍵情報";
            this.chkHyKagi.UseVisualStyleBackColor = true;
            // 
            // chkHySetubi
            // 
            this.chkHySetubi.AutoSize = true;
            this.chkHySetubi.Location = new System.Drawing.Point(45, 81);
            this.chkHySetubi.Name = "chkHySetubi";
            this.chkHySetubi.Size = new System.Drawing.Size(99, 22);
            this.chkHySetubi.TabIndex = 2;
            this.chkHySetubi.Text = "部屋設備情報";
            this.chkHySetubi.UseVisualStyleBackColor = true;
            // 
            // chkHyBase
            // 
            this.chkHyBase.AutoSize = true;
            this.chkHyBase.Location = new System.Drawing.Point(25, 25);
            this.chkHyBase.Name = "chkHyBase";
            this.chkHyBase.Size = new System.Drawing.Size(99, 22);
            this.chkHyBase.TabIndex = 0;
            this.chkHyBase.Text = "部屋基本情報";
            this.chkHyBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase160
            // 
            this.tabPageBase160.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase160.Controls.Add(this.grpSorule);
            this.tabPageBase160.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase160.Name = "tabPageBase160";
            this.tabPageBase160.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase160.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase160.TabIndex = 5;
            this.tabPageBase160.Text = "送金ルール";
            // 
            // grpSorule
            // 
            this.grpSorule.Controls.Add(this.chkSoruleSosaki);
            this.grpSorule.Controls.Add(this.chkSoruleKojo);
            this.grpSorule.Controls.Add(this.chkSoruleNkin);
            this.grpSorule.Controls.Add(this.chkSoruleBase);
            this.grpSorule.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpSorule.Location = new System.Drawing.Point(10, 5);
            this.grpSorule.Name = "grpSorule";
            this.grpSorule.Size = new System.Drawing.Size(282, 265);
            this.grpSorule.TabIndex = 0;
            this.grpSorule.TabStop = false;
            this.grpSorule.Text = "送金ルール情報";
            // 
            // chkSoruleSosaki
            // 
            this.chkSoruleSosaki.AutoSize = true;
            this.chkSoruleSosaki.Location = new System.Drawing.Point(45, 53);
            this.chkSoruleSosaki.Name = "chkSoruleSosaki";
            this.chkSoruleSosaki.Size = new System.Drawing.Size(147, 22);
            this.chkSoruleSosaki.TabIndex = 1;
            this.chkSoruleSosaki.Text = "送金ルール送金先情報";
            this.chkSoruleSosaki.UseVisualStyleBackColor = true;
            // 
            // chkSoruleKojo
            // 
            this.chkSoruleKojo.AutoSize = true;
            this.chkSoruleKojo.Location = new System.Drawing.Point(45, 109);
            this.chkSoruleKojo.Name = "chkSoruleKojo";
            this.chkSoruleKojo.Size = new System.Drawing.Size(159, 22);
            this.chkSoruleKojo.TabIndex = 3;
            this.chkSoruleKojo.Text = "送金ルール控除項目情報";
            this.chkSoruleKojo.UseVisualStyleBackColor = true;
            // 
            // chkSoruleNkin
            // 
            this.chkSoruleNkin.AutoSize = true;
            this.chkSoruleNkin.Location = new System.Drawing.Point(45, 81);
            this.chkSoruleNkin.Name = "chkSoruleNkin";
            this.chkSoruleNkin.Size = new System.Drawing.Size(159, 22);
            this.chkSoruleNkin.TabIndex = 2;
            this.chkSoruleNkin.Text = "送金ルール入金項目情報";
            this.chkSoruleNkin.UseVisualStyleBackColor = true;
            // 
            // chkSoruleBase
            // 
            this.chkSoruleBase.AutoSize = true;
            this.chkSoruleBase.Location = new System.Drawing.Point(25, 25);
            this.chkSoruleBase.Name = "chkSoruleBase";
            this.chkSoruleBase.Size = new System.Drawing.Size(135, 22);
            this.chkSoruleBase.TabIndex = 0;
            this.chkSoruleBase.Text = "送金ルール基本情報";
            this.chkSoruleBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase170
            // 
            this.tabPageBase170.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase170.Controls.Add(this.grpKy);
            this.tabPageBase170.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase170.Name = "tabPageBase170";
            this.tabPageBase170.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase170.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase170.TabIndex = 6;
            this.tabPageBase170.Text = "契約情報";
            // 
            // grpKy
            // 
            this.grpKy.Controls.Add(this.chkKySzenmeisai);
            this.grpKy.Controls.Add(this.chkKySzen);
            this.grpKy.Controls.Add(this.chkKyKai);
            this.grpKy.Controls.Add(this.CheckBox14);
            this.grpKy.Controls.Add(this.CheckBox24);
            this.grpKy.Controls.Add(this.CheckBox23);
            this.grpKy.Controls.Add(this.chkKyHosyonin);
            this.grpKy.Controls.Add(this.chkKyMemo);
            this.grpKy.Controls.Add(this.chkKyTokuyaku);
            this.grpKy.Controls.Add(this.CheckBox18);
            this.grpKy.Controls.Add(this.chkKySorule);
            this.grpKy.Controls.Add(this.CheckBox16);
            this.grpKy.Controls.Add(this.CheckBox15);
            this.grpKy.Controls.Add(this.chkKyNyukyo);
            this.grpKy.Controls.Add(this.chkKyNkinkomkNx);
            this.grpKy.Controls.Add(this.chkKyNkinkomk);
            this.grpKy.Controls.Add(this.chkKyHoken);
            this.grpKy.Controls.Add(this.chkKyKojoRule);
            this.grpKy.Controls.Add(this.chkKyKys);
            this.grpKy.Controls.Add(this.CheckBox5);
            this.grpKy.Controls.Add(this.chkKyRireki);
            this.grpKy.Controls.Add(this.chkKyHendo);
            this.grpKy.Controls.Add(this.chkKyCar);
            this.grpKy.Controls.Add(this.chkKyBase);
            this.grpKy.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpKy.Location = new System.Drawing.Point(10, 5);
            this.grpKy.Name = "grpKy";
            this.grpKy.Size = new System.Drawing.Size(819, 265);
            this.grpKy.TabIndex = 0;
            this.grpKy.TabStop = false;
            this.grpKy.Text = "契約情報";
            // 
            // chkKySzenmeisai
            // 
            this.chkKySzenmeisai.AutoSize = true;
            this.chkKySzenmeisai.Location = new System.Drawing.Point(400, 80);
            this.chkKySzenmeisai.Name = "chkKySzenmeisai";
            this.chkKySzenmeisai.Size = new System.Drawing.Size(147, 22);
            this.chkKySzenmeisai.TabIndex = 23;
            this.chkKySzenmeisai.Text = "契約修繕見積詳細情報";
            this.chkKySzenmeisai.UseVisualStyleBackColor = true;
            // 
            // chkKySzen
            // 
            this.chkKySzen.AutoSize = true;
            this.chkKySzen.Location = new System.Drawing.Point(400, 48);
            this.chkKySzen.Name = "chkKySzen";
            this.chkKySzen.Size = new System.Drawing.Size(123, 22);
            this.chkKySzen.TabIndex = 22;
            this.chkKySzen.Text = "契約修繕見積情報";
            this.chkKySzen.UseVisualStyleBackColor = true;
            // 
            // chkKyKai
            // 
            this.chkKyKai.AutoSize = true;
            this.chkKyKai.Location = new System.Drawing.Point(220, 221);
            this.chkKyKai.Name = "chkKyKai";
            this.chkKyKai.Size = new System.Drawing.Size(99, 22);
            this.chkKyKai.TabIndex = 14;
            this.chkKyKai.Text = "契約解約情報";
            this.chkKyKai.UseVisualStyleBackColor = true;
            // 
            // CheckBox14
            // 
            this.CheckBox14.AutoSize = true;
            this.CheckBox14.Location = new System.Drawing.Point(576, 104);
            this.CheckBox14.Name = "CheckBox14";
            this.CheckBox14.Size = new System.Drawing.Size(171, 22);
            this.CheckBox14.TabIndex = 17;
            this.CheckBox14.Text = "契約変動費親メーター情報";
            this.CheckBox14.UseVisualStyleBackColor = true;
            // 
            // CheckBox24
            // 
            this.CheckBox24.AutoSize = true;
            this.CheckBox24.Location = new System.Drawing.Point(576, 48);
            this.CheckBox24.Name = "CheckBox24";
            this.CheckBox24.Size = new System.Drawing.Size(147, 22);
            this.CheckBox24.TabIndex = 15;
            this.CheckBox24.Text = "契約解約確認事項情報";
            this.CheckBox24.UseVisualStyleBackColor = true;
            // 
            // CheckBox23
            // 
            this.CheckBox23.AutoSize = true;
            this.CheckBox23.Location = new System.Drawing.Point(576, 76);
            this.CheckBox23.Name = "CheckBox23";
            this.CheckBox23.Size = new System.Drawing.Size(123, 22);
            this.CheckBox23.TabIndex = 16;
            this.CheckBox23.Text = "契約同時契約情報";
            this.CheckBox23.UseVisualStyleBackColor = true;
            // 
            // chkKyHosyonin
            // 
            this.chkKyHosyonin.AutoSize = true;
            this.chkKyHosyonin.Location = new System.Drawing.Point(45, 137);
            this.chkKyHosyonin.Name = "chkKyHosyonin";
            this.chkKyHosyonin.Size = new System.Drawing.Size(111, 22);
            this.chkKyHosyonin.TabIndex = 4;
            this.chkKyHosyonin.Text = "契約保証人情報";
            this.chkKyHosyonin.UseVisualStyleBackColor = true;
            // 
            // chkKyMemo
            // 
            this.chkKyMemo.AutoSize = true;
            this.chkKyMemo.Location = new System.Drawing.Point(220, 53);
            this.chkKyMemo.Name = "chkKyMemo";
            this.chkKyMemo.Size = new System.Drawing.Size(99, 22);
            this.chkKyMemo.TabIndex = 8;
            this.chkKyMemo.Text = "契約メモ情報";
            this.chkKyMemo.UseVisualStyleBackColor = true;
            // 
            // chkKyTokuyaku
            // 
            this.chkKyTokuyaku.AutoSize = true;
            this.chkKyTokuyaku.Location = new System.Drawing.Point(45, 221);
            this.chkKyTokuyaku.Name = "chkKyTokuyaku";
            this.chkKyTokuyaku.Size = new System.Drawing.Size(123, 22);
            this.chkKyTokuyaku.TabIndex = 7;
            this.chkKyTokuyaku.Text = "契約特約事項情報";
            this.chkKyTokuyaku.UseVisualStyleBackColor = true;
            // 
            // CheckBox18
            // 
            this.CheckBox18.AutoSize = true;
            this.CheckBox18.Location = new System.Drawing.Point(576, 216);
            this.CheckBox18.Name = "CheckBox18";
            this.CheckBox18.Size = new System.Drawing.Size(171, 22);
            this.CheckBox18.TabIndex = 21;
            this.CheckBox18.Text = "契約原状回復目安単価情報";
            this.CheckBox18.UseVisualStyleBackColor = true;
            // 
            // chkKySorule
            // 
            this.chkKySorule.AutoSize = true;
            this.chkKySorule.Location = new System.Drawing.Point(220, 192);
            this.chkKySorule.Name = "chkKySorule";
            this.chkKySorule.Size = new System.Drawing.Size(135, 22);
            this.chkKySorule.TabIndex = 13;
            this.chkKySorule.Text = "契約送金ルール情報";
            this.chkKySorule.UseVisualStyleBackColor = true;
            // 
            // CheckBox16
            // 
            this.CheckBox16.AutoSize = true;
            this.CheckBox16.Location = new System.Drawing.Point(576, 187);
            this.CheckBox16.Name = "CheckBox16";
            this.CheckBox16.Size = new System.Drawing.Size(183, 22);
            this.CheckBox16.TabIndex = 20;
            this.CheckBox16.Text = "契約敷金保証金随時処理情報";
            this.CheckBox16.UseVisualStyleBackColor = true;
            // 
            // CheckBox15
            // 
            this.CheckBox15.AutoSize = true;
            this.CheckBox15.Location = new System.Drawing.Point(576, 159);
            this.CheckBox15.Name = "CheckBox15";
            this.CheckBox15.Size = new System.Drawing.Size(147, 22);
            this.CheckBox15.TabIndex = 19;
            this.CheckBox15.Text = "契約関連ファイル情報";
            this.CheckBox15.UseVisualStyleBackColor = true;
            // 
            // chkKyNyukyo
            // 
            this.chkKyNyukyo.AutoSize = true;
            this.chkKyNyukyo.Location = new System.Drawing.Point(45, 109);
            this.chkKyNyukyo.Name = "chkKyNyukyo";
            this.chkKyNyukyo.Size = new System.Drawing.Size(111, 22);
            this.chkKyNyukyo.TabIndex = 3;
            this.chkKyNyukyo.Text = "契約入居者情報";
            this.chkKyNyukyo.UseVisualStyleBackColor = true;
            // 
            // chkKyNkinkomkNx
            // 
            this.chkKyNkinkomkNx.AutoSize = true;
            this.chkKyNkinkomkNx.Location = new System.Drawing.Point(220, 109);
            this.chkKyNkinkomkNx.Name = "chkKyNkinkomkNx";
            this.chkKyNkinkomkNx.Size = new System.Drawing.Size(147, 22);
            this.chkKyNkinkomkNx.TabIndex = 10;
            this.chkKyNkinkomkNx.Text = "契約次回入金項目情報";
            this.chkKyNkinkomkNx.UseVisualStyleBackColor = true;
            // 
            // chkKyNkinkomk
            // 
            this.chkKyNkinkomk.AutoSize = true;
            this.chkKyNkinkomk.Location = new System.Drawing.Point(220, 81);
            this.chkKyNkinkomk.Name = "chkKyNkinkomk";
            this.chkKyNkinkomk.Size = new System.Drawing.Size(123, 22);
            this.chkKyNkinkomk.TabIndex = 9;
            this.chkKyNkinkomk.Text = "契約入金項目情報";
            this.chkKyNkinkomk.UseVisualStyleBackColor = true;
            // 
            // chkKyHoken
            // 
            this.chkKyHoken.AutoSize = true;
            this.chkKyHoken.Location = new System.Drawing.Point(45, 193);
            this.chkKyHoken.Name = "chkKyHoken";
            this.chkKyHoken.Size = new System.Drawing.Size(99, 22);
            this.chkKyHoken.TabIndex = 6;
            this.chkKyHoken.Text = "契約保険情報";
            this.chkKyHoken.UseVisualStyleBackColor = true;
            // 
            // chkKyKojoRule
            // 
            this.chkKyKojoRule.AutoSize = true;
            this.chkKyKojoRule.Location = new System.Drawing.Point(220, 164);
            this.chkKyKojoRule.Name = "chkKyKojoRule";
            this.chkKyKojoRule.Size = new System.Drawing.Size(135, 22);
            this.chkKyKojoRule.TabIndex = 12;
            this.chkKyKojoRule.Text = "契約控除ルール情報";
            this.chkKyKojoRule.UseVisualStyleBackColor = true;
            // 
            // chkKyKys
            // 
            this.chkKyKys.AutoSize = true;
            this.chkKyKys.Location = new System.Drawing.Point(45, 81);
            this.chkKyKys.Name = "chkKyKys";
            this.chkKyKys.Size = new System.Drawing.Size(111, 22);
            this.chkKyKys.TabIndex = 2;
            this.chkKyKys.Text = "契約契約者情報";
            this.chkKyKys.UseVisualStyleBackColor = true;
            // 
            // CheckBox5
            // 
            this.CheckBox5.AutoSize = true;
            this.CheckBox5.Location = new System.Drawing.Point(576, 132);
            this.CheckBox5.Name = "CheckBox5";
            this.CheckBox5.Size = new System.Drawing.Size(123, 22);
            this.CheckBox5.TabIndex = 18;
            this.CheckBox5.Text = "契約空室待ち情報";
            this.CheckBox5.UseVisualStyleBackColor = true;
            // 
            // chkKyRireki
            // 
            this.chkKyRireki.AutoSize = true;
            this.chkKyRireki.Location = new System.Drawing.Point(45, 53);
            this.chkKyRireki.Name = "chkKyRireki";
            this.chkKyRireki.Size = new System.Drawing.Size(99, 22);
            this.chkKyRireki.TabIndex = 1;
            this.chkKyRireki.Text = "契約履歴情報";
            this.chkKyRireki.UseVisualStyleBackColor = true;
            // 
            // chkKyHendo
            // 
            this.chkKyHendo.AutoSize = true;
            this.chkKyHendo.Location = new System.Drawing.Point(220, 137);
            this.chkKyHendo.Name = "chkKyHendo";
            this.chkKyHendo.Size = new System.Drawing.Size(183, 22);
            this.chkKyHendo.TabIndex = 11;
            this.chkKyHendo.Text = "契約変動費各戸メーター情報";
            this.chkKyHendo.UseVisualStyleBackColor = true;
            // 
            // chkKyCar
            // 
            this.chkKyCar.AutoSize = true;
            this.chkKyCar.Location = new System.Drawing.Point(45, 165);
            this.chkKyCar.Name = "chkKyCar";
            this.chkKyCar.Size = new System.Drawing.Size(87, 22);
            this.chkKyCar.TabIndex = 5;
            this.chkKyCar.Text = "契約車情報";
            this.chkKyCar.UseVisualStyleBackColor = true;
            // 
            // chkKyBase
            // 
            this.chkKyBase.AutoSize = true;
            this.chkKyBase.Location = new System.Drawing.Point(25, 25);
            this.chkKyBase.Name = "chkKyBase";
            this.chkKyBase.Size = new System.Drawing.Size(99, 22);
            this.chkKyBase.TabIndex = 0;
            this.chkKyBase.Text = "契約基本情報";
            this.chkKyBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase180
            // 
            this.tabPageBase180.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase180.Controls.Add(this.grpSq);
            this.tabPageBase180.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase180.Name = "tabPageBase180";
            this.tabPageBase180.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase180.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase180.TabIndex = 7;
            this.tabPageBase180.Text = "請求情報";
            // 
            // grpSq
            // 
            this.grpSq.Controls.Add(this.chkSqSqKojo);
            this.grpSq.Controls.Add(this.chkSqKoteiKojo);
            this.grpSq.Controls.Add(this.chkSqHendokensin);
            this.grpSq.Controls.Add(this.chkSqUnyotaino);
            this.grpSq.Controls.Add(this.chkSqSq);
            this.grpSq.Controls.Add(this.chkSqKajyo);
            this.grpSq.Location = new System.Drawing.Point(10, 5);
            this.grpSq.Name = "grpSq";
            this.grpSq.Size = new System.Drawing.Size(282, 265);
            this.grpSq.TabIndex = 0;
            this.grpSq.TabStop = false;
            this.grpSq.Text = "請求情報";
            // 
            // chkSqSqKojo
            // 
            this.chkSqSqKojo.AutoSize = true;
            this.chkSqSqKojo.Location = new System.Drawing.Point(25, 165);
            this.chkSqSqKojo.Name = "chkSqSqKojo";
            this.chkSqSqKojo.Size = new System.Drawing.Size(123, 22);
            this.chkSqSqKojo.TabIndex = 5;
            this.chkSqSqKojo.Text = "家主請求控除情報";
            this.chkSqSqKojo.UseVisualStyleBackColor = true;
            // 
            // chkSqKoteiKojo
            // 
            this.chkSqKoteiKojo.AutoSize = true;
            this.chkSqKoteiKojo.Location = new System.Drawing.Point(25, 137);
            this.chkSqKoteiKojo.Name = "chkSqKoteiKojo";
            this.chkSqKoteiKojo.Size = new System.Drawing.Size(123, 22);
            this.chkSqKoteiKojo.TabIndex = 4;
            this.chkSqKoteiKojo.Text = "家主固定控除情報";
            this.chkSqKoteiKojo.UseVisualStyleBackColor = true;
            // 
            // chkSqHendokensin
            // 
            this.chkSqHendokensin.AutoSize = true;
            this.chkSqHendokensin.Location = new System.Drawing.Point(25, 109);
            this.chkSqHendokensin.Name = "chkSqHendokensin";
            this.chkSqHendokensin.Size = new System.Drawing.Size(111, 22);
            this.chkSqHendokensin.TabIndex = 3;
            this.chkSqHendokensin.Text = "変動費検針情報";
            this.chkSqHendokensin.UseVisualStyleBackColor = true;
            // 
            // chkSqUnyotaino
            // 
            this.chkSqUnyotaino.AutoSize = true;
            this.chkSqUnyotaino.Location = new System.Drawing.Point(25, 53);
            this.chkSqUnyotaino.Name = "chkSqUnyotaino";
            this.chkSqUnyotaino.Size = new System.Drawing.Size(111, 22);
            this.chkSqUnyotaino.TabIndex = 1;
            this.chkSqUnyotaino.Text = "未収滞納金情報";
            this.chkSqUnyotaino.UseVisualStyleBackColor = true;
            // 
            // chkSqSq
            // 
            this.chkSqSq.AutoSize = true;
            this.chkSqSq.Location = new System.Drawing.Point(25, 81);
            this.chkSqSq.Name = "chkSqSq";
            this.chkSqSq.Size = new System.Drawing.Size(111, 22);
            this.chkSqSq.TabIndex = 2;
            this.chkSqSq.Text = "その他請求情報";
            this.chkSqSq.UseVisualStyleBackColor = true;
            // 
            // chkSqKajyo
            // 
            this.chkSqKajyo.AutoSize = true;
            this.chkSqKajyo.Location = new System.Drawing.Point(25, 25);
            this.chkSqKajyo.Name = "chkSqKajyo";
            this.chkSqKajyo.Size = new System.Drawing.Size(87, 22);
            this.chkSqKajyo.TabIndex = 0;
            this.chkSqKajyo.Text = "預り金情報";
            this.chkSqKajyo.UseVisualStyleBackColor = true;
            // 
            // tabPageBase190
            // 
            this.tabPageBase190.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase190.Controls.Add(this.grpSzen);
            this.tabPageBase190.Controls.Add(this.grpClaim);
            this.tabPageBase190.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase190.Name = "tabPageBase190";
            this.tabPageBase190.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase190.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase190.TabIndex = 18;
            this.tabPageBase190.Text = "クレーム修繕情報";
            // 
            // grpSzen
            // 
            this.grpSzen.Controls.Add(this.chkSzenRelfile);
            this.grpSzen.Controls.Add(this.chkSzenClaim);
            this.grpSzen.Controls.Add(this.chkSzenMemo);
            this.grpSzen.Controls.Add(this.chkSzenSzen);
            this.grpSzen.Controls.Add(this.chkSzenSzenmeisai);
            this.grpSzen.Controls.Add(this.chkSzenBase);
            this.grpSzen.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpSzen.Location = new System.Drawing.Point(312, 8);
            this.grpSzen.Name = "grpSzen";
            this.grpSzen.Size = new System.Drawing.Size(282, 265);
            this.grpSzen.TabIndex = 2;
            this.grpSzen.TabStop = false;
            this.grpSzen.Text = "修繕情報";
            // 
            // chkSzenRelfile
            // 
            this.chkSzenRelfile.AutoSize = true;
            this.chkSzenRelfile.Location = new System.Drawing.Point(48, 144);
            this.chkSzenRelfile.Name = "chkSzenRelfile";
            this.chkSzenRelfile.Size = new System.Drawing.Size(147, 22);
            this.chkSzenRelfile.TabIndex = 6;
            this.chkSzenRelfile.Text = "修繕関連ファイル情報";
            this.chkSzenRelfile.UseVisualStyleBackColor = true;
            // 
            // chkSzenClaim
            // 
            this.chkSzenClaim.AutoSize = true;
            this.chkSzenClaim.Location = new System.Drawing.Point(48, 112);
            this.chkSzenClaim.Name = "chkSzenClaim";
            this.chkSzenClaim.Size = new System.Drawing.Size(171, 22);
            this.chkSzenClaim.TabIndex = 4;
            this.chkSzenClaim.Text = "修繕クレーム関連付け情報";
            this.chkSzenClaim.UseVisualStyleBackColor = true;
            // 
            // chkSzenMemo
            // 
            this.chkSzenMemo.AutoSize = true;
            this.chkSzenMemo.Location = new System.Drawing.Point(48, 176);
            this.chkSzenMemo.Name = "chkSzenMemo";
            this.chkSzenMemo.Size = new System.Drawing.Size(99, 22);
            this.chkSzenMemo.TabIndex = 3;
            this.chkSzenMemo.Text = "修繕メモ情報";
            this.chkSzenMemo.UseVisualStyleBackColor = true;
            // 
            // chkSzenSzen
            // 
            this.chkSzenSzen.AutoSize = true;
            this.chkSzenSzen.Location = new System.Drawing.Point(45, 53);
            this.chkSzenSzen.Name = "chkSzenSzen";
            this.chkSzenSzen.Size = new System.Drawing.Size(99, 22);
            this.chkSzenSzen.TabIndex = 1;
            this.chkSzenSzen.Text = "修繕見積情報";
            this.chkSzenSzen.UseVisualStyleBackColor = true;
            // 
            // chkSzenSzenmeisai
            // 
            this.chkSzenSzenmeisai.AutoSize = true;
            this.chkSzenSzenmeisai.Location = new System.Drawing.Point(45, 81);
            this.chkSzenSzenmeisai.Name = "chkSzenSzenmeisai";
            this.chkSzenSzenmeisai.Size = new System.Drawing.Size(123, 22);
            this.chkSzenSzenmeisai.TabIndex = 2;
            this.chkSzenSzenmeisai.Text = "修繕見積詳細情報";
            this.chkSzenSzenmeisai.UseVisualStyleBackColor = true;
            // 
            // chkSzenBase
            // 
            this.chkSzenBase.AutoSize = true;
            this.chkSzenBase.Location = new System.Drawing.Point(25, 25);
            this.chkSzenBase.Name = "chkSzenBase";
            this.chkSzenBase.Size = new System.Drawing.Size(99, 22);
            this.chkSzenBase.TabIndex = 0;
            this.chkSzenBase.Text = "修繕基本情報";
            this.chkSzenBase.UseVisualStyleBackColor = true;
            // 
            // grpClaim
            // 
            this.grpClaim.Controls.Add(this.chkClaimTaiorireki);
            this.grpClaim.Controls.Add(this.chkClaimRelfile);
            this.grpClaim.Controls.Add(this.chkClaimBase);
            this.grpClaim.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpClaim.Location = new System.Drawing.Point(16, 8);
            this.grpClaim.Name = "grpClaim";
            this.grpClaim.Size = new System.Drawing.Size(282, 265);
            this.grpClaim.TabIndex = 1;
            this.grpClaim.TabStop = false;
            this.grpClaim.Text = "クレーム情報";
            // 
            // chkClaimTaiorireki
            // 
            this.chkClaimTaiorireki.AutoSize = true;
            this.chkClaimTaiorireki.Location = new System.Drawing.Point(45, 53);
            this.chkClaimTaiorireki.Name = "chkClaimTaiorireki";
            this.chkClaimTaiorireki.Size = new System.Drawing.Size(147, 22);
            this.chkClaimTaiorireki.TabIndex = 1;
            this.chkClaimTaiorireki.Text = "クレーム対応履歴情報";
            this.chkClaimTaiorireki.UseVisualStyleBackColor = true;
            // 
            // chkClaimRelfile
            // 
            this.chkClaimRelfile.AutoSize = true;
            this.chkClaimRelfile.Location = new System.Drawing.Point(45, 81);
            this.chkClaimRelfile.Name = "chkClaimRelfile";
            this.chkClaimRelfile.Size = new System.Drawing.Size(171, 22);
            this.chkClaimRelfile.TabIndex = 2;
            this.chkClaimRelfile.Text = "クレーム関連ファイル情報";
            this.chkClaimRelfile.UseVisualStyleBackColor = true;
            // 
            // chkClaimBase
            // 
            this.chkClaimBase.AutoSize = true;
            this.chkClaimBase.Location = new System.Drawing.Point(25, 25);
            this.chkClaimBase.Name = "chkClaimBase";
            this.chkClaimBase.Size = new System.Drawing.Size(123, 22);
            this.chkClaimBase.TabIndex = 0;
            this.chkClaimBase.Text = "クレーム基本情報";
            this.chkClaimBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase200
            // 
            this.tabPageBase200.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase200.Controls.Add(this.grpSyskanri);
            this.tabPageBase200.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase200.Name = "tabPageBase200";
            this.tabPageBase200.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase200.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase200.TabIndex = 19;
            this.tabPageBase200.Text = "初期設定";
            // 
            // grpSyskanri
            // 
            this.grpSyskanri.Controls.Add(this.chkSyskanriNkinkomkmerge);
            this.grpSyskanri.Controls.Add(this.chkSyskanriHenkanmoji);
            this.grpSyskanri.Controls.Add(this.chkSyskanriZei);
            this.grpSyskanri.Controls.Add(this.chkSyskanriBase);
            this.grpSyskanri.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpSyskanri.Location = new System.Drawing.Point(24, 16);
            this.grpSyskanri.Name = "grpSyskanri";
            this.grpSyskanri.Size = new System.Drawing.Size(237, 240);
            this.grpSyskanri.TabIndex = 4;
            this.grpSyskanri.TabStop = false;
            this.grpSyskanri.Text = "初期設定";
            // 
            // chkSyskanriNkinkomkmerge
            // 
            this.chkSyskanriNkinkomkmerge.AutoSize = true;
            this.chkSyskanriNkinkomkmerge.Location = new System.Drawing.Point(32, 120);
            this.chkSyskanriNkinkomkmerge.Name = "chkSyskanriNkinkomkmerge";
            this.chkSyskanriNkinkomkmerge.Size = new System.Drawing.Size(123, 22);
            this.chkSyskanriNkinkomkmerge.TabIndex = 3;
            this.chkSyskanriNkinkomkmerge.Text = "入金項目集約情報";
            this.chkSyskanriNkinkomkmerge.UseVisualStyleBackColor = true;
            // 
            // chkSyskanriHenkanmoji
            // 
            this.chkSyskanriHenkanmoji.AutoSize = true;
            this.chkSyskanriHenkanmoji.Location = new System.Drawing.Point(32, 88);
            this.chkSyskanriHenkanmoji.Name = "chkSyskanriHenkanmoji";
            this.chkSyskanriHenkanmoji.Size = new System.Drawing.Size(99, 22);
            this.chkSyskanriHenkanmoji.TabIndex = 2;
            this.chkSyskanriHenkanmoji.Text = "変換文字情報";
            this.chkSyskanriHenkanmoji.UseVisualStyleBackColor = true;
            // 
            // chkSyskanriZei
            // 
            this.chkSyskanriZei.AutoSize = true;
            this.chkSyskanriZei.Location = new System.Drawing.Point(32, 56);
            this.chkSyskanriZei.Name = "chkSyskanriZei";
            this.chkSyskanriZei.Size = new System.Drawing.Size(87, 22);
            this.chkSyskanriZei.TabIndex = 1;
            this.chkSyskanriZei.Text = "税編集情報";
            this.chkSyskanriZei.UseVisualStyleBackColor = true;
            // 
            // chkSyskanriBase
            // 
            this.chkSyskanriBase.AutoSize = true;
            this.chkSyskanriBase.Location = new System.Drawing.Point(25, 25);
            this.chkSyskanriBase.Name = "chkSyskanriBase";
            this.chkSyskanriBase.Size = new System.Drawing.Size(123, 22);
            this.chkSyskanriBase.TabIndex = 0;
            this.chkSyskanriBase.Text = "初期設定基本情報";
            this.chkSyskanriBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase210
            // 
            this.tabPageBase210.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase210.Controls.Add(this.grpRendo);
            this.tabPageBase210.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase210.Name = "tabPageBase210";
            this.tabPageBase210.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase210.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase210.TabIndex = 21;
            this.tabPageBase210.Text = "物件データ連動情報";
            // 
            // grpRendo
            // 
            this.grpRendo.Controls.Add(this.chkRendoMapdisp);
            this.grpRendo.Controls.Add(this.chkRendoBtoBgroup);
            this.grpRendo.Controls.Add(this.chkRendoHysosin);
            this.grpRendo.Controls.Add(this.chkRendoHyrui);
            this.grpRendo.Controls.Add(this.chkRendoKokokuSuumo);
            this.grpRendo.Controls.Add(this.chkRendoKokokuAthome);
            this.grpRendo.Controls.Add(this.chkRendoKokokuHomes);
            this.grpRendo.Controls.Add(this.chkRendoKokokuJisyaweb);
            this.grpRendo.Controls.Add(this.chkRendoSosinSuumo);
            this.grpRendo.Controls.Add(this.chkRendoSosinAthome);
            this.grpRendo.Controls.Add(this.chkRendoSosinHomes);
            this.grpRendo.Controls.Add(this.chkRendoSosinJisyaweb);
            this.grpRendo.Controls.Add(this.chkRendoSosinBase);
            this.grpRendo.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpRendo.Location = new System.Drawing.Point(16, 16);
            this.grpRendo.Name = "grpRendo";
            this.grpRendo.Size = new System.Drawing.Size(567, 240);
            this.grpRendo.TabIndex = 7;
            this.grpRendo.TabStop = false;
            this.grpRendo.Text = "物件データ連動情報";
            // 
            // chkRendoMapdisp
            // 
            this.chkRendoMapdisp.AutoSize = true;
            this.chkRendoMapdisp.Location = new System.Drawing.Point(360, 137);
            this.chkRendoMapdisp.Name = "chkRendoMapdisp";
            this.chkRendoMapdisp.Size = new System.Drawing.Size(147, 22);
            this.chkRendoMapdisp.TabIndex = 15;
            this.chkRendoMapdisp.Text = "地図表示詳細設定情報";
            this.chkRendoMapdisp.UseVisualStyleBackColor = true;
            // 
            // chkRendoBtoBgroup
            // 
            this.chkRendoBtoBgroup.AutoSize = true;
            this.chkRendoBtoBgroup.Location = new System.Drawing.Point(360, 109);
            this.chkRendoBtoBgroup.Name = "chkRendoBtoBgroup";
            this.chkRendoBtoBgroup.Size = new System.Drawing.Size(151, 22);
            this.chkRendoBtoBgroup.TabIndex = 14;
            this.chkRendoBtoBgroup.Text = "BtoBグループ設定情報";
            this.chkRendoBtoBgroup.UseVisualStyleBackColor = true;
            // 
            // chkRendoHysosin
            // 
            this.chkRendoHysosin.AutoSize = true;
            this.chkRendoHysosin.Location = new System.Drawing.Point(360, 81);
            this.chkRendoHysosin.Name = "chkRendoHysosin";
            this.chkRendoHysosin.Size = new System.Drawing.Size(111, 22);
            this.chkRendoHysosin.TabIndex = 13;
            this.chkRendoHysosin.Text = "部屋毎送信情報";
            this.chkRendoHysosin.UseVisualStyleBackColor = true;
            // 
            // chkRendoHyrui
            // 
            this.chkRendoHyrui.AutoSize = true;
            this.chkRendoHyrui.Location = new System.Drawing.Point(360, 53);
            this.chkRendoHyrui.Name = "chkRendoHyrui";
            this.chkRendoHyrui.Size = new System.Drawing.Size(171, 22);
            this.chkRendoHyrui.TabIndex = 12;
            this.chkRendoHyrui.Text = "ポータル連動部屋分類情報";
            this.chkRendoHyrui.UseVisualStyleBackColor = true;
            // 
            // chkRendoKokokuSuumo
            // 
            this.chkRendoKokokuSuumo.AutoSize = true;
            this.chkRendoKokokuSuumo.Location = new System.Drawing.Point(188, 137);
            this.chkRendoKokokuSuumo.Name = "chkRendoKokokuSuumo";
            this.chkRendoKokokuSuumo.Size = new System.Drawing.Size(144, 22);
            this.chkRendoKokokuSuumo.TabIndex = 11;
            this.chkRendoKokokuSuumo.Text = "広告補足SUUMO情報";
            this.chkRendoKokokuSuumo.UseVisualStyleBackColor = true;
            // 
            // chkRendoKokokuAthome
            // 
            this.chkRendoKokokuAthome.AutoSize = true;
            this.chkRendoKokokuAthome.Location = new System.Drawing.Point(188, 109);
            this.chkRendoKokokuAthome.Name = "chkRendoKokokuAthome";
            this.chkRendoKokokuAthome.Size = new System.Drawing.Size(144, 22);
            this.chkRendoKokokuAthome.TabIndex = 10;
            this.chkRendoKokokuAthome.Text = "広告補足athome情報";
            this.chkRendoKokokuAthome.UseVisualStyleBackColor = true;
            // 
            // chkRendoKokokuHomes
            // 
            this.chkRendoKokokuHomes.AutoSize = true;
            this.chkRendoKokokuHomes.Location = new System.Drawing.Point(188, 81);
            this.chkRendoKokokuHomes.Name = "chkRendoKokokuHomes";
            this.chkRendoKokokuHomes.Size = new System.Drawing.Size(142, 22);
            this.chkRendoKokokuHomes.TabIndex = 9;
            this.chkRendoKokokuHomes.Text = "広告補足HOMES情報";
            this.chkRendoKokokuHomes.UseVisualStyleBackColor = true;
            // 
            // chkRendoKokokuJisyaweb
            // 
            this.chkRendoKokokuJisyaweb.AutoSize = true;
            this.chkRendoKokokuJisyaweb.Location = new System.Drawing.Point(188, 53);
            this.chkRendoKokokuJisyaweb.Name = "chkRendoKokokuJisyaweb";
            this.chkRendoKokokuJisyaweb.Size = new System.Drawing.Size(147, 22);
            this.chkRendoKokokuJisyaweb.TabIndex = 8;
            this.chkRendoKokokuJisyaweb.Text = "広告補足自社web情報";
            this.chkRendoKokokuJisyaweb.UseVisualStyleBackColor = true;
            // 
            // chkRendoSosinSuumo
            // 
            this.chkRendoSosinSuumo.AutoSize = true;
            this.chkRendoSosinSuumo.Location = new System.Drawing.Point(25, 137);
            this.chkRendoSosinSuumo.Name = "chkRendoSosinSuumo";
            this.chkRendoSosinSuumo.Size = new System.Drawing.Size(144, 22);
            this.chkRendoSosinSuumo.TabIndex = 7;
            this.chkRendoSosinSuumo.Text = "送信設定SUUMO情報";
            this.chkRendoSosinSuumo.UseVisualStyleBackColor = true;
            // 
            // chkRendoSosinAthome
            // 
            this.chkRendoSosinAthome.AutoSize = true;
            this.chkRendoSosinAthome.Location = new System.Drawing.Point(25, 109);
            this.chkRendoSosinAthome.Name = "chkRendoSosinAthome";
            this.chkRendoSosinAthome.Size = new System.Drawing.Size(144, 22);
            this.chkRendoSosinAthome.TabIndex = 6;
            this.chkRendoSosinAthome.Text = "送信設定athome情報";
            this.chkRendoSosinAthome.UseVisualStyleBackColor = true;
            // 
            // chkRendoSosinHomes
            // 
            this.chkRendoSosinHomes.AutoSize = true;
            this.chkRendoSosinHomes.Location = new System.Drawing.Point(25, 81);
            this.chkRendoSosinHomes.Name = "chkRendoSosinHomes";
            this.chkRendoSosinHomes.Size = new System.Drawing.Size(142, 22);
            this.chkRendoSosinHomes.TabIndex = 5;
            this.chkRendoSosinHomes.Text = "送信設定HOMES情報";
            this.chkRendoSosinHomes.UseVisualStyleBackColor = true;
            // 
            // chkRendoSosinJisyaweb
            // 
            this.chkRendoSosinJisyaweb.AutoSize = true;
            this.chkRendoSosinJisyaweb.Location = new System.Drawing.Point(25, 53);
            this.chkRendoSosinJisyaweb.Name = "chkRendoSosinJisyaweb";
            this.chkRendoSosinJisyaweb.Size = new System.Drawing.Size(147, 22);
            this.chkRendoSosinJisyaweb.TabIndex = 1;
            this.chkRendoSosinJisyaweb.Text = "送信設定自社web情報";
            this.chkRendoSosinJisyaweb.UseVisualStyleBackColor = true;
            // 
            // chkRendoSosinBase
            // 
            this.chkRendoSosinBase.AutoSize = true;
            this.chkRendoSosinBase.Location = new System.Drawing.Point(25, 25);
            this.chkRendoSosinBase.Name = "chkRendoSosinBase";
            this.chkRendoSosinBase.Size = new System.Drawing.Size(123, 22);
            this.chkRendoSosinBase.TabIndex = 0;
            this.chkRendoSosinBase.Text = "送信設定基本情報";
            this.chkRendoSosinBase.UseVisualStyleBackColor = true;
            // 
            // tabPageBase900
            // 
            this.tabPageBase900.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageBase900.Controls.Add(this.Label22);
            this.tabPageBase900.Location = new System.Drawing.Point(4, 27);
            this.tabPageBase900.Name = "tabPageBase900";
            this.tabPageBase900.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBase900.Size = new System.Drawing.Size(847, 327);
            this.tabPageBase900.TabIndex = 8;
            this.tabPageBase900.Text = " -";
            // 
            // Label22
            // 
            this.Label22.AutoSize = true;
            this.Label22.Font = new System.Drawing.Font("メイリオ", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label22.ForeColor = System.Drawing.Color.Red;
            this.Label22.Location = new System.Drawing.Point(307, 92);
            this.Label22.Name = "Label22";
            this.Label22.Size = new System.Drawing.Size(232, 96);
            this.Label22.TabIndex = 0;
            this.Label22.Text = "作成中";
            // 
            // tabPageHanyo110
            // 
            this.tabPageHanyo110.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHanyo110.Controls.Add(this.grpHMst);
            this.tabPageHanyo110.Location = new System.Drawing.Point(4, 27);
            this.tabPageHanyo110.Name = "tabPageHanyo110";
            this.tabPageHanyo110.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageHanyo110.Size = new System.Drawing.Size(847, 327);
            this.tabPageHanyo110.TabIndex = 12;
            this.tabPageHanyo110.Text = " 没選択1(汎用)";
            // 
            // grpHMst
            // 
            this.grpHMst.Controls.Add(this.Label138);
            this.grpHMst.Controls.Add(this.Label150);
            this.grpHMst.Controls.Add(this.Label151);
            this.grpHMst.Controls.Add(this.Label152);
            this.grpHMst.Controls.Add(this.Label153);
            this.grpHMst.Controls.Add(this.Label154);
            this.grpHMst.Controls.Add(this.Label155);
            this.grpHMst.Controls.Add(this.Label156);
            this.grpHMst.Controls.Add(this.Label157);
            this.grpHMst.Controls.Add(this.CheckBox2);
            this.grpHMst.Controls.Add(this.CheckBox3);
            this.grpHMst.Controls.Add(this.CheckBox6);
            this.grpHMst.Controls.Add(this.CheckBox9);
            this.grpHMst.Controls.Add(this.CheckBox10);
            this.grpHMst.Controls.Add(this.CheckBox11);
            this.grpHMst.Controls.Add(this.CheckBox12);
            this.grpHMst.Controls.Add(this.CheckBox13);
            this.grpHMst.Controls.Add(this.CheckBox17);
            this.grpHMst.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpHMst.Location = new System.Drawing.Point(10, 5);
            this.grpHMst.Name = "grpHMst";
            this.grpHMst.Size = new System.Drawing.Size(814, 265);
            this.grpHMst.TabIndex = 0;
            this.grpHMst.TabStop = false;
            this.grpHMst.Text = " 各マスタ情報 ";
            // 
            // Label138
            // 
            this.Label138.AutoSize = true;
            this.Label138.ForeColor = System.Drawing.Color.Black;
            this.Label138.Location = new System.Drawing.Point(222, 188);
            this.Label138.Name = "Label138";
            this.Label138.Size = new System.Drawing.Size(119, 18);
            this.Label138.TabIndex = 17;
            this.Label138.Text = "   (例:オートロック)";
            // 
            // Label150
            // 
            this.Label150.AutoSize = true;
            this.Label150.ForeColor = System.Drawing.Color.Black;
            this.Label150.Location = new System.Drawing.Point(222, 93);
            this.Label150.Name = "Label150";
            this.Label150.Size = new System.Drawing.Size(170, 18);
            this.Label150.TabIndex = 13;
            this.Label150.Text = "   (例:建物(専用部)-クレーム)";
            // 
            // Label151
            // 
            this.Label151.AutoSize = true;
            this.Label151.ForeColor = System.Drawing.Color.Black;
            this.Label151.Location = new System.Drawing.Point(222, 50);
            this.Label151.Name = "Label151";
            this.Label151.Size = new System.Drawing.Size(131, 18);
            this.Label151.TabIndex = 11;
            this.Label151.Text = "   (例:第一種住宅地域)";
            // 
            // Label152
            // 
            this.Label152.AutoSize = true;
            this.Label152.ForeColor = System.Drawing.Color.Black;
            this.Label152.Location = new System.Drawing.Point(222, 142);
            this.Label152.Name = "Label152";
            this.Label152.Size = new System.Drawing.Size(162, 18);
            this.Label152.TabIndex = 15;
            this.Label152.Text = "   (口径や料金表などの設定)";
            // 
            // Label153
            // 
            this.Label153.AutoSize = true;
            this.Label153.ForeColor = System.Drawing.Color.Black;
            this.Label153.Location = new System.Drawing.Point(22, 234);
            this.Label153.Name = "Label153";
            this.Label153.Size = new System.Drawing.Size(155, 18);
            this.Label153.TabIndex = 9;
            this.Label153.Text = "   (例:定期借地借家権契約)";
            // 
            // Label154
            // 
            this.Label154.AutoSize = true;
            this.Label154.ForeColor = System.Drawing.Color.Black;
            this.Label154.Location = new System.Drawing.Point(22, 188);
            this.Label154.Name = "Label154";
            this.Label154.Size = new System.Drawing.Size(119, 18);
            this.Label154.TabIndex = 7;
            this.Label154.Text = "   (例:住宅総合保険)";
            // 
            // Label155
            // 
            this.Label155.AutoSize = true;
            this.Label155.ForeColor = System.Drawing.Color.Black;
            this.Label155.Location = new System.Drawing.Point(22, 142);
            this.Label155.Name = "Label155";
            this.Label155.Size = new System.Drawing.Size(131, 18);
            this.Label155.TabIndex = 5;
            this.Label155.Text = "   (例:東部、東エリア)";
            // 
            // Label156
            // 
            this.Label156.AutoSize = true;
            this.Label156.ForeColor = System.Drawing.Color.Black;
            this.Label156.Location = new System.Drawing.Point(22, 96);
            this.Label156.Name = "Label156";
            this.Label156.Size = new System.Drawing.Size(155, 18);
            this.Label156.TabIndex = 3;
            this.Label156.Text = "   (例:都城市立明道小学校)";
            // 
            // Label157
            // 
            this.Label157.AutoSize = true;
            this.Label157.ForeColor = System.Drawing.Color.Black;
            this.Label157.Location = new System.Drawing.Point(22, 50);
            this.Label157.Name = "Label157";
            this.Label157.Size = new System.Drawing.Size(184, 18);
            this.Label157.TabIndex = 1;
            this.Label157.Text = "   (例:宮崎交通-バスセンター前)";
            // 
            // CheckBox2
            // 
            this.CheckBox2.AutoSize = true;
            this.CheckBox2.Location = new System.Drawing.Point(225, 117);
            this.CheckBox2.Name = "CheckBox2";
            this.CheckBox2.Size = new System.Drawing.Size(111, 22);
            this.CheckBox2.TabIndex = 14;
            this.CheckBox2.Text = "変動費設定内容";
            this.CheckBox2.UseVisualStyleBackColor = true;
            // 
            // CheckBox3
            // 
            this.CheckBox3.AutoSize = true;
            this.CheckBox3.Location = new System.Drawing.Point(225, 71);
            this.CheckBox3.Name = "CheckBox3";
            this.CheckBox3.Size = new System.Drawing.Size(147, 22);
            this.CheckBox3.TabIndex = 12;
            this.CheckBox3.Text = "クレーム分類設定内容";
            this.CheckBox3.UseVisualStyleBackColor = true;
            // 
            // CheckBox6
            // 
            this.CheckBox6.AutoSize = true;
            this.CheckBox6.Location = new System.Drawing.Point(25, 25);
            this.CheckBox6.Name = "CheckBox6";
            this.CheckBox6.Size = new System.Drawing.Size(111, 22);
            this.CheckBox6.TabIndex = 0;
            this.CheckBox6.Text = "バス交通マスタ";
            this.CheckBox6.UseVisualStyleBackColor = true;
            // 
            // CheckBox9
            // 
            this.CheckBox9.AutoSize = true;
            this.CheckBox9.Location = new System.Drawing.Point(25, 71);
            this.CheckBox9.Name = "CheckBox9";
            this.CheckBox9.Size = new System.Drawing.Size(99, 22);
            this.CheckBox9.TabIndex = 2;
            this.CheckBox9.Text = "学校区マスタ";
            this.CheckBox9.UseVisualStyleBackColor = true;
            // 
            // CheckBox10
            // 
            this.CheckBox10.AutoSize = true;
            this.CheckBox10.Location = new System.Drawing.Point(225, 25);
            this.CheckBox10.Name = "CheckBox10";
            this.CheckBox10.Size = new System.Drawing.Size(87, 22);
            this.CheckBox10.TabIndex = 10;
            this.CheckBox10.Text = "特約マスタ";
            this.CheckBox10.UseVisualStyleBackColor = true;
            // 
            // CheckBox11
            // 
            this.CheckBox11.AutoSize = true;
            this.CheckBox11.Location = new System.Drawing.Point(25, 209);
            this.CheckBox11.Name = "CheckBox11";
            this.CheckBox11.Size = new System.Drawing.Size(111, 22);
            this.CheckBox11.TabIndex = 8;
            this.CheckBox11.Text = "契約分類マスタ";
            this.CheckBox11.UseVisualStyleBackColor = true;
            // 
            // CheckBox12
            // 
            this.CheckBox12.AutoSize = true;
            this.CheckBox12.Location = new System.Drawing.Point(25, 163);
            this.CheckBox12.Name = "CheckBox12";
            this.CheckBox12.Size = new System.Drawing.Size(111, 22);
            this.CheckBox12.TabIndex = 6;
            this.CheckBox12.Text = "保険種類マスタ";
            this.CheckBox12.UseVisualStyleBackColor = true;
            // 
            // CheckBox13
            // 
            this.CheckBox13.AutoSize = true;
            this.CheckBox13.Location = new System.Drawing.Point(225, 163);
            this.CheckBox13.Name = "CheckBox13";
            this.CheckBox13.Size = new System.Drawing.Size(123, 22);
            this.CheckBox13.TabIndex = 16;
            this.CheckBox13.Text = "鍵タイトルマスタ";
            this.CheckBox13.UseVisualStyleBackColor = true;
            // 
            // CheckBox17
            // 
            this.CheckBox17.AutoSize = true;
            this.CheckBox17.Location = new System.Drawing.Point(25, 117);
            this.CheckBox17.Name = "CheckBox17";
            this.CheckBox17.Size = new System.Drawing.Size(99, 22);
            this.CheckBox17.TabIndex = 4;
            this.CheckBox17.Text = "エリアマスタ";
            this.CheckBox17.UseVisualStyleBackColor = true;
            // 
            // tabPageHanyo120
            // 
            this.tabPageHanyo120.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHanyo120.Controls.Add(this.grpHGy);
            this.tabPageHanyo120.Controls.Add(this.grpHOw);
            this.tabPageHanyo120.Controls.Add(this.grpHJisya);
            this.tabPageHanyo120.Location = new System.Drawing.Point(4, 27);
            this.tabPageHanyo120.Name = "tabPageHanyo120";
            this.tabPageHanyo120.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageHanyo120.Size = new System.Drawing.Size(847, 327);
            this.tabPageHanyo120.TabIndex = 13;
            this.tabPageHanyo120.Text = " 没選択2(汎用)";
            // 
            // grpHGy
            // 
            this.grpHGy.Controls.Add(this.CheckBox4);
            this.grpHGy.Controls.Add(this.CheckBox30);
            this.grpHGy.Controls.Add(this.CheckBox31);
            this.grpHGy.Controls.Add(this.CheckBox32);
            this.grpHGy.Controls.Add(this.CheckBox33);
            this.grpHGy.Controls.Add(this.CheckBox34);
            this.grpHGy.Controls.Add(this.CheckBox35);
            this.grpHGy.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpHGy.Location = new System.Drawing.Point(586, 6);
            this.grpHGy.Name = "grpHGy";
            this.grpHGy.Size = new System.Drawing.Size(245, 265);
            this.grpHGy.TabIndex = 2;
            this.grpHGy.TabStop = false;
            this.grpHGy.Text = " 業者情報 ";
            // 
            // CheckBox4
            // 
            this.CheckBox4.AutoSize = true;
            this.CheckBox4.Location = new System.Drawing.Point(25, 192);
            this.CheckBox4.Name = "CheckBox4";
            this.CheckBox4.Size = new System.Drawing.Size(99, 22);
            this.CheckBox4.TabIndex = 6;
            this.CheckBox4.Text = "施工業者情報";
            this.CheckBox4.UseVisualStyleBackColor = true;
            // 
            // CheckBox30
            // 
            this.CheckBox30.AutoSize = true;
            this.CheckBox30.Location = new System.Drawing.Point(25, 164);
            this.CheckBox30.Name = "CheckBox30";
            this.CheckBox30.Size = new System.Drawing.Size(123, 22);
            this.CheckBox30.TabIndex = 5;
            this.CheckBox30.Text = "施設保守業者情報";
            this.CheckBox30.UseVisualStyleBackColor = true;
            // 
            // CheckBox31
            // 
            this.CheckBox31.AutoSize = true;
            this.CheckBox31.Location = new System.Drawing.Point(25, 52);
            this.CheckBox31.Name = "CheckBox31";
            this.CheckBox31.Size = new System.Drawing.Size(99, 22);
            this.CheckBox31.TabIndex = 1;
            this.CheckBox31.Text = "修繕業者情報";
            this.CheckBox31.UseVisualStyleBackColor = true;
            // 
            // CheckBox32
            // 
            this.CheckBox32.AutoSize = true;
            this.CheckBox32.Location = new System.Drawing.Point(25, 136);
            this.CheckBox32.Name = "CheckBox32";
            this.CheckBox32.Size = new System.Drawing.Size(123, 22);
            this.CheckBox32.TabIndex = 4;
            this.CheckBox32.Text = "家賃保証業者情報";
            this.CheckBox32.UseVisualStyleBackColor = true;
            // 
            // CheckBox33
            // 
            this.CheckBox33.AutoSize = true;
            this.CheckBox33.Location = new System.Drawing.Point(25, 108);
            this.CheckBox33.Name = "CheckBox33";
            this.CheckBox33.Size = new System.Drawing.Size(99, 22);
            this.CheckBox33.TabIndex = 3;
            this.CheckBox33.Text = "保険業者情報";
            this.CheckBox33.UseVisualStyleBackColor = true;
            // 
            // CheckBox34
            // 
            this.CheckBox34.AutoSize = true;
            this.CheckBox34.Location = new System.Drawing.Point(25, 80);
            this.CheckBox34.Name = "CheckBox34";
            this.CheckBox34.Size = new System.Drawing.Size(147, 22);
            this.CheckBox34.TabIndex = 2;
            this.CheckBox34.Text = "ライフライン業者情報";
            this.CheckBox34.UseVisualStyleBackColor = true;
            // 
            // CheckBox35
            // 
            this.CheckBox35.AutoSize = true;
            this.CheckBox35.Location = new System.Drawing.Point(25, 25);
            this.CheckBox35.Name = "CheckBox35";
            this.CheckBox35.Size = new System.Drawing.Size(99, 22);
            this.CheckBox35.TabIndex = 0;
            this.CheckBox35.Text = "仲介業者情報";
            this.CheckBox35.UseVisualStyleBackColor = true;
            // 
            // grpHOw
            // 
            this.grpHOw.Controls.Add(this.CheckBox1);
            this.grpHOw.Controls.Add(this.CheckBox26);
            this.grpHOw.Controls.Add(this.CheckBox27);
            this.grpHOw.Controls.Add(this.CheckBox28);
            this.grpHOw.Controls.Add(this.CheckBox29);
            this.grpHOw.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpHOw.Location = new System.Drawing.Point(269, 6);
            this.grpHOw.Name = "grpHOw";
            this.grpHOw.Size = new System.Drawing.Size(311, 265);
            this.grpHOw.TabIndex = 1;
            this.grpHOw.TabStop = false;
            this.grpHOw.Text = "家主情報";
            // 
            // CheckBox1
            // 
            this.CheckBox1.AutoSize = true;
            this.CheckBox1.Location = new System.Drawing.Point(45, 136);
            this.CheckBox1.Name = "CheckBox1";
            this.CheckBox1.Size = new System.Drawing.Size(217, 22);
            this.CheckBox1.TabIndex = 4;
            this.CheckBox1.Text = "(未使用の家主データは移行しない)";
            this.CheckBox1.UseVisualStyleBackColor = true;
            // 
            // CheckBox26
            // 
            this.CheckBox26.AutoSize = true;
            this.CheckBox26.Location = new System.Drawing.Point(45, 80);
            this.CheckBox26.Name = "CheckBox26";
            this.CheckBox26.Size = new System.Drawing.Size(123, 22);
            this.CheckBox26.TabIndex = 2;
            this.CheckBox26.Text = "家主イベント情報";
            this.CheckBox26.UseVisualStyleBackColor = true;
            // 
            // CheckBox27
            // 
            this.CheckBox27.AutoSize = true;
            this.CheckBox27.Location = new System.Drawing.Point(45, 53);
            this.CheckBox27.Name = "CheckBox27";
            this.CheckBox27.Size = new System.Drawing.Size(99, 22);
            this.CheckBox27.TabIndex = 1;
            this.CheckBox27.Text = "家主口座情報";
            this.CheckBox27.UseVisualStyleBackColor = true;
            // 
            // CheckBox28
            // 
            this.CheckBox28.AutoSize = true;
            this.CheckBox28.Location = new System.Drawing.Point(45, 108);
            this.CheckBox28.Name = "CheckBox28";
            this.CheckBox28.Size = new System.Drawing.Size(99, 22);
            this.CheckBox28.TabIndex = 3;
            this.CheckBox28.Text = "家主メモ情報";
            this.CheckBox28.UseVisualStyleBackColor = true;
            // 
            // CheckBox29
            // 
            this.CheckBox29.AutoSize = true;
            this.CheckBox29.Location = new System.Drawing.Point(25, 25);
            this.CheckBox29.Name = "CheckBox29";
            this.CheckBox29.Size = new System.Drawing.Size(99, 22);
            this.CheckBox29.TabIndex = 0;
            this.CheckBox29.Text = "家主基本情報";
            this.CheckBox29.UseVisualStyleBackColor = true;
            // 
            // grpHJisya
            // 
            this.grpHJisya.Controls.Add(this.CheckBox19);
            this.grpHJisya.Controls.Add(this.CheckBox20);
            this.grpHJisya.Controls.Add(this.CheckBox21);
            this.grpHJisya.Controls.Add(this.CheckBox22);
            this.grpHJisya.Controls.Add(this.CheckBox25);
            this.grpHJisya.Font = new System.Drawing.Font("メイリオ", 9F);
            this.grpHJisya.Location = new System.Drawing.Point(10, 5);
            this.grpHJisya.Name = "grpHJisya";
            this.grpHJisya.Size = new System.Drawing.Size(253, 265);
            this.grpHJisya.TabIndex = 0;
            this.grpHJisya.TabStop = false;
            this.grpHJisya.Text = " 自社情報 ";
            // 
            // CheckBox19
            // 
            this.CheckBox19.AutoSize = true;
            this.CheckBox19.Location = new System.Drawing.Point(45, 109);
            this.CheckBox19.Name = "CheckBox19";
            this.CheckBox19.Size = new System.Drawing.Size(99, 22);
            this.CheckBox19.TabIndex = 3;
            this.CheckBox19.Text = "口座振替情報";
            this.CheckBox19.UseVisualStyleBackColor = true;
            // 
            // CheckBox20
            // 
            this.CheckBox20.AutoSize = true;
            this.CheckBox20.Location = new System.Drawing.Point(44, 81);
            this.CheckBox20.Name = "CheckBox20";
            this.CheckBox20.Size = new System.Drawing.Size(111, 22);
            this.CheckBox20.TabIndex = 2;
            this.CheckBox20.Text = "振込依頼人情報";
            this.CheckBox20.UseVisualStyleBackColor = true;
            // 
            // CheckBox21
            // 
            this.CheckBox21.AutoSize = true;
            this.CheckBox21.Location = new System.Drawing.Point(45, 53);
            this.CheckBox21.Name = "CheckBox21";
            this.CheckBox21.Size = new System.Drawing.Size(99, 22);
            this.CheckBox21.TabIndex = 1;
            this.CheckBox21.Text = "自社口座情報";
            this.CheckBox21.UseVisualStyleBackColor = true;
            // 
            // CheckBox22
            // 
            this.CheckBox22.AutoSize = true;
            this.CheckBox22.Location = new System.Drawing.Point(44, 137);
            this.CheckBox22.Name = "CheckBox22";
            this.CheckBox22.Size = new System.Drawing.Size(99, 22);
            this.CheckBox22.TabIndex = 4;
            this.CheckBox22.Text = "自社メモ情報";
            this.CheckBox22.UseVisualStyleBackColor = true;
            // 
            // CheckBox25
            // 
            this.CheckBox25.AutoSize = true;
            this.CheckBox25.Location = new System.Drawing.Point(25, 25);
            this.CheckBox25.Name = "CheckBox25";
            this.CheckBox25.Size = new System.Drawing.Size(99, 22);
            this.CheckBox25.TabIndex = 0;
            this.CheckBox25.Text = "自社基本情報";
            this.CheckBox25.UseVisualStyleBackColor = true;
            // 
            // tabPageHanyo130
            // 
            this.tabPageHanyo130.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHanyo130.Location = new System.Drawing.Point(4, 27);
            this.tabPageHanyo130.Name = "tabPageHanyo130";
            this.tabPageHanyo130.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageHanyo130.Size = new System.Drawing.Size(847, 327);
            this.tabPageHanyo130.TabIndex = 17;
            this.tabPageHanyo130.Text = " 没選択3(汎用)";
            // 
            // tabPageHanyo140
            // 
            this.tabPageHanyo140.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHanyo140.Location = new System.Drawing.Point(4, 27);
            this.tabPageHanyo140.Name = "tabPageHanyo140";
            this.tabPageHanyo140.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageHanyo140.Size = new System.Drawing.Size(847, 327);
            this.tabPageHanyo140.TabIndex = 16;
            this.tabPageHanyo140.Text = " 没選択4(汎用)";
            // 
            // tabPageHanyo150
            // 
            this.tabPageHanyo150.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHanyo150.Location = new System.Drawing.Point(4, 27);
            this.tabPageHanyo150.Name = "tabPageHanyo150";
            this.tabPageHanyo150.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageHanyo150.Size = new System.Drawing.Size(847, 327);
            this.tabPageHanyo150.TabIndex = 15;
            this.tabPageHanyo150.Text = " 没選択5(汎用)";
            // 
            // TabPage1
            // 
            this.TabPage1.BackColor = System.Drawing.SystemColors.Menu;
            this.TabPage1.Controls.Add(this.grpKizon4);
            this.TabPage1.Location = new System.Drawing.Point(4, 27);
            this.TabPage1.Name = "TabPage1";
            this.TabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.TabPage1.Size = new System.Drawing.Size(847, 327);
            this.TabPage1.TabIndex = 20;
            this.TabPage1.Text = "TabPage1";
            // 
            // grpKizon4
            // 
            this.grpKizon4.Controls.Add(this.lblSelectPageCnt3);
            this.grpKizon4.Controls.Add(this.Label26);
            this.grpKizon4.Controls.Add(this.Label73);
            this.grpKizon4.Controls.Add(this.Label70);
            this.grpKizon4.Controls.Add(this.Label71);
            this.grpKizon4.Controls.Add(this.Label72);
            this.grpKizon4.Controls.Add(this.Label74);
            this.grpKizon4.Controls.Add(this.chkHSongai);
            this.grpKizon4.Controls.Add(this.chkHJisyaKoza);
            this.grpKizon4.Controls.Add(this.chkHSetubi);
            this.grpKizon4.Controls.Add(this.chkHDataFmt);
            this.grpKizon4.Controls.Add(this.Label69);
            this.grpKizon4.Controls.Add(this.Label68);
            this.grpKizon4.Controls.Add(this.Label67);
            this.grpKizon4.Controls.Add(this.Label66);
            this.grpKizon4.Controls.Add(this.Label65);
            this.grpKizon4.Controls.Add(this.Label64);
            this.grpKizon4.Controls.Add(this.Label63);
            this.grpKizon4.Controls.Add(this.Label62);
            this.grpKizon4.Controls.Add(this.Label61);
            this.grpKizon4.Controls.Add(this.Label60);
            this.grpKizon4.Controls.Add(this.Label59);
            this.grpKizon4.Controls.Add(this.chkHNkinKomk);
            this.grpKizon4.Controls.Add(this.chkHTaiyo);
            this.grpKizon4.Controls.Add(this.chkHKozo);
            this.grpKizon4.Controls.Add(this.chkHKozaSyu);
            this.grpKizon4.Controls.Add(this.chkHEki);
            this.grpKizon4.Controls.Add(this.chkHEnsen);
            this.grpKizon4.Controls.Add(this.chkHKinyuSiten);
            this.grpKizon4.Controls.Add(this.chkHYouto);
            this.grpKizon4.Controls.Add(this.chkHBkBunrui);
            this.grpKizon4.Controls.Add(this.chkHKinyu);
            this.grpKizon4.Controls.Add(this.chkHKyBunrui);
            this.grpKizon4.Controls.Add(this.chkHKagi);
            this.grpKizon4.Controls.Add(this.chkHHouKenri);
            this.grpKizon4.Controls.Add(this.chkHNkinKbn);
            this.grpKizon4.Controls.Add(this.chkHHyBunrui);
            this.grpKizon4.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpKizon4.Location = new System.Drawing.Point(6, 7);
            this.grpKizon4.Name = "grpKizon4";
            this.grpKizon4.Size = new System.Drawing.Size(835, 265);
            this.grpKizon4.TabIndex = 1;
            this.grpKizon4.TabStop = false;
            this.grpKizon4.Text = " 紐付設定項目 ";
            // 
            // lblSelectPageCnt3
            // 
            this.lblSelectPageCnt3.AutoSize = true;
            this.lblSelectPageCnt3.Location = new System.Drawing.Point(794, 242);
            this.lblSelectPageCnt3.Name = "lblSelectPageCnt3";
            this.lblSelectPageCnt3.Size = new System.Drawing.Size(31, 18);
            this.lblSelectPageCnt3.TabIndex = 36;
            this.lblSelectPageCnt3.Text = "- / -";
            // 
            // Label26
            // 
            this.Label26.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.Label26.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label26.ForeColor = System.Drawing.Color.Red;
            this.Label26.Location = new System.Drawing.Point(82, 80);
            this.Label26.Name = "Label26";
            this.Label26.Size = new System.Drawing.Size(610, 159);
            this.Label26.TabIndex = 35;
            this.Label26.Text = "※没\r\nここで用意している紐付設定については、前画面の関連する親情報(物件情報や部屋情報など)にチェックがある場合に自動呼出ししています。\r\n(自動で呼び出すよう" +
    "にしている為、ここのチェックは現時点で機能していません)";
            // 
            // Label73
            // 
            this.Label73.BackColor = System.Drawing.SystemColors.Menu;
            this.Label73.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label73.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Label73.Location = new System.Drawing.Point(11, 20);
            this.Label73.Name = "Label73";
            this.Label73.Size = new System.Drawing.Size(795, 40);
            this.Label73.TabIndex = 0;
            this.Label73.Text = "紐付を行う必要がある項目に自動でチェックが入ります。\r\n(例えば、物件情報の項目にチェックを入れた場合、それに関連する項目(物件用途など)に自動でチェックが入りま" +
    "す)";
            this.Label73.UseCompatibleTextRendering = true;
            // 
            // Label70
            // 
            this.Label70.AutoSize = true;
            this.Label70.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label70.ForeColor = System.Drawing.Color.Black;
            this.Label70.Location = new System.Drawing.Point(347, 234);
            this.Label70.Name = "Label70";
            this.Label70.Size = new System.Drawing.Size(129, 17);
            this.Label70.TabIndex = 24;
            this.Label70.Text = "   (該当入金項目の選定)";
            // 
            // Label71
            // 
            this.Label71.AutoSize = true;
            this.Label71.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label71.ForeColor = System.Drawing.Color.Black;
            this.Label71.Location = new System.Drawing.Point(507, 188);
            this.Label71.Name = "Label71";
            this.Label71.Size = new System.Drawing.Size(140, 17);
            this.Label71.TabIndex = 30;
            this.Label71.Text = "   (全銀フォーマット選択)";
            // 
            // Label72
            // 
            this.Label72.AutoSize = true;
            this.Label72.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label72.ForeColor = System.Drawing.Color.Black;
            this.Label72.Location = new System.Drawing.Point(507, 142);
            this.Label72.Name = "Label72";
            this.Label72.Size = new System.Drawing.Size(107, 17);
            this.Label72.TabIndex = 28;
            this.Label72.Text = "   (自社口座の選択)";
            // 
            // Label74
            // 
            this.Label74.AutoSize = true;
            this.Label74.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label74.ForeColor = System.Drawing.Color.Black;
            this.Label74.Location = new System.Drawing.Point(182, 142);
            this.Label74.Name = "Label74";
            this.Label74.Size = new System.Drawing.Size(122, 17);
            this.Label74.TabIndex = 12;
            this.Label74.Text = "   (例:収納-物置-有り)";
            // 
            // chkHSongai
            // 
            this.chkHSongai.AutoSize = true;
            this.chkHSongai.ForeColor = System.Drawing.Color.Black;
            this.chkHSongai.Location = new System.Drawing.Point(350, 209);
            this.chkHSongai.Name = "chkHSongai";
            this.chkHSongai.Size = new System.Drawing.Size(87, 22);
            this.chkHSongai.TabIndex = 23;
            this.chkHSongai.Text = "損害保険料";
            this.chkHSongai.UseVisualStyleBackColor = true;
            // 
            // chkHJisyaKoza
            // 
            this.chkHJisyaKoza.AutoSize = true;
            this.chkHJisyaKoza.ForeColor = System.Drawing.Color.Black;
            this.chkHJisyaKoza.Location = new System.Drawing.Point(510, 117);
            this.chkHJisyaKoza.Name = "chkHJisyaKoza";
            this.chkHJisyaKoza.Size = new System.Drawing.Size(147, 22);
            this.chkHJisyaKoza.TabIndex = 27;
            this.chkHJisyaKoza.Text = "振込・振替・家賃口座";
            this.chkHJisyaKoza.UseVisualStyleBackColor = true;
            // 
            // chkHSetubi
            // 
            this.chkHSetubi.AutoSize = true;
            this.chkHSetubi.ForeColor = System.Drawing.Color.Black;
            this.chkHSetubi.Location = new System.Drawing.Point(185, 117);
            this.chkHSetubi.Name = "chkHSetubi";
            this.chkHSetubi.Size = new System.Drawing.Size(75, 22);
            this.chkHSetubi.TabIndex = 11;
            this.chkHSetubi.Text = "部屋設備";
            this.chkHSetubi.UseVisualStyleBackColor = true;
            // 
            // chkHDataFmt
            // 
            this.chkHDataFmt.AutoSize = true;
            this.chkHDataFmt.ForeColor = System.Drawing.Color.Black;
            this.chkHDataFmt.Location = new System.Drawing.Point(510, 163);
            this.chkHDataFmt.Name = "chkHDataFmt";
            this.chkHDataFmt.Size = new System.Drawing.Size(135, 22);
            this.chkHDataFmt.TabIndex = 29;
            this.chkHDataFmt.Text = "データフォーマット";
            this.chkHDataFmt.UseVisualStyleBackColor = true;
            // 
            // Label69
            // 
            this.Label69.AutoSize = true;
            this.Label69.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label69.ForeColor = System.Drawing.Color.Black;
            this.Label69.Location = new System.Drawing.Point(347, 188);
            this.Label69.Name = "Label69";
            this.Label69.Size = new System.Drawing.Size(111, 17);
            this.Label69.TabIndex = 22;
            this.Label69.Text = "   (例:賃料(毎月時))";
            // 
            // Label68
            // 
            this.Label68.AutoSize = true;
            this.Label68.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label68.ForeColor = System.Drawing.Color.Black;
            this.Label68.Location = new System.Drawing.Point(347, 142);
            this.Label68.Name = "Label68";
            this.Label68.Size = new System.Drawing.Size(68, 17);
            this.Label68.TabIndex = 20;
            this.Label68.Text = "   (例:振替)";
            // 
            // Label67
            // 
            this.Label67.AutoSize = true;
            this.Label67.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label67.ForeColor = System.Drawing.Color.Black;
            this.Label67.Location = new System.Drawing.Point(182, 234);
            this.Label67.Name = "Label67";
            this.Label67.Size = new System.Drawing.Size(68, 17);
            this.Label67.TabIndex = 16;
            this.Label67.Text = "   (例:仲介)";
            // 
            // Label66
            // 
            this.Label66.AutoSize = true;
            this.Label66.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label66.ForeColor = System.Drawing.Color.Black;
            this.Label66.Location = new System.Drawing.Point(347, 96);
            this.Label66.Name = "Label66";
            this.Label66.Size = new System.Drawing.Size(90, 17);
            this.Label66.TabIndex = 18;
            this.Label66.Text = "   (例:普通預金)";
            // 
            // Label65
            // 
            this.Label65.AutoSize = true;
            this.Label65.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label65.ForeColor = System.Drawing.Color.Black;
            this.Label65.Location = new System.Drawing.Point(182, 188);
            this.Label65.Name = "Label65";
            this.Label65.Size = new System.Drawing.Size(145, 17);
            this.Label65.TabIndex = 14;
            this.Label65.Text = "   (例:普通建物賃貸借契約)";
            // 
            // Label64
            // 
            this.Label64.AutoSize = true;
            this.Label64.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label64.ForeColor = System.Drawing.Color.Black;
            this.Label64.Location = new System.Drawing.Point(22, 142);
            this.Label64.Name = "Label64";
            this.Label64.Size = new System.Drawing.Size(145, 17);
            this.Label64.TabIndex = 4;
            this.Label64.Text = "   (例:鉄筋コンクリート造)";
            // 
            // Label63
            // 
            this.Label63.AutoSize = true;
            this.Label63.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label63.ForeColor = System.Drawing.Color.Black;
            this.Label63.Location = new System.Drawing.Point(507, 96);
            this.Label63.Name = "Label63";
            this.Label63.Size = new System.Drawing.Size(96, 17);
            this.Label63.TabIndex = 26;
            this.Label63.Text = "   (共用鍵の設定)";
            // 
            // Label62
            // 
            this.Label62.AutoSize = true;
            this.Label62.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label62.ForeColor = System.Drawing.Color.Black;
            this.Label62.Location = new System.Drawing.Point(182, 96);
            this.Label62.Name = "Label62";
            this.Label62.Size = new System.Drawing.Size(133, 17);
            this.Label62.TabIndex = 10;
            this.Label62.Text = "   (例:アパート(事業用))";
            // 
            // Label61
            // 
            this.Label61.AutoSize = true;
            this.Label61.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label61.ForeColor = System.Drawing.Color.Black;
            this.Label61.Location = new System.Drawing.Point(22, 234);
            this.Label61.Name = "Label61";
            this.Label61.Size = new System.Drawing.Size(90, 17);
            this.Label61.TabIndex = 8;
            this.Label61.Text = "   (例:アパート)";
            // 
            // Label60
            // 
            this.Label60.AutoSize = true;
            this.Label60.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label60.ForeColor = System.Drawing.Color.Black;
            this.Label60.Location = new System.Drawing.Point(22, 188);
            this.Label60.Name = "Label60";
            this.Label60.Size = new System.Drawing.Size(140, 17);
            this.Label60.TabIndex = 6;
            this.Label60.Text = "   (例：土砂災害警戒区域)";
            // 
            // Label59
            // 
            this.Label59.AutoSize = true;
            this.Label59.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label59.ForeColor = System.Drawing.Color.Black;
            this.Label59.Location = new System.Drawing.Point(22, 96);
            this.Label59.Name = "Label59";
            this.Label59.Size = new System.Drawing.Size(123, 17);
            this.Label59.TabIndex = 2;
            this.Label59.Text = "   (例:第一種住宅地域)";
            // 
            // chkHNkinKomk
            // 
            this.chkHNkinKomk.AutoSize = true;
            this.chkHNkinKomk.ForeColor = System.Drawing.Color.Black;
            this.chkHNkinKomk.Location = new System.Drawing.Point(350, 163);
            this.chkHNkinKomk.Name = "chkHNkinKomk";
            this.chkHNkinKomk.Size = new System.Drawing.Size(75, 22);
            this.chkHNkinKomk.TabIndex = 21;
            this.chkHNkinKomk.Text = "入金項目";
            this.chkHNkinKomk.UseVisualStyleBackColor = true;
            // 
            // chkHTaiyo
            // 
            this.chkHTaiyo.AutoSize = true;
            this.chkHTaiyo.ForeColor = System.Drawing.Color.Black;
            this.chkHTaiyo.Location = new System.Drawing.Point(185, 209);
            this.chkHTaiyo.Name = "chkHTaiyo";
            this.chkHTaiyo.Size = new System.Drawing.Size(75, 22);
            this.chkHTaiyo.TabIndex = 15;
            this.chkHTaiyo.Text = "取引態様";
            this.chkHTaiyo.UseVisualStyleBackColor = true;
            // 
            // chkHKozo
            // 
            this.chkHKozo.AutoSize = true;
            this.chkHKozo.ForeColor = System.Drawing.Color.Black;
            this.chkHKozo.Location = new System.Drawing.Point(25, 117);
            this.chkHKozo.Name = "chkHKozo";
            this.chkHKozo.Size = new System.Drawing.Size(75, 22);
            this.chkHKozo.TabIndex = 3;
            this.chkHKozo.Text = "物件構造";
            this.chkHKozo.UseVisualStyleBackColor = true;
            // 
            // chkHKozaSyu
            // 
            this.chkHKozaSyu.AutoSize = true;
            this.chkHKozaSyu.ForeColor = System.Drawing.Color.Black;
            this.chkHKozaSyu.Location = new System.Drawing.Point(350, 70);
            this.chkHKozaSyu.Name = "chkHKozaSyu";
            this.chkHKozaSyu.Size = new System.Drawing.Size(75, 22);
            this.chkHKozaSyu.TabIndex = 17;
            this.chkHKozaSyu.Text = "口座種別";
            this.chkHKozaSyu.UseVisualStyleBackColor = true;
            // 
            // chkHEki
            // 
            this.chkHEki.AutoSize = true;
            this.chkHEki.ForeColor = System.Drawing.Color.Black;
            this.chkHEki.Location = new System.Drawing.Point(670, 209);
            this.chkHEki.Name = "chkHEki";
            this.chkHEki.Size = new System.Drawing.Size(39, 22);
            this.chkHEki.TabIndex = 34;
            this.chkHEki.Text = "駅";
            this.chkHEki.UseVisualStyleBackColor = true;
            // 
            // chkHEnsen
            // 
            this.chkHEnsen.AutoSize = true;
            this.chkHEnsen.ForeColor = System.Drawing.Color.Black;
            this.chkHEnsen.Location = new System.Drawing.Point(670, 163);
            this.chkHEnsen.Name = "chkHEnsen";
            this.chkHEnsen.Size = new System.Drawing.Size(51, 22);
            this.chkHEnsen.TabIndex = 33;
            this.chkHEnsen.Text = "沿線";
            this.chkHEnsen.UseVisualStyleBackColor = true;
            // 
            // chkHKinyuSiten
            // 
            this.chkHKinyuSiten.AutoSize = true;
            this.chkHKinyuSiten.ForeColor = System.Drawing.Color.Black;
            this.chkHKinyuSiten.Location = new System.Drawing.Point(670, 117);
            this.chkHKinyuSiten.Name = "chkHKinyuSiten";
            this.chkHKinyuSiten.Size = new System.Drawing.Size(111, 22);
            this.chkHKinyuSiten.TabIndex = 32;
            this.chkHKinyuSiten.Text = "金融機関本支店";
            this.chkHKinyuSiten.UseVisualStyleBackColor = true;
            // 
            // chkHYouto
            // 
            this.chkHYouto.AutoSize = true;
            this.chkHYouto.ForeColor = System.Drawing.Color.Black;
            this.chkHYouto.Location = new System.Drawing.Point(25, 70);
            this.chkHYouto.Name = "chkHYouto";
            this.chkHYouto.Size = new System.Drawing.Size(75, 22);
            this.chkHYouto.TabIndex = 1;
            this.chkHYouto.Text = "物件用途";
            this.chkHYouto.UseVisualStyleBackColor = true;
            // 
            // chkHBkBunrui
            // 
            this.chkHBkBunrui.AutoSize = true;
            this.chkHBkBunrui.ForeColor = System.Drawing.Color.Black;
            this.chkHBkBunrui.Location = new System.Drawing.Point(25, 209);
            this.chkHBkBunrui.Name = "chkHBkBunrui";
            this.chkHBkBunrui.Size = new System.Drawing.Size(75, 22);
            this.chkHBkBunrui.TabIndex = 7;
            this.chkHBkBunrui.Text = "物件分類";
            this.chkHBkBunrui.UseVisualStyleBackColor = true;
            // 
            // chkHKinyu
            // 
            this.chkHKinyu.AutoSize = true;
            this.chkHKinyu.ForeColor = System.Drawing.Color.Black;
            this.chkHKinyu.Location = new System.Drawing.Point(665, 71);
            this.chkHKinyu.Name = "chkHKinyu";
            this.chkHKinyu.Size = new System.Drawing.Size(75, 22);
            this.chkHKinyu.TabIndex = 31;
            this.chkHKinyu.Text = "金融機関";
            this.chkHKinyu.UseVisualStyleBackColor = true;
            // 
            // chkHKyBunrui
            // 
            this.chkHKyBunrui.AutoSize = true;
            this.chkHKyBunrui.ForeColor = System.Drawing.Color.Black;
            this.chkHKyBunrui.Location = new System.Drawing.Point(185, 163);
            this.chkHKyBunrui.Name = "chkHKyBunrui";
            this.chkHKyBunrui.Size = new System.Drawing.Size(75, 22);
            this.chkHKyBunrui.TabIndex = 13;
            this.chkHKyBunrui.Text = "契約分類";
            this.chkHKyBunrui.UseVisualStyleBackColor = true;
            // 
            // chkHKagi
            // 
            this.chkHKagi.AutoSize = true;
            this.chkHKagi.ForeColor = System.Drawing.Color.Black;
            this.chkHKagi.Location = new System.Drawing.Point(510, 70);
            this.chkHKagi.Name = "chkHKagi";
            this.chkHKagi.Size = new System.Drawing.Size(63, 22);
            this.chkHKagi.TabIndex = 25;
            this.chkHKagi.Text = "鍵情報";
            this.chkHKagi.UseVisualStyleBackColor = true;
            // 
            // chkHHouKenri
            // 
            this.chkHHouKenri.AutoSize = true;
            this.chkHHouKenri.ForeColor = System.Drawing.Color.Black;
            this.chkHHouKenri.Location = new System.Drawing.Point(25, 163);
            this.chkHHouKenri.Name = "chkHHouKenri";
            this.chkHHouKenri.Size = new System.Drawing.Size(111, 22);
            this.chkHHouKenri.TabIndex = 5;
            this.chkHHouKenri.Text = "物件法令・権利";
            this.chkHHouKenri.UseVisualStyleBackColor = true;
            // 
            // chkHNkinKbn
            // 
            this.chkHNkinKbn.AutoSize = true;
            this.chkHNkinKbn.ForeColor = System.Drawing.Color.Black;
            this.chkHNkinKbn.Location = new System.Drawing.Point(350, 117);
            this.chkHNkinKbn.Name = "chkHNkinKbn";
            this.chkHNkinKbn.Size = new System.Drawing.Size(75, 22);
            this.chkHNkinKbn.TabIndex = 19;
            this.chkHNkinKbn.Text = "入金区分";
            this.chkHNkinKbn.UseVisualStyleBackColor = true;
            // 
            // chkHHyBunrui
            // 
            this.chkHHyBunrui.AutoSize = true;
            this.chkHHyBunrui.ForeColor = System.Drawing.Color.Black;
            this.chkHHyBunrui.Location = new System.Drawing.Point(185, 71);
            this.chkHHyBunrui.Name = "chkHHyBunrui";
            this.chkHHyBunrui.Size = new System.Drawing.Size(75, 22);
            this.chkHHyBunrui.TabIndex = 9;
            this.chkHHyBunrui.Text = "部屋分類";
            this.chkHHyBunrui.UseVisualStyleBackColor = true;
            // 
            // pnlRekiClear
            // 
            this.pnlRekiClear.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnlRekiClear.Controls.Add(this.btnRekiClear);
            this.pnlRekiClear.Controls.Add(this.Label79);
            this.pnlRekiClear.Controls.Add(this.Label298);
            this.pnlRekiClear.Location = new System.Drawing.Point(563, 36);
            this.pnlRekiClear.Name = "pnlRekiClear";
            this.pnlRekiClear.Size = new System.Drawing.Size(326, 45);
            this.pnlRekiClear.TabIndex = 45;
            // 
            // btnRekiClear
            // 
            this.btnRekiClear.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRekiClear.Image = ((System.Drawing.Image)(resources.GetObject("btnRekiClear.Image")));
            this.btnRekiClear.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRekiClear.Location = new System.Drawing.Point(229, 5);
            this.btnRekiClear.Name = "btnRekiClear";
            this.btnRekiClear.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnRekiClear.Size = new System.Drawing.Size(88, 30);
            this.btnRekiClear.TabIndex = 2;
            this.btnRekiClear.Text = " クリア";
            this.btnRekiClear.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnRekiClear.UseVisualStyleBackColor = true;
            this.btnRekiClear.Click += new System.EventHandler(this.btnRekiClear_Click);
            // 
            // Label79
            // 
            this.Label79.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label79.ForeColor = System.Drawing.Color.ForestGreen;
            this.Label79.Location = new System.Drawing.Point(18, 2);
            this.Label79.Name = "Label79";
            this.Label79.Size = new System.Drawing.Size(32, 18);
            this.Label79.TabIndex = 1;
            this.Label79.Text = "緑色";
            // 
            // Label298
            // 
            this.Label298.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label298.ForeColor = System.Drawing.Color.Black;
            this.Label298.Location = new System.Drawing.Point(3, 3);
            this.Label298.Name = "Label298";
            this.Label298.Size = new System.Drawing.Size(228, 38);
            this.Label298.TabIndex = 0;
            this.Label298.Text = "※　　　に着色されているチェック\r\n　ボックスはコンバート実行済み項目です。";
            // 
            // lblDatacvSelectDescription1
            // 
            this.lblDatacvSelectDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblDatacvSelectDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDatacvSelectDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblDatacvSelectDescription1.Location = new System.Drawing.Point(30, 20);
            this.lblDatacvSelectDescription1.Name = "lblDatacvSelectDescription1";
            this.lblDatacvSelectDescription1.Size = new System.Drawing.Size(876, 80);
            this.lblDatacvSelectDescription1.TabIndex = 0;
            this.lblDatacvSelectDescription1.Text = "コンバート対象項目の選択を行います。\r\nコンバートを行う項目にチェックを入れて、[次へ] ボタンを押して下さい。";
            this.lblDatacvSelectDescription1.UseCompatibleTextRendering = true;
            // 
            // tabPageJikko
            // 
            this.tabPageJikko.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageJikko.Controls.Add(this.lblDatacvJikkoCaution);
            this.tabPageJikko.Controls.Add(this.lblDatacvJikkoDescription1);
            this.tabPageJikko.Controls.Add(this.grpTotalProcess);
            this.tabPageJikko.Location = new System.Drawing.Point(4, 27);
            this.tabPageJikko.Name = "tabPageJikko";
            this.tabPageJikko.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageJikko.Size = new System.Drawing.Size(912, 549);
            this.tabPageJikko.TabIndex = 3;
            this.tabPageJikko.Text = "移行処理";
            // 
            // lblDatacvJikkoCaution
            // 
            this.lblDatacvJikkoCaution.BackColor = System.Drawing.SystemColors.Menu;
            this.lblDatacvJikkoCaution.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDatacvJikkoCaution.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.lblDatacvJikkoCaution.Location = new System.Drawing.Point(30, 45);
            this.lblDatacvJikkoCaution.Name = "lblDatacvJikkoCaution";
            this.lblDatacvJikkoCaution.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblDatacvJikkoCaution.Size = new System.Drawing.Size(851, 60);
            this.lblDatacvJikkoCaution.TabIndex = 119;
            this.lblDatacvJikkoCaution.Text = "※コンバート処理中は、賃貸革命を使用しないで下さい。\r\n※「キャンセル」ボタンで処理を中止します (途中再開(リジューム)はできません)。\r\n※中間ファイル書込処" +
    "理完了後に紐付設定画面が表示されます。項目の紐付設定を行って下さい。";
            this.lblDatacvJikkoCaution.UseCompatibleTextRendering = true;
            // 
            // lblDatacvJikkoDescription1
            // 
            this.lblDatacvJikkoDescription1.BackColor = System.Drawing.SystemColors.Menu;
            this.lblDatacvJikkoDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDatacvJikkoDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblDatacvJikkoDescription1.Location = new System.Drawing.Point(30, 20);
            this.lblDatacvJikkoDescription1.Name = "lblDatacvJikkoDescription1";
            this.lblDatacvJikkoDescription1.Size = new System.Drawing.Size(851, 80);
            this.lblDatacvJikkoDescription1.TabIndex = 115;
            this.lblDatacvJikkoDescription1.Text = "データコンバート処理を行います。処理が完了したら、「次へ」ボタンが表示されます。表示後、「次へ」ボタンを押して下さい。\r\n";
            this.lblDatacvJikkoDescription1.UseCompatibleTextRendering = true;
            // 
            // grpTotalProcess
            // 
            this.grpTotalProcess.Controls.Add(this.lblCVItem);
            this.grpTotalProcess.Controls.Add(this.txtPartialSituation);
            this.grpTotalProcess.Controls.Add(this.Label31);
            this.grpTotalProcess.Controls.Add(this.lblPgbPartial);
            this.grpTotalProcess.Controls.Add(this.lblPgbTotal);
            this.grpTotalProcess.Controls.Add(this.txtTotalSituation);
            this.grpTotalProcess.Controls.Add(this.pgbpartial);
            this.grpTotalProcess.Controls.Add(this.pgbTotal);
            this.grpTotalProcess.Controls.Add(this.Label32);
            this.grpTotalProcess.Controls.Add(this.lblTotalSituation);
            this.grpTotalProcess.Location = new System.Drawing.Point(22, 121);
            this.grpTotalProcess.Name = "grpTotalProcess";
            this.grpTotalProcess.Size = new System.Drawing.Size(859, 408);
            this.grpTotalProcess.TabIndex = 113;
            this.grpTotalProcess.TabStop = false;
            this.grpTotalProcess.Text = "【進捗状況】";
            // 
            // lblCVItem
            // 
            this.lblCVItem.Location = new System.Drawing.Point(16, 87);
            this.lblCVItem.Name = "lblCVItem";
            this.lblCVItem.Size = new System.Drawing.Size(336, 16);
            this.lblCVItem.TabIndex = 114;
            this.lblCVItem.Text = "処理項目";
            this.lblCVItem.UseCompatibleTextRendering = true;
            // 
            // txtPartialSituation
            // 
            this.txtPartialSituation.Location = new System.Drawing.Point(452, 202);
            this.txtPartialSituation.Multiline = true;
            this.txtPartialSituation.Name = "txtPartialSituation";
            this.txtPartialSituation.Size = new System.Drawing.Size(365, 190);
            this.txtPartialSituation.TabIndex = 116;
            // 
            // Label31
            // 
            this.Label31.Location = new System.Drawing.Point(452, 169);
            this.Label31.Name = "Label31";
            this.Label31.Size = new System.Drawing.Size(80, 16);
            this.Label31.TabIndex = 113;
            this.Label31.Text = "処理完了項目";
            this.Label31.UseCompatibleTextRendering = true;
            // 
            // lblPgbPartial
            // 
            this.lblPgbPartial.Location = new System.Drawing.Point(796, 107);
            this.lblPgbPartial.Name = "lblPgbPartial";
            this.lblPgbPartial.Size = new System.Drawing.Size(52, 20);
            this.lblPgbPartial.TabIndex = 117;
            this.lblPgbPartial.Text = "0 %";
            this.lblPgbPartial.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblPgbPartial.UseCompatibleTextRendering = true;
            // 
            // lblPgbTotal
            // 
            this.lblPgbTotal.Location = new System.Drawing.Point(797, 50);
            this.lblPgbTotal.Name = "lblPgbTotal";
            this.lblPgbTotal.Size = new System.Drawing.Size(51, 20);
            this.lblPgbTotal.TabIndex = 115;
            this.lblPgbTotal.Text = "0 %";
            this.lblPgbTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblPgbTotal.UseCompatibleTextRendering = true;
            // 
            // txtTotalSituation
            // 
            this.txtTotalSituation.Location = new System.Drawing.Point(40, 202);
            this.txtTotalSituation.Multiline = true;
            this.txtTotalSituation.Name = "txtTotalSituation";
            this.txtTotalSituation.Size = new System.Drawing.Size(365, 190);
            this.txtTotalSituation.TabIndex = 114;
            // 
            // pgbpartial
            // 
            this.pgbpartial.Location = new System.Drawing.Point(40, 107);
            this.pgbpartial.Name = "pgbpartial";
            this.pgbpartial.Size = new System.Drawing.Size(750, 20);
            this.pgbpartial.TabIndex = 115;
            // 
            // pgbTotal
            // 
            this.pgbTotal.Location = new System.Drawing.Point(40, 50);
            this.pgbTotal.Name = "pgbTotal";
            this.pgbTotal.Size = new System.Drawing.Size(751, 20);
            this.pgbTotal.TabIndex = 113;
            // 
            // Label32
            // 
            this.Label32.Location = new System.Drawing.Point(40, 169);
            this.Label32.Name = "Label32";
            this.Label32.Size = new System.Drawing.Size(80, 16);
            this.Label32.TabIndex = 112;
            this.Label32.Text = "処理経過状況";
            this.Label32.UseCompatibleTextRendering = true;
            // 
            // lblTotalSituation
            // 
            this.lblTotalSituation.Location = new System.Drawing.Point(16, 29);
            this.lblTotalSituation.Name = "lblTotalSituation";
            this.lblTotalSituation.Size = new System.Drawing.Size(336, 16);
            this.lblTotalSituation.TabIndex = 104;
            this.lblTotalSituation.Text = "コンバート処理中…";
            this.lblTotalSituation.UseCompatibleTextRendering = true;
            // 
            // tabPageEndOK
            // 
            this.tabPageEndOK.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageEndOK.Controls.Add(this.lblDatacvEndDescription1);
            this.tabPageEndOK.Controls.Add(this.lblHDatacvEndLabel);
            this.tabPageEndOK.Controls.Add(this.PictureBox15);
            this.tabPageEndOK.Location = new System.Drawing.Point(4, 27);
            this.tabPageEndOK.Name = "tabPageEndOK";
            this.tabPageEndOK.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageEndOK.Size = new System.Drawing.Size(912, 549);
            this.tabPageEndOK.TabIndex = 4;
            this.tabPageEndOK.Text = " 終了";
            // 
            // lblDatacvEndDescription1
            // 
            this.lblDatacvEndDescription1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDatacvEndDescription1.ForeColor = System.Drawing.Color.Navy;
            this.lblDatacvEndDescription1.Location = new System.Drawing.Point(34, 49);
            this.lblDatacvEndDescription1.Name = "lblDatacvEndDescription1";
            this.lblDatacvEndDescription1.Size = new System.Drawing.Size(820, 51);
            this.lblDatacvEndDescription1.TabIndex = 129;
            this.lblDatacvEndDescription1.Text = "[ログ確認]ボタンよりコンバート結果を確認して下さい。";
            this.lblDatacvEndDescription1.UseCompatibleTextRendering = true;
            // 
            // lblHDatacvEndLabel
            // 
            this.lblHDatacvEndLabel.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblHDatacvEndLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lblHDatacvEndLabel.Location = new System.Drawing.Point(30, 20);
            this.lblHDatacvEndLabel.Name = "lblHDatacvEndLabel";
            this.lblHDatacvEndLabel.Size = new System.Drawing.Size(820, 80);
            this.lblHDatacvEndLabel.TabIndex = 131;
            this.lblHDatacvEndLabel.Text = "「賃貸革命10」のコンバートが完了しました。";
            this.lblHDatacvEndLabel.UseCompatibleTextRendering = true;
            // 
            // PictureBox15
            // 
            this.PictureBox15.Image = ((System.Drawing.Image)(resources.GetObject("PictureBox15.Image")));
            this.PictureBox15.Location = new System.Drawing.Point(466, 360);
            this.PictureBox15.Name = "PictureBox15";
            this.PictureBox15.Size = new System.Drawing.Size(420, 170);
            this.PictureBox15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PictureBox15.TabIndex = 130;
            this.PictureBox15.TabStop = false;
            // 
            // tabPageEndError
            // 
            this.tabPageEndError.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageEndError.Controls.Add(this.GroupBox23);
            this.tabPageEndError.Controls.Add(this.lblDatacvErrorDescription1);
            this.tabPageEndError.Controls.Add(this.GroupBox19);
            this.tabPageEndError.Controls.Add(this.GroupBox20);
            this.tabPageEndError.Controls.Add(this.lblDatacvErrorLabel);
            this.tabPageEndError.Controls.Add(this.PictureBox23);
            this.tabPageEndError.Location = new System.Drawing.Point(4, 27);
            this.tabPageEndError.Name = "tabPageEndError";
            this.tabPageEndError.Size = new System.Drawing.Size(912, 549);
            this.tabPageEndError.TabIndex = 10;
            this.tabPageEndError.Text = "異常終了";
            // 
            // GroupBox23
            // 
            this.GroupBox23.BackColor = System.Drawing.SystemColors.Menu;
            this.GroupBox23.Controls.Add(this.Label163);
            this.GroupBox23.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox23.Location = new System.Drawing.Point(31, 329);
            this.GroupBox23.Name = "GroupBox23";
            this.GroupBox23.Size = new System.Drawing.Size(420, 200);
            this.GroupBox23.TabIndex = 134;
            this.GroupBox23.TabStop = false;
            this.GroupBox23.Text = "【対処方法例】";
            // 
            // Label163
            // 
            this.Label163.Location = new System.Drawing.Point(10, 25);
            this.Label163.Name = "Label163";
            this.Label163.Size = new System.Drawing.Size(400, 150);
            this.Label163.TabIndex = 117;
            this.Label163.Text = "(例) 接続設定情報画面の賃貸革命V7と10の接続情報が正しく登録されているか確認して下さい。";
            this.Label163.UseCompatibleTextRendering = true;
            // 
            // lblDatacvErrorDescription1
            // 
            this.lblDatacvErrorDescription1.ForeColor = System.Drawing.Color.Red;
            this.lblDatacvErrorDescription1.Location = new System.Drawing.Point(31, 47);
            this.lblDatacvErrorDescription1.Name = "lblDatacvErrorDescription1";
            this.lblDatacvErrorDescription1.Size = new System.Drawing.Size(820, 51);
            this.lblDatacvErrorDescription1.TabIndex = 134;
            this.lblDatacvErrorDescription1.Text = "コンバート処理を正常に終了することができませんでした。\r\n[ログ確認]ボタンよりコンバート結果を確認して下さい。";
            this.lblDatacvErrorDescription1.UseCompatibleTextRendering = true;
            // 
            // GroupBox19
            // 
            this.GroupBox19.BackColor = System.Drawing.SystemColors.Menu;
            this.GroupBox19.Controls.Add(this.Label170);
            this.GroupBox19.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox19.Location = new System.Drawing.Point(466, 123);
            this.GroupBox19.Name = "GroupBox19";
            this.GroupBox19.Size = new System.Drawing.Size(420, 200);
            this.GroupBox19.TabIndex = 133;
            this.GroupBox19.TabStop = false;
            this.GroupBox19.Text = "【エラー内容】";
            // 
            // Label170
            // 
            this.Label170.Location = new System.Drawing.Point(10, 25);
            this.Label170.Name = "Label170";
            this.Label170.Size = new System.Drawing.Size(400, 150);
            this.Label170.TabIndex = 117;
            this.Label170.Text = "(例) 賃貸革命V7と10の設定値が逆になっている可能性があります。";
            this.Label170.UseCompatibleTextRendering = true;
            // 
            // GroupBox20
            // 
            this.GroupBox20.BackColor = System.Drawing.SystemColors.Menu;
            this.GroupBox20.Controls.Add(this.Label176);
            this.GroupBox20.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox20.Location = new System.Drawing.Point(31, 123);
            this.GroupBox20.Name = "GroupBox20";
            this.GroupBox20.Size = new System.Drawing.Size(420, 200);
            this.GroupBox20.TabIndex = 132;
            this.GroupBox20.TabStop = false;
            this.GroupBox20.Text = "【エラー個所】";
            // 
            // Label176
            // 
            this.Label176.Location = new System.Drawing.Point(10, 25);
            this.Label176.Name = "Label176";
            this.Label176.Size = new System.Drawing.Size(400, 150);
            this.Label176.TabIndex = 117;
            this.Label176.Text = "(例) 接続設定画面";
            this.Label176.UseCompatibleTextRendering = true;
            // 
            // lblDatacvErrorLabel
            // 
            this.lblDatacvErrorLabel.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDatacvErrorLabel.ForeColor = System.Drawing.Color.Red;
            this.lblDatacvErrorLabel.Location = new System.Drawing.Point(27, 18);
            this.lblDatacvErrorLabel.Name = "lblDatacvErrorLabel";
            this.lblDatacvErrorLabel.Size = new System.Drawing.Size(820, 80);
            this.lblDatacvErrorLabel.TabIndex = 131;
            this.lblDatacvErrorLabel.Text = "コンバートを異常終了しました。";
            this.lblDatacvErrorLabel.UseCompatibleTextRendering = true;
            // 
            // PictureBox23
            // 
            this.PictureBox23.Image = ((System.Drawing.Image)(resources.GetObject("PictureBox23.Image")));
            this.PictureBox23.Location = new System.Drawing.Point(466, 360);
            this.PictureBox23.Name = "PictureBox23";
            this.PictureBox23.Size = new System.Drawing.Size(420, 170);
            this.PictureBox23.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PictureBox23.TabIndex = 135;
            this.PictureBox23.TabStop = false;
            // 
            // tabPageEndCancel
            // 
            this.tabPageEndCancel.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageEndCancel.Controls.Add(this.lblDatacvCancelDescription1);
            this.tabPageEndCancel.Controls.Add(this.GroupBox21);
            this.tabPageEndCancel.Controls.Add(this.GroupBox22);
            this.tabPageEndCancel.Controls.Add(this.lblDatacvCancelLabel);
            this.tabPageEndCancel.Controls.Add(this.PictureBox25);
            this.tabPageEndCancel.Location = new System.Drawing.Point(4, 27);
            this.tabPageEndCancel.Name = "tabPageEndCancel";
            this.tabPageEndCancel.Size = new System.Drawing.Size(912, 549);
            this.tabPageEndCancel.TabIndex = 11;
            this.tabPageEndCancel.Text = "キャンセル終了";
            // 
            // lblDatacvCancelDescription1
            // 
            this.lblDatacvCancelDescription1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.lblDatacvCancelDescription1.Location = new System.Drawing.Point(31, 47);
            this.lblDatacvCancelDescription1.Name = "lblDatacvCancelDescription1";
            this.lblDatacvCancelDescription1.Size = new System.Drawing.Size(820, 51);
            this.lblDatacvCancelDescription1.TabIndex = 134;
            this.lblDatacvCancelDescription1.Text = "途中キャンセルした場合は不完全なデータの移行となる為、データの整合性が取れない場合があります。\r\n完全なデータの移行を行う為、本プログラムを再起動し、コンバート作" +
    "業を再度行って下さい。";
            this.lblDatacvCancelDescription1.UseCompatibleTextRendering = true;
            // 
            // GroupBox21
            // 
            this.GroupBox21.BackColor = System.Drawing.SystemColors.Menu;
            this.GroupBox21.Controls.Add(this.Label186);
            this.GroupBox21.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox21.Location = new System.Drawing.Point(466, 123);
            this.GroupBox21.Name = "GroupBox21";
            this.GroupBox21.Size = new System.Drawing.Size(420, 200);
            this.GroupBox21.TabIndex = 133;
            this.GroupBox21.TabStop = false;
            this.GroupBox21.Text = "【キャンセル内容】";
            // 
            // Label186
            // 
            this.Label186.Location = new System.Drawing.Point(10, 25);
            this.Label186.Name = "Label186";
            this.Label186.Size = new System.Drawing.Size(400, 150);
            this.Label186.TabIndex = 117;
            this.Label186.Text = "(例) 紐付設定処理中";
            this.Label186.UseCompatibleTextRendering = true;
            // 
            // GroupBox22
            // 
            this.GroupBox22.BackColor = System.Drawing.SystemColors.Menu;
            this.GroupBox22.Controls.Add(this.Label192);
            this.GroupBox22.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox22.Location = new System.Drawing.Point(31, 123);
            this.GroupBox22.Name = "GroupBox22";
            this.GroupBox22.Size = new System.Drawing.Size(400, 200);
            this.GroupBox22.TabIndex = 132;
            this.GroupBox22.TabStop = false;
            this.GroupBox22.Text = "【キャンセル箇所】";
            // 
            // Label192
            // 
            this.Label192.Location = new System.Drawing.Point(10, 25);
            this.Label192.Name = "Label192";
            this.Label192.Size = new System.Drawing.Size(384, 150);
            this.Label192.TabIndex = 117;
            this.Label192.Text = "(例) 紐付設定画面";
            this.Label192.UseCompatibleTextRendering = true;
            // 
            // lblDatacvCancelLabel
            // 
            this.lblDatacvCancelLabel.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDatacvCancelLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.lblDatacvCancelLabel.Location = new System.Drawing.Point(27, 18);
            this.lblDatacvCancelLabel.Name = "lblDatacvCancelLabel";
            this.lblDatacvCancelLabel.Size = new System.Drawing.Size(820, 80);
            this.lblDatacvCancelLabel.TabIndex = 131;
            this.lblDatacvCancelLabel.Text = "コンバート処理がキャンセルされました。";
            this.lblDatacvCancelLabel.UseCompatibleTextRendering = true;
            // 
            // PictureBox25
            // 
            this.PictureBox25.Image = ((System.Drawing.Image)(resources.GetObject("PictureBox25.Image")));
            this.PictureBox25.Location = new System.Drawing.Point(466, 360);
            this.PictureBox25.Name = "PictureBox25";
            this.PictureBox25.Size = new System.Drawing.Size(420, 170);
            this.PictureBox25.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PictureBox25.TabIndex = 135;
            this.PictureBox25.TabStop = false;
            // 
            // tabPageIkkatu
            // 
            this.tabPageIkkatu.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageIkkatu.Location = new System.Drawing.Point(4, 27);
            this.tabPageIkkatu.Name = "tabPageIkkatu";
            this.tabPageIkkatu.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageIkkatu.Size = new System.Drawing.Size(912, 549);
            this.tabPageIkkatu.TabIndex = 9;
            this.tabPageIkkatu.Text = " 必須項目一括設定(未定)";
            // 
            // tabPageHanyoJizen
            // 
            this.tabPageHanyoJizen.BackColor = System.Drawing.SystemColors.Menu;
            this.tabPageHanyoJizen.Location = new System.Drawing.Point(4, 27);
            this.tabPageHanyoJizen.Name = "tabPageHanyoJizen";
            this.tabPageHanyoJizen.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageHanyoJizen.Size = new System.Drawing.Size(912, 549);
            this.tabPageHanyoJizen.TabIndex = 8;
            this.tabPageHanyoJizen.Text = " 事前調整作業(汎用)";
            // 
            // btnDevTabChange
            // 
            this.btnDevTabChange.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDevTabChange.Location = new System.Drawing.Point(66, 638);
            this.btnDevTabChange.Name = "btnDevTabChange";
            this.btnDevTabChange.Size = new System.Drawing.Size(180, 30);
            this.btnDevTabChange.TabIndex = 4;
            this.btnDevTabChange.Text = "開発用画面表示ON/OFF";
            this.btnDevTabChange.UseVisualStyleBackColor = true;
            this.btnDevTabChange.Visible = false;
            this.btnDevTabChange.Click += new System.EventHandler(this.btnDevTabChange_Click);
            // 
            // Label393
            // 
            this.Label393.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label393.Location = new System.Drawing.Point(40, 186);
            this.Label393.Name = "Label393";
            this.Label393.Size = new System.Drawing.Size(340, 65);
            this.Label393.TabIndex = 5;
            this.Label393.UseCompatibleTextRendering = true;
            // 
            // Label392
            // 
            this.Label392.AutoSize = true;
            this.Label392.Location = new System.Drawing.Point(794, 310);
            this.Label392.Name = "Label392";
            this.Label392.Size = new System.Drawing.Size(35, 18);
            this.Label392.TabIndex = 0;
            // 
            // CheckBox51
            // 
            this.CheckBox51.AutoSize = true;
            this.CheckBox51.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CheckBox51.Location = new System.Drawing.Point(20, 161);
            this.CheckBox51.Name = "CheckBox51";
            this.CheckBox51.Size = new System.Drawing.Size(147, 22);
            this.CheckBox51.TabIndex = 4;
            this.CheckBox51.Text = "控除支払データの確認";
            this.CheckBox51.UseVisualStyleBackColor = true;
            // 
            // Label391
            // 
            this.Label391.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label391.Location = new System.Drawing.Point(40, 50);
            this.Label391.Name = "Label391";
            this.Label391.Size = new System.Drawing.Size(340, 33);
            this.Label391.TabIndex = 1;
            this.Label391.UseCompatibleTextRendering = true;
            // 
            // CheckBox50
            // 
            this.CheckBox50.AutoSize = true;
            this.CheckBox50.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CheckBox50.Location = new System.Drawing.Point(20, 25);
            this.CheckBox50.Name = "CheckBox50";
            this.CheckBox50.Size = new System.Drawing.Size(195, 22);
            this.CheckBox50.TabIndex = 0;
            this.CheckBox50.Text = "分割入金の未収分データの確認";
            this.CheckBox50.UseVisualStyleBackColor = true;
            // 
            // Panel24
            // 
            this.Panel24.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.Panel24.Location = new System.Drawing.Point(40, 254);
            this.Panel24.Name = "Panel24";
            this.Panel24.Size = new System.Drawing.Size(340, 60);
            this.Panel24.TabIndex = 6;
            // 
            // Label390
            // 
            this.Label390.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label390.Location = new System.Drawing.Point(218, 31);
            this.Label390.Name = "Label390";
            this.Label390.Size = new System.Drawing.Size(83, 16);
            this.Label390.TabIndex = 3;
            this.Label390.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label389
            // 
            this.Label389.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label389.Location = new System.Drawing.Point(131, 31);
            this.Label389.Name = "Label389";
            this.Label389.Size = new System.Drawing.Size(92, 16);
            this.Label389.TabIndex = 2;
            this.Label389.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label388
            // 
            this.Label388.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label388.Location = new System.Drawing.Point(307, 31);
            this.Label388.Name = "Label388";
            this.Label388.Size = new System.Drawing.Size(22, 16);
            this.Label388.TabIndex = 4;
            this.Label388.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label387
            // 
            this.Label387.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label387.Location = new System.Drawing.Point(8, 8);
            this.Label387.Name = "Label387";
            this.Label387.Size = new System.Drawing.Size(321, 16);
            this.Label387.TabIndex = 0;
            this.Label387.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label387.UseCompatibleTextRendering = true;
            // 
            // Panel23
            // 
            this.Panel23.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.Panel23.Location = new System.Drawing.Point(40, 86);
            this.Panel23.Name = "Panel23";
            this.Panel23.Size = new System.Drawing.Size(340, 60);
            this.Panel23.TabIndex = 3;
            // 
            // Label386
            // 
            this.Label386.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label386.Location = new System.Drawing.Point(218, 31);
            this.Label386.Name = "Label386";
            this.Label386.Size = new System.Drawing.Size(83, 16);
            this.Label386.TabIndex = 3;
            this.Label386.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label385
            // 
            this.Label385.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label385.Location = new System.Drawing.Point(131, 31);
            this.Label385.Name = "Label385";
            this.Label385.Size = new System.Drawing.Size(92, 16);
            this.Label385.TabIndex = 2;
            this.Label385.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label384
            // 
            this.Label384.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label384.Location = new System.Drawing.Point(307, 31);
            this.Label384.Name = "Label384";
            this.Label384.Size = new System.Drawing.Size(22, 16);
            this.Label384.TabIndex = 4;
            this.Label384.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label383
            // 
            this.Label383.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label383.Location = new System.Drawing.Point(8, 8);
            this.Label383.Name = "Label383";
            this.Label383.Size = new System.Drawing.Size(321, 16);
            this.Label383.TabIndex = 0;
            this.Label383.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label383.UseCompatibleTextRendering = true;
            // 
            // Label382
            // 
            this.Label382.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label382.Location = new System.Drawing.Point(461, 50);
            this.Label382.Name = "Label382";
            this.Label382.Size = new System.Drawing.Size(340, 85);
            this.Label382.TabIndex = 8;
            this.Label382.UseCompatibleTextRendering = true;
            // 
            // CheckBox49
            // 
            this.CheckBox49.AutoSize = true;
            this.CheckBox49.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CheckBox49.Location = new System.Drawing.Point(441, 25);
            this.CheckBox49.Name = "CheckBox49";
            this.CheckBox49.Size = new System.Drawing.Size(231, 22);
            this.CheckBox49.TabIndex = 7;
            this.CheckBox49.Text = "修繕項目毎契約者送金率データの確認";
            this.CheckBox49.UseVisualStyleBackColor = true;
            // 
            // Panel22
            // 
            this.Panel22.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.Panel22.Location = new System.Drawing.Point(461, 138);
            this.Panel22.Name = "Panel22";
            this.Panel22.Size = new System.Drawing.Size(340, 60);
            this.Panel22.TabIndex = 9;
            // 
            // Label381
            // 
            this.Label381.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label381.Location = new System.Drawing.Point(218, 31);
            this.Label381.Name = "Label381";
            this.Label381.Size = new System.Drawing.Size(83, 16);
            this.Label381.TabIndex = 3;
            this.Label381.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label380
            // 
            this.Label380.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label380.Location = new System.Drawing.Point(131, 31);
            this.Label380.Name = "Label380";
            this.Label380.Size = new System.Drawing.Size(92, 16);
            this.Label380.TabIndex = 2;
            this.Label380.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label379
            // 
            this.Label379.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label379.Location = new System.Drawing.Point(307, 31);
            this.Label379.Name = "Label379";
            this.Label379.Size = new System.Drawing.Size(22, 16);
            this.Label379.TabIndex = 4;
            this.Label379.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label378
            // 
            this.Label378.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label378.Location = new System.Drawing.Point(8, 8);
            this.Label378.Name = "Label378";
            this.Label378.Size = new System.Drawing.Size(321, 16);
            this.Label378.TabIndex = 0;
            this.Label378.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label378.UseCompatibleTextRendering = true;
            // 
            // lblLine0
            // 
            this.lblLine0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine0.Location = new System.Drawing.Point(319, 65);
            this.lblLine0.Name = "lblLine0";
            this.lblLine0.Size = new System.Drawing.Size(916, 2);
            this.lblLine0.TabIndex = 139;
            // 
            // pnlRefresh
            // 
            this.pnlRefresh.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnlRefresh.Controls.Add(this.Label251);
            this.pnlRefresh.Controls.Add(this.btnRefresh);
            this.pnlRefresh.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.pnlRefresh.Location = new System.Drawing.Point(316, 626);
            this.pnlRefresh.Name = "pnlRefresh";
            this.pnlRefresh.Size = new System.Drawing.Size(389, 50);
            this.pnlRefresh.TabIndex = 46;
            this.pnlRefresh.Visible = false;
            // 
            // Label251
            // 
            this.Label251.ForeColor = System.Drawing.Color.Black;
            this.Label251.Location = new System.Drawing.Point(4, 5);
            this.Label251.Name = "Label251";
            this.Label251.Size = new System.Drawing.Size(278, 38);
            this.Label251.TabIndex = 0;
            this.Label251.Text = "※表示されている抽出件数を再読み込みします。\r\n　データの調整を行った後に押して下さい。";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRefresh.Image = ((System.Drawing.Image)(resources.GetObject("btnRefresh.Image")));
            this.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRefresh.Location = new System.Drawing.Point(288, 8);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnRefresh.Size = new System.Drawing.Size(88, 30);
            this.btnRefresh.TabIndex = 2;
            this.btnRefresh.Text = " 再読込";
            this.btnRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnBack
            // 
            this.btnBack.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnBack.Font = new System.Drawing.Font("メイリオ", 11.25F);
            this.btnBack.Image = ((System.Drawing.Image)(resources.GetObject("btnBack.Image")));
            this.btnBack.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnBack.Location = new System.Drawing.Point(817, 640);
            this.btnBack.Name = "btnBack";
            this.btnBack.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnBack.Size = new System.Drawing.Size(120, 30);
            this.btnBack.TabIndex = 5;
            this.btnBack.Text = "   戻  る";
            this.btnBack.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("メイリオ", 11.25F);
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNext.Location = new System.Drawing.Point(952, 640);
            this.btnNext.Name = "btnNext";
            this.btnNext.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnNext.Size = new System.Drawing.Size(120, 30);
            this.btnNext.TabIndex = 6;
            this.btnNext.Text = "   次  へ";
            this.btnNext.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnEnd
            // 
            this.btnEnd.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnEnd.Font = new System.Drawing.Font("メイリオ", 11.25F);
            this.btnEnd.Image = ((System.Drawing.Image)(resources.GetObject("btnEnd.Image")));
            this.btnEnd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEnd.Location = new System.Drawing.Point(1089, 640);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnEnd.Size = new System.Drawing.Size(120, 30);
            this.btnEnd.TabIndex = 7;
            this.btnEnd.Text = "   終  了";
            this.btnEnd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEnd.UseVisualStyleBackColor = true;
            this.btnEnd.Click += new System.EventHandler(this.btnEnd_Click);
            // 
            // Button21
            // 
            this.Button21.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Button21.Image = ((System.Drawing.Image)(resources.GetObject("Button21.Image")));
            this.Button21.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Button21.Location = new System.Drawing.Point(8, 27);
            this.Button21.Name = "Button21";
            this.Button21.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.Button21.Size = new System.Drawing.Size(100, 25);
            this.Button21.TabIndex = 1;
            this.Button21.Text = "リスト出力";
            this.Button21.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button21.UseVisualStyleBackColor = true;
            // 
            // Button20
            // 
            this.Button20.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Button20.Image = ((System.Drawing.Image)(resources.GetObject("Button20.Image")));
            this.Button20.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Button20.Location = new System.Drawing.Point(8, 27);
            this.Button20.Name = "Button20";
            this.Button20.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.Button20.Size = new System.Drawing.Size(100, 25);
            this.Button20.TabIndex = 1;
            this.Button20.Text = "リスト出力";
            this.Button20.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button20.UseVisualStyleBackColor = true;
            // 
            // Button19
            // 
            this.Button19.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Button19.Image = ((System.Drawing.Image)(resources.GetObject("Button19.Image")));
            this.Button19.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Button19.Location = new System.Drawing.Point(8, 27);
            this.Button19.Name = "Button19";
            this.Button19.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.Button19.Size = new System.Drawing.Size(100, 25);
            this.Button19.TabIndex = 1;
            this.Button19.Text = "リスト出力";
            this.Button19.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button19.UseVisualStyleBackColor = true;
            // 
            // lblTitleH
            // 
            this.lblTitleH.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTitleH.Font = new System.Drawing.Font("メイリオ", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblTitleH.ForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.lblTitleH.Location = new System.Drawing.Point(33, 18);
            this.lblTitleH.Name = "lblTitleH";
            this.lblTitleH.Size = new System.Drawing.Size(250, 26);
            this.lblTitleH.TabIndex = 140;
            this.lblTitleH.Text = "[中間テーブル → 賃貸革命10]";
            this.lblTitleH.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pgbCheck
            // 
            this.pgbCheck.Location = new System.Drawing.Point(176, 3);
            this.pgbCheck.Name = "pgbCheck";
            this.pgbCheck.Size = new System.Drawing.Size(190, 20);
            this.pgbCheck.TabIndex = 141;
            // 
            // lblPgbCheck
            // 
            this.lblPgbCheck.Font = new System.Drawing.Font("メイリオ", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPgbCheck.Location = new System.Drawing.Point(372, 4);
            this.lblPgbCheck.Name = "lblPgbCheck";
            this.lblPgbCheck.Size = new System.Drawing.Size(51, 20);
            this.lblPgbCheck.TabIndex = 142;
            this.lblPgbCheck.Text = "0 %";
            this.lblPgbCheck.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblPgbCheck.UseCompatibleTextRendering = true;
            // 
            // lblCheckSituation
            // 
            this.lblCheckSituation.Font = new System.Drawing.Font("メイリオ", 8.25F);
            this.lblCheckSituation.Location = new System.Drawing.Point(4, 4);
            this.lblCheckSituation.Name = "lblCheckSituation";
            this.lblCheckSituation.Size = new System.Drawing.Size(166, 20);
            this.lblCheckSituation.TabIndex = 143;
            this.lblCheckSituation.Text = "...";
            this.lblCheckSituation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblCheckSituation.UseCompatibleTextRendering = true;
            // 
            // pnlPrgChk
            // 
            this.pnlPrgChk.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlPrgChk.Controls.Add(this.pgbCheck);
            this.pnlPrgChk.Controls.Add(this.lblPgbCheck);
            this.pnlPrgChk.Controls.Add(this.lblCheckSituation);
            this.pnlPrgChk.Location = new System.Drawing.Point(320, 18);
            this.pnlPrgChk.Name = "pnlPrgChk";
            this.pnlPrgChk.Size = new System.Drawing.Size(441, 29);
            this.pnlPrgChk.TabIndex = 144;
            // 
            // MainFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 682);
            this.Controls.Add(this.pnlPrgChk);
            this.Controls.Add(this.pnlMenuDatacv);
            this.Controls.Add(this.lblLine0);
            this.Controls.Add(this.pnlRefresh);
            this.Controls.Add(this.btnDevTabChange);
            this.Controls.Add(this.lblHidden1);
            this.Controls.Add(this.tabCtrlMain);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.lblTitleH);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainFrm";
            this.Text = "データコンバート";
            this.Load += new System.EventHandler(this.FrmMain_Load);
            this.pnlMenuDatacv.ResumeLayout(false);
            this.pnlPrgDCConv.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCConv2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCConv1)).EndInit();
            this.pnlPrgDCRelation.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCRelation2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCRelation1)).EndInit();
            this.pnlPrgDCFileWrite.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCFileWrite2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCFileWrite1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCEnd2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCJikko2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCSelect2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCEnd1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCJikko1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picArwDCSelect1)).EndInit();
            this.pnlPrgDCEnd.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCEnd2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCEnd1)).EndInit();
            this.pnlPrgDCJikko.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCJikko2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCJikko1)).EndInit();
            this.pnlPrgDCSelect.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCSelect2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCSelect1)).EndInit();
            this.pnlPrgDCHajimeni.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCHajimeni2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picIcoDCHajimeni1)).EndInit();
            this.tabCtrlMain.ResumeLayout(false);
            this.tabPageDev.ResumeLayout(false);
            this.GroupBox6.ResumeLayout(false);
            this.GroupBox6.PerformLayout();
            this.grpDevSettingX.ResumeLayout(false);
            this.grpDevSettingX.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox1)).EndInit();
            this.grpDevSetting1.ResumeLayout(false);
            this.grpDevSetting1.PerformLayout();
            this.grpDevSetting2.ResumeLayout(false);
            this.grpDevSetting2.PerformLayout();
            this.grpDevSetting3.ResumeLayout(false);
            this.grpDevSetting3.PerformLayout();
            this.GroupBox3.ResumeLayout(false);
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.grptaihi.ResumeLayout(false);
            this.grptaihi.PerformLayout();
            this.tabPageStart.ResumeLayout(false);
            this.grpHFirstNaiyo.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox2)).EndInit();
            this.tabPageSession.ResumeLayout(false);
            this.tabPageSession.PerformLayout();
            this.grpTimeOut.ResumeLayout(false);
            this.grpTimeOut.PerformLayout();
            this.grp10ConnectInfo.ResumeLayout(false);
            this.grp10ConnectInfo.PerformLayout();
            this.grpV10Authent.ResumeLayout(false);
            this.grpV10Authent.PerformLayout();
            this.tabPageSyoki.ResumeLayout(false);
            this.pnlUserName.ResumeLayout(false);
            this.pnlUserName.PerformLayout();
            this.pnlLogPath.ResumeLayout(false);
            this.pnlLogPath.PerformLayout();
            this.tabPageMenu.ResumeLayout(false);
            this.grpMenuJizen.ResumeLayout(false);
            this.grpMenuDatacv.ResumeLayout(false);
            this.grpMenuJigo.ResumeLayout(false);
            this.tabPageJizen.ResumeLayout(false);
            this.pnlJizenListPath.ResumeLayout(false);
            this.pnlJizenListPath.PerformLayout();
            this.tabCtrlJizen.ResumeLayout(false);
            this.tabPageHJizen1.ResumeLayout(false);
            this.GroupBox5.ResumeLayout(false);
            this.GroupBox5.PerformLayout();
            this.pnlHJizenTyukan.ResumeLayout(false);
            this.pnlHJizenTyukan.PerformLayout();
            this.tabPageJigo.ResumeLayout(false);
            this.tabCtrlJigo.ResumeLayout(false);
            this.tabPageJigo1.ResumeLayout(false);
            this.grpJigoDonyuji.ResumeLayout(false);
            this.pnlJigoCmtSoKotiku.ResumeLayout(false);
            this.pnlJigoCmtNkNyuryoku.ResumeLayout(false);
            this.pnlJigoCmtSqKotiku.ResumeLayout(false);
            this.grpJigoUserSagyo.ResumeLayout(false);
            this.grpJigoUserSagyo.PerformLayout();
            this.pnlHJigoCmtSyudo.ResumeLayout(false);
            this.tabPageHajimeni.ResumeLayout(false);
            this.GroupBox10.ResumeLayout(false);
            this.pnlDcFstCmtH99.ResumeLayout(false);
            this.pnlDcFstCmtH01.ResumeLayout(false);
            this.pnlDcFstCmtH02.ResumeLayout(false);
            this.pnlDcFstCmt11.ResumeLayout(false);
            this.pnlDcFstCmt06.ResumeLayout(false);
            this.pnlDcFstCmt05.ResumeLayout(false);
            this.pnlDcFstCmt07.ResumeLayout(false);
            this.pnlDcFstCmt10.ResumeLayout(false);
            this.pnlDcFstCmt02.ResumeLayout(false);
            this.pnlDcFstCmt08.ResumeLayout(false);
            this.pnlDcFstCmt04.ResumeLayout(false);
            this.pnlDcFstCmt03.ResumeLayout(false);
            this.pnlDcFstCmt01.ResumeLayout(false);
            this.pnlDcFstCmt09.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox14)).EndInit();
            this.tabPageSelect.ResumeLayout(false);
            this.grpMiddleFile.ResumeLayout(false);
            this.grpMiddleFile.PerformLayout();
            this.tabCtrlCVItem.ResumeLayout(false);
            this.tabPageKizon110.ResumeLayout(false);
            this.tabPageKizon110.PerformLayout();
            this.grpKizon1.ResumeLayout(false);
            this.pnlKiMstHendo.ResumeLayout(false);
            this.pnlKiMstHendo.PerformLayout();
            this.pnlKiMstTokuyaku.ResumeLayout(false);
            this.pnlKiMstTokuyaku.PerformLayout();
            this.pnlKiMstKasyoClaimrui.ResumeLayout(false);
            this.pnlKiMstKasyoClaimrui.PerformLayout();
            this.pnlKiMstTitle.ResumeLayout(false);
            this.pnlKiMstTitle.PerformLayout();
            this.pnlKiMstArea.ResumeLayout(false);
            this.pnlKiMstArea.PerformLayout();
            this.pnlKiMstSchool.ResumeLayout(false);
            this.pnlKiMstSchool.PerformLayout();
            this.pnlKiMstHokenrui.ResumeLayout(false);
            this.pnlKiMstHokenrui.PerformLayout();
            this.pnlKiMstBus.ResumeLayout(false);
            this.pnlKiMstBus.PerformLayout();
            this.tabPageKizon120.ResumeLayout(false);
            this.tabPageKizon120.PerformLayout();
            this.grpKizon3.ResumeLayout(false);
            this.pnlKiGySyuzenBase.ResumeLayout(false);
            this.pnlKiGySyuzenBase.PerformLayout();
            this.pnlKiGyYatinhosyoBase.ResumeLayout(false);
            this.pnlKiGyYatinhosyoBase.PerformLayout();
            this.pnlKiGyLifelineBase.ResumeLayout(false);
            this.pnlKiGyLifelineBase.PerformLayout();
            this.pnlKiGyHokenBase.ResumeLayout(false);
            this.pnlKiGyHokenBase.PerformLayout();
            this.pnlKiGySisetuBase.ResumeLayout(false);
            this.pnlKiGySisetuBase.PerformLayout();
            this.pnlKiGyCyukaiBase.ResumeLayout(false);
            this.pnlKiGyCyukaiBase.PerformLayout();
            this.pnlKiGySekoBase.ResumeLayout(false);
            this.pnlKiGySekoBase.PerformLayout();
            this.grpKizon5.ResumeLayout(false);
            this.pnlKiOw.ResumeLayout(false);
            this.pnlKiOw.PerformLayout();
            this.pnlKiJisya.ResumeLayout(false);
            this.pnlKiJisya.PerformLayout();
            this.pnlKiSyskanriBase.ResumeLayout(false);
            this.pnlKiSyskanriBase.PerformLayout();
            this.tabPageKizon130.ResumeLayout(false);
            this.tabPageKizon130.PerformLayout();
            this.grpKizon2.ResumeLayout(false);
            this.pnlRendo.ResumeLayout(false);
            this.pnlRendo.PerformLayout();
            this.pnlKiSq.ResumeLayout(false);
            this.pnlKiSq.PerformLayout();
            this.pnlSzen.ResumeLayout(false);
            this.pnlSzen.PerformLayout();
            this.pnlKiClaim.ResumeLayout(false);
            this.pnlKiClaim.PerformLayout();
            this.pnlKiKy.ResumeLayout(false);
            this.pnlKiKy.PerformLayout();
            this.pnlKiKys.ResumeLayout(false);
            this.pnlKiKys.PerformLayout();
            this.pnlKiHy.ResumeLayout(false);
            this.pnlKiHy.PerformLayout();
            this.pnlKiBk.ResumeLayout(false);
            this.pnlKiBk.PerformLayout();
            this.grpKizonKagi.ResumeLayout(false);
            this.grpKizonKagi.PerformLayout();
            this.tabPageBase110.ResumeLayout(false);
            this.grpMst.ResumeLayout(false);
            this.grpMst.PerformLayout();
            this.tabPageBase120.ResumeLayout(false);
            this.grpGy.ResumeLayout(false);
            this.grpGy.PerformLayout();
            this.tabPageBase130.ResumeLayout(false);
            this.grpKys.ResumeLayout(false);
            this.grpKys.PerformLayout();
            this.grpOw.ResumeLayout(false);
            this.grpOw.PerformLayout();
            this.grpJisya.ResumeLayout(false);
            this.grpJisya.PerformLayout();
            this.tabPageBase140.ResumeLayout(false);
            this.grpBk.ResumeLayout(false);
            this.grpBk.PerformLayout();
            this.grpKagiSelect.ResumeLayout(false);
            this.grpKagiSelect.PerformLayout();
            this.tabPageBase150.ResumeLayout(false);
            this.grpHy.ResumeLayout(false);
            this.grpHy.PerformLayout();
            this.tabPageBase160.ResumeLayout(false);
            this.grpSorule.ResumeLayout(false);
            this.grpSorule.PerformLayout();
            this.tabPageBase170.ResumeLayout(false);
            this.grpKy.ResumeLayout(false);
            this.grpKy.PerformLayout();
            this.tabPageBase180.ResumeLayout(false);
            this.grpSq.ResumeLayout(false);
            this.grpSq.PerformLayout();
            this.tabPageBase190.ResumeLayout(false);
            this.grpSzen.ResumeLayout(false);
            this.grpSzen.PerformLayout();
            this.grpClaim.ResumeLayout(false);
            this.grpClaim.PerformLayout();
            this.tabPageBase200.ResumeLayout(false);
            this.grpSyskanri.ResumeLayout(false);
            this.grpSyskanri.PerformLayout();
            this.tabPageBase210.ResumeLayout(false);
            this.grpRendo.ResumeLayout(false);
            this.grpRendo.PerformLayout();
            this.tabPageBase900.ResumeLayout(false);
            this.tabPageBase900.PerformLayout();
            this.tabPageHanyo110.ResumeLayout(false);
            this.grpHMst.ResumeLayout(false);
            this.grpHMst.PerformLayout();
            this.tabPageHanyo120.ResumeLayout(false);
            this.grpHGy.ResumeLayout(false);
            this.grpHGy.PerformLayout();
            this.grpHOw.ResumeLayout(false);
            this.grpHOw.PerformLayout();
            this.grpHJisya.ResumeLayout(false);
            this.grpHJisya.PerformLayout();
            this.TabPage1.ResumeLayout(false);
            this.grpKizon4.ResumeLayout(false);
            this.grpKizon4.PerformLayout();
            this.pnlRekiClear.ResumeLayout(false);
            this.tabPageJikko.ResumeLayout(false);
            this.grpTotalProcess.ResumeLayout(false);
            this.grpTotalProcess.PerformLayout();
            this.tabPageEndOK.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox15)).EndInit();
            this.tabPageEndError.ResumeLayout(false);
            this.GroupBox23.ResumeLayout(false);
            this.GroupBox19.ResumeLayout(false);
            this.GroupBox20.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox23)).EndInit();
            this.tabPageEndCancel.ResumeLayout(false);
            this.GroupBox21.ResumeLayout(false);
            this.GroupBox22.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox25)).EndInit();
            this.pnlRefresh.ResumeLayout(false);
            this.pnlPrgChk.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        internal Button btnConnectTest;
        internal Button btnAllChk;
        internal CheckBox chkRelationFile;
        internal Label Label9;
        internal Button btnMidFileCheck;
        internal Button btnDefConInfoRead;
        internal Panel pnlMenuDatacv;
        internal PictureBox picArwDCEnd2;
        internal PictureBox picArwDCJikko2;
        internal PictureBox picArwDCSelect2;
        internal PictureBox picArwDCEnd1;
        internal PictureBox picArwDCJikko1;
        internal PictureBox picArwDCSelect1;
        internal Panel pnlPrgDCEnd;
        internal PictureBox picIcoDCEnd2;
        internal PictureBox picIcoDCEnd1;
        internal Label lblPrgDCEnd;
        internal Panel pnlPrgDCJikko;
        internal PictureBox picIcoDCJikko2;
        internal PictureBox picIcoDCJikko1;
        internal Label lblPrgDCJikko;
        internal Panel pnlPrgDCSelect;
        internal PictureBox picIcoDCSelect2;
        internal PictureBox picIcoDCSelect1;
        internal Label lblPrgDCSelect;
        internal Panel pnlPrgDCHajimeni;
        internal PictureBox picIcoDCHajimeni2;
        internal PictureBox picIcoDCHajimeni1;
        internal Label lblPrgDCHajimeni;
        internal Button btnBack;
        internal Button btnEnd;
        private Button btnNext;
        internal Label lblHidden1;
        internal TabControl tabCtrlMain;
        internal TabPage tabPageHajimeni;
        internal PictureBox PictureBox14;
        internal Label lblDatacvHajimeniDescription1;
        internal GroupBox GroupBox10;
        internal Label Label126;
        internal Label Label132;
        internal Label Label122;
        internal Label Label104;
        internal Label Label123;
        internal Label Label105;
        internal Label Label127;
        internal Label Label129;
        internal Label Label133;
        internal Label Label135;
        internal TabPage tabPageSyoki;
        internal TextBox txtRecUser;
        internal Label lblExecutor;
        internal Button btnLogDirSeach;
        internal TextBox txtLogDirPath;
        internal Label lblLogDir;
        internal Label lblSyokiDescription1;
        internal TabPage tabPageSession;
        internal Label lblSessionCaution;
        internal Label lblTimeOutSec;
        internal TextBox txtTimeOut;
        internal Label Label36;
        internal Label Label35;
        internal Label lblNetworklib;
        internal Label Label30;
        internal Label Label13;
        internal GroupBox grp10ConnectInfo;
        internal GroupBox grpV10Authent;
        internal RadioButton optV10Authent2;
        internal RadioButton optV10Authent1;
        internal TextBox txtV10Pass;
        internal TextBox txtV10User;
        internal TextBox txtV10Catalog;
        internal TextBox txtV10Server;
        internal Label lblV10;
        internal TabPage tabPageJizen;
        internal TabControl tabCtrlJizen;
        internal Label lblCautionDescription;
        internal Label lblJizenDescription1;
        internal TabPage tabPageSelect;
        internal Button btnRekiClear;
        internal GroupBox grpMiddleFile;
        internal Button btnMidDirLogSeach;
        internal TextBox txtMidDirLogPath;
        internal Label lblMiddleFileLog;
        internal TextBox txtMidDirPath;
        internal Label lblDatacvSelectCaution;
        internal Label lblDatacvSelectDescription1;
        internal TabControl tabCtrlCVItem;
        internal TabPage tabPageKizon110;
        internal GroupBox grpKizon1;
        internal CheckBox chkKiMstHendo;
        internal CheckBox chkKiMstKasyoClaimrui;
        internal CheckBox chkKiMstBus;
        internal CheckBox chkKiMstSchool;
        internal CheckBox chkKiMstTokuyaku;
        internal CheckBox chkKiMstHokenrui;
        internal CheckBox chkKiMstTitle;
        internal CheckBox chkKiMstArea;
        internal TabPage tabPageKizon120;
        internal TabPage tabPageKizon130;
        internal TabPage tabPageBase110;
        internal Label lblRelItemInfo;
        internal GroupBox grpMst;
        internal CheckBox chkMstHendoitiran;
        internal CheckBox chkMstHendo;
        internal CheckBox chkMstKasyoClaimrui;
        internal CheckBox chkMstBus;
        internal CheckBox chkMstSchool;
        internal CheckBox chkMstTokuyaku;
        internal CheckBox chkMstHokenrui;
        internal CheckBox chkMstBusKotu;
        internal CheckBox chkMstKagititle;
        internal CheckBox chkMstArea;
        internal TabPage tabPageBase120;
        internal GroupBox grpGy;
        internal CheckBox chkGySyuzenMemo;
        internal CheckBox chkGyHokenMemo;
        internal CheckBox chkGySyuzenKoza;
        internal CheckBox chkGySekoBase;
        internal CheckBox chkGySisetuBase;
        internal CheckBox chkGySyuzenBase;
        internal CheckBox chkGyYatinhosyoMemo;
        internal CheckBox chkGyYatinhosyoBase;
        internal CheckBox chkGyCyukaiKoza;
        internal CheckBox chkGyCyukaiMemo;
        internal CheckBox chkGyHokenBase;
        internal CheckBox chkGyLifelineBase;
        internal CheckBox chkGyHokenKoza;
        internal CheckBox chkGyCyukaiBase;
        internal TabPage tabPageBase130;
        internal GroupBox grpKys;
        internal CheckBox chkKysHosyonin;
        internal CheckBox chkKysSyogoKana;
        internal CheckBox chkKysKoza;
        internal CheckBox chkKysMemo;
        internal CheckBox chkKysBase;
        internal GroupBox grpOw;
        internal CheckBox chkOwEvent;
        internal CheckBox chkOwKoza;
        internal CheckBox chkOwMemo;
        internal CheckBox chkOwBase;
        internal GroupBox grpJisya;
        internal CheckBox chkFBKozafurikae;
        internal CheckBox chkFBFuriirai;
        internal CheckBox chkJisyaKoza;
        internal CheckBox chkJisyaMemo;
        internal CheckBox chkJisyaBase;
        internal TabPage tabPageBase140;
        internal GroupBox grpBk;
        internal CheckBox chkBkSyo;
        internal CheckBox chkBkHendo;
        internal CheckBox chkBkKinrincyusyajo;
        internal CheckBox chkBkSansyofile;
        internal CheckBox chkBkSzeniji;
        internal CheckBox chkBkSyuhen;
        internal CheckBox chkBkSetudo;
        internal CheckBox chkBkKotu;
        internal CheckBox chkBkKenri;
        internal CheckBox chkBkGomi;
        internal CheckBox chkBkSyosai;
        internal CheckBox chkBkKagi;
        internal CheckBox chkBkMemo;
        internal CheckBox chkBkBase;
        internal TabPage tabPageBase150;
        internal GroupBox grpHy;
        internal CheckBox chkHySyo;
        internal CheckBox chkHySansyofile;
        internal CheckBox chkHyGenjotanka;
        internal CheckBox chkHyKenri;
        internal CheckBox chkHyConfirm;
        internal CheckBox chkHyHendo;
        internal CheckBox chkHyCommonsalespoint;
        internal CheckBox chkHyMenseki;
        internal CheckBox chkHyNkinkomk;
        internal CheckBox chkHyMadoriutiwake;
        internal CheckBox chkHySzeniji;
        internal CheckBox chkHyTokuyaku;
        internal CheckBox chkHyParking;
        internal CheckBox chkHySyosai;
        internal CheckBox chkHyMemo;
        internal CheckBox chkHyKagi;
        internal CheckBox chkHySetubi;
        internal CheckBox chkHyBase;
        internal TabPage tabPageBase160;
        internal GroupBox grpSorule;
        internal CheckBox chkSoruleSosaki;
        internal CheckBox chkSoruleKojo;
        internal CheckBox chkSoruleNkin;
        internal CheckBox chkSoruleBase;
        internal TabPage tabPageBase170;
        internal GroupBox grpKy;
        internal CheckBox chkKyKai;
        internal CheckBox CheckBox14;
        internal CheckBox CheckBox24;
        internal CheckBox CheckBox23;
        internal CheckBox chkKyHosyonin;
        internal CheckBox chkKyMemo;
        internal CheckBox chkKyTokuyaku;
        internal CheckBox CheckBox18;
        internal CheckBox chkKySorule;
        internal CheckBox CheckBox16;
        internal CheckBox CheckBox15;
        internal CheckBox chkKyNyukyo;
        internal CheckBox chkKyNkinkomkNx;
        internal CheckBox chkKyNkinkomk;
        internal CheckBox chkKyHoken;
        internal CheckBox chkKyKojoRule;
        internal CheckBox chkKyKys;
        internal CheckBox CheckBox5;
        internal CheckBox chkKyRireki;
        internal CheckBox chkKyHendo;
        internal CheckBox chkKyCar;
        internal CheckBox chkKyBase;
        internal TabPage tabPageBase180;
        internal GroupBox grpSq;
        internal CheckBox chkSqSqKojo;
        internal CheckBox chkSqKoteiKojo;
        internal CheckBox chkSqHendokensin;
        internal CheckBox chkSqUnyotaino;
        internal CheckBox chkSqSq;
        internal CheckBox chkSqKajyo;
        internal TabPage tabPageBase900;
        internal TabPage tabPageHanyo110;
        internal GroupBox grpHMst;
        internal Label Label138;
        internal Label Label150;
        internal Label Label151;
        internal Label Label152;
        internal Label Label153;
        internal Label Label154;
        internal Label Label155;
        internal Label Label156;
        internal Label Label157;
        internal CheckBox CheckBox2;
        internal CheckBox CheckBox3;
        internal CheckBox CheckBox6;
        internal CheckBox CheckBox9;
        internal CheckBox CheckBox10;
        internal CheckBox CheckBox11;
        internal CheckBox CheckBox12;
        internal CheckBox CheckBox13;
        internal CheckBox CheckBox17;
        internal TabPage tabPageHanyo120;
        internal GroupBox grpHGy;
        internal CheckBox CheckBox4;
        internal CheckBox CheckBox30;
        internal CheckBox CheckBox31;
        internal CheckBox CheckBox32;
        internal CheckBox CheckBox33;
        internal CheckBox CheckBox34;
        internal CheckBox CheckBox35;
        internal GroupBox grpHOw;
        internal CheckBox CheckBox1;
        internal CheckBox CheckBox26;
        internal CheckBox CheckBox27;
        internal CheckBox CheckBox28;
        internal CheckBox CheckBox29;
        internal GroupBox grpHJisya;
        internal CheckBox CheckBox19;
        internal CheckBox CheckBox20;
        internal CheckBox CheckBox21;
        internal CheckBox CheckBox22;
        internal CheckBox CheckBox25;
        internal TabPage tabPageHanyo130;
        internal TabPage tabPageHanyo140;
        internal TabPage tabPageHanyo150;
        internal TabPage tabPageJikko;
        internal Label lblDatacvJikkoCaution;
        internal Label lblDatacvJikkoDescription1;
        internal Label lblPgbPartial;
        internal TextBox txtPartialSituation;
        internal ProgressBar pgbpartial;
        internal Label lblCVItem;
        internal Label Label31;
        internal GroupBox grpTotalProcess;
        internal Label lblPgbTotal;
        internal TextBox txtTotalSituation;
        internal ProgressBar pgbTotal;
        internal Label Label32;
        internal Label lblTotalSituation;
        internal TabPage tabPageEndOK;
        internal PictureBox PictureBox15;
        internal Label lblDatacvEndDescription1;
        internal TabPage tabPageDev;
        internal GroupBox grpDevSetting1;
        internal Label Label51;
        internal CheckBox chkKyNewest;
        internal Label Label48;
        internal CheckBox chkChildItemControl;
        internal Label Label46;
        internal Label Label45;
        internal Label Label44;
        internal CheckBox chkRelTblDrop;
        internal Label Label43;
        internal Label Label42;
        internal CheckBox chkV7ViewDrop;
        internal CheckBox chkLogTblDrop;
        internal CheckBox chkOverWrite;
        internal Label Label52;
        internal GroupBox grpDevSetting2;
        internal Button btnRelDirSeach;
        internal TextBox txtRelationDirPath;
        internal Label lblRelationDir;
        internal Label Label11;
        internal CheckBox chkRelation;
        internal Label Label47;
        internal CheckBox chkMidNotStop;
        internal CheckBox chkCVStart;
        internal Label Label8;
        internal Label Label10;
        internal CheckBox chkMiddleFile;
        internal GroupBox grpDevSetting3;
        internal Label Label55;
        internal Label Label54;
        internal CheckBox chkRommDuplicate;
        internal CheckBox chkEmptyRoomNo;
        internal Label Label53;
        internal Label Label50;
        internal Label Label49;
        internal Label Label41;
        internal TextBox txtLogOutputCnt;
        internal GroupBox grpRelation;
        internal GroupBox GroupBox3;
        internal Label Label39;
        internal Label Label33;
        internal Label Label6;
        internal Label Label56;
        internal Label Label16;
        internal Label Label15;
        internal Label Label14;
        internal Label Label12;
        internal Label Label18;
        internal GroupBox GroupBox2;
        internal CheckBox chkMstKasyorui;
        internal CheckBox chkMstGenjotokuyaku;
        internal Label lblTokuyakuInfo;
        internal GroupBox GroupBox1;
        internal Label Label34;
        internal Label Label29;
        internal CheckBox chkMstKozasyubetu;
        internal CheckBox chkMstSetubi;
        internal CheckBox chkMstKozo;
        internal CheckBox chkMstTorihikitaiyo;
        internal CheckBox chkMstNkinkbn;
        internal CheckBox chkMstHyrui;
        internal CheckBox chkMstNkinkomok;
        internal CheckBox chkMstBkrui;
        internal GroupBox grptaihi;
        internal Label Label7;
        internal Button btnFileSeach;
        internal Label Label5;
        internal CheckBox chkGyYatinhosyoKoza;
        internal CheckBox CheckBox66;
        internal Label lblGy;
        internal Label lblMst;
        internal Label lblKys;
        internal Label lblOwner;
        internal Label lblJisya;
        internal Label Label17;
        internal Label Label19;
        internal Label Label20;
        internal CheckBox chkMs25;
        internal CheckBox chkMsKeiyakusyubetu;
        internal CheckBox chkMs23;
        internal CheckBox chkMs24;
        internal CheckBox chkMsOwevent;
        internal CheckBox chkMsYane;
        internal CheckBox chkMsKeikaikakunin;
        internal CheckBox chkMsKinyu;
        internal CheckBox chkMsKinyuten;
        internal TabPage tabPageIkkatu;
        internal TabPage tabPageHanyoJizen;
        internal TabPage tabPageEndError;
        internal GroupBox GroupBox23;
        internal Label Label163;
        internal PictureBox PictureBox23;
        internal Label lblDatacvErrorDescription1;
        internal GroupBox GroupBox19;
        internal Label Label170;
        internal GroupBox GroupBox20;
        internal Label Label176;
        internal Label lblDatacvErrorLabel;
        internal TabPage tabPageEndCancel;
        internal PictureBox PictureBox25;
        internal Label lblDatacvCancelDescription1;
        internal GroupBox GroupBox21;
        internal Label Label186;
        internal GroupBox GroupBox22;
        internal Label Label192;
        internal Label lblDatacvCancelLabel;
        internal GroupBox grpKagiSelect;
        internal RadioButton optKyKagi;
        internal RadioButton optHyKagi;
        internal Button btnDevTabChange;
        internal Label Label22;
        internal Label lblHidden2;
        internal Label lblLine2;
        internal Label lblLine1;
        internal Label lblHidden3;
        internal GroupBox grpTimeOut;
        internal Label Label119;
        internal Label Label116;
        internal Button btnDoui;
        internal PictureBox PictureBox1;
        internal CheckBox chkOpOw;
        internal CheckBox chkOpKys;
        internal CheckBox chkMstYatinKoza;
        internal CheckBox chkFBNsSyutoku;
        internal CheckBox chkFBFuritesuryo;
        internal TabPage tabPageBase190;
        internal GroupBox grpClaim;
        internal CheckBox chkClaimTaiorireki;
        internal CheckBox chkClaimRelfile;
        internal CheckBox chkClaimBase;
        internal TabPage tabPageBase200;
        internal GroupBox grpSyskanri;
        internal CheckBox chkSyskanriBase;
        internal CheckBox chkJisyaTanto;
        internal CheckBox chkSyskanriNkinkomkmerge;
        internal CheckBox chkSyskanriHenkanmoji;
        internal CheckBox chkSyskanriZei;
        internal CheckBox chkKySzenmeisai;
        internal CheckBox chkKySzen;
        internal GroupBox grpSzen;
        internal CheckBox chkSzenMemo;
        internal CheckBox chkSzenSzen;
        internal CheckBox chkSzenSzenmeisai;
        internal CheckBox chkSzenBase;
        internal CheckBox chkSzenClaim;
        internal CheckBox chkMstBikotitle;
        internal CheckBox chkMstBikolst;
        internal Label Label101;
        internal Label Label97;
        internal TabPage TabPage1;
        internal GroupBox grpKizon4;
        internal Label lblSelectPageCnt3;
        internal Label Label26;
        internal Label Label73;
        internal Label Label70;
        internal Label Label71;
        internal Label Label72;
        internal Label Label74;
        internal CheckBox chkHSongai;
        internal CheckBox chkHJisyaKoza;
        internal CheckBox chkHSetubi;
        internal CheckBox chkHDataFmt;
        internal Label Label69;
        internal Label Label68;
        internal Label Label67;
        internal Label Label66;
        internal Label Label65;
        internal Label Label64;
        internal Label Label63;
        internal Label Label62;
        internal Label Label61;
        internal Label Label60;
        internal Label Label59;
        internal CheckBox chkHNkinKomk;
        internal CheckBox chkHTaiyo;
        internal CheckBox chkHKozo;
        internal CheckBox chkHKozaSyu;
        internal CheckBox chkHEki;
        internal CheckBox chkHEnsen;
        internal CheckBox chkHKinyuSiten;
        internal CheckBox chkHYouto;
        internal CheckBox chkHBkBunrui;
        internal CheckBox chkHKinyu;
        internal CheckBox chkHKyBunrui;
        internal CheckBox chkHKagi;
        internal CheckBox chkHHouKenri;
        internal CheckBox chkHNkinKbn;
        internal CheckBox chkHHyBunrui;
        internal GroupBox grpKizon2;
        internal CheckBox chkKiSzenBase;
        internal CheckBox chkKiClaimBase;
        internal CheckBox chkKiSqBase;
        internal GroupBox grpKizonKagi;
        internal RadioButton optKiKyKagi;
        internal RadioButton optKiHyKagi;
        internal Label Label86;
        internal CheckBox chkKiKysBase;
        internal CheckBox chkKiKyBase;
        internal CheckBox chkKiHySetubi;
        internal CheckBox chkKiHyBase;
        internal CheckBox chkKiBkBase;
        internal Label Label78;
        internal Label Label222;
        internal Label Label223;
        internal Label lblKiMstTokuyakuCnt;
        internal Label Label218;
        internal Label Label219;
        internal Label Label220;
        internal Label lblKiMstHokenruiCnt;
        internal Label Label76;
        internal Label Label77;
        internal Label Label216;
        internal Label lblKiMstAreaCnt;
        internal Label Label75;
        internal Label Label213;
        internal Label Label214;
        internal Label lblKiMstSchoolCnt;
        internal Label Label210;
        internal Label Label181;
        internal Label Label211;
        internal Label lblKiMstBusCnt;
        internal GroupBox grpKizon3;
        internal CheckBox chkKiGySekoBase;
        internal CheckBox chkKiGySisetuBase;
        internal CheckBox chkKiGySyuzenBase;
        internal CheckBox chkKiGyYatinhosyoBase;
        internal CheckBox chkKiGyHokenBase;
        internal CheckBox chkKiGyLifelineBase;
        internal CheckBox chkKiGyCyukaiBase;
        internal Label Label226;
        internal Label Label227;
        internal Label Label228;
        internal Label lblKiMstKasyoClaimruiCnt;
        internal Label Label80;
        internal Label Label232;
        internal Label Label233;
        internal Label lblKiMstKagititleCnt;
        internal Label Label81;
        internal Label Label82;
        internal Label Label230;
        internal Label lblKiMstHendoCnt;
        internal Label Label255;
        internal Label Label256;
        internal Label lblKiGySekoBaseCnt;
        internal Label Label252;
        internal Label Label253;
        internal Label lblKiGySisetuBaseCnt;
        internal Label Label249;
        internal Label Label250;
        internal Label lblKiGyYatinhosyoBaseCnt;
        internal Label Label246;
        internal Label Label247;
        internal Label lblKiGyHokenBaseCnt;
        internal Label Label243;
        internal Label Label244;
        internal Label lblKiGyLifelineBaseCnt;
        internal Label Label240;
        internal Label Label241;
        internal Label lblKiGySyuzenBaseCnt;
        internal Label Label87;
        internal Label Label238;
        internal Label lblKiGyCyukaiBaseCnt;
        internal GroupBox grpKizon5;
        internal CheckBox chkKiSyskanriBase;
        internal CheckBox chkKiJisyaBase;
        internal CheckBox chkKiOwBase;
        internal Label Label267;
        internal Label Label268;
        internal Label lblKiOwBaseCnt;
        internal Label Label263;
        internal Label Label264;
        internal Label Label265;
        internal Label lblKiJisyaBaseCnt;
        internal Label Label258;
        internal Label Label259;
        internal Label Label260;
        internal Label lblKiSyskanriBaseCnt;
        internal Label lblSelectPageCnt1;
        internal Label lblSelectPageCnt2;
        internal Label Label272;
        internal Label Label261;
        internal Label Label174;
        internal Label Label294;
        internal Label Label295;
        internal Label lblKiSzenBaseCnt;
        internal Label Label180;
        internal Label Label292;
        internal Label lblKiClaimBaseCnt;
        internal Label Label288;
        internal Label Label289;
        internal Label Label290;
        internal Label lblKiSqMiBaseCnt;
        internal Label Label284;
        internal Label Label285;
        internal Label Label286;
        internal Label lblKiKyBaseCnt;
        internal Label Label280;
        internal Label Label281;
        internal Label Label282;
        internal Label lblKiKysBaseCnt;
        internal Label Label277;
        internal Label Label278;
        internal Label lblKiHySetubiCnt;
        internal Label Label270;
        internal Label Label275;
        internal Label lblKiHyBaseCnt;
        internal Label Label271;
        internal Label Label273;
        internal Label lblKiBkBaseCnt;
        internal Panel pnlRekiClear;
        internal Label Label79;
        internal Label Label298;
        internal TabPage tabPageMenu;
        internal TabPage tabPageStart;
        internal Label lblFirstDescription1;
        internal Button btnJizenListDirSeach;
        internal TextBox txtJizenListPath;
        internal TabPage tabPageJigo;
        internal Label Label326;
        internal Label Label325;
        internal Label lblMenuCaution;
        internal GroupBox grpMenuJizen;
        internal Button btnMenuJizen;
        internal Label lblKMenuJizen;
        internal GroupBox grpMenuDatacv;
        internal Button btnMenuDatacv;
        internal Label lblKMenuDatacv;
        internal GroupBox grpMenuJigo;
        internal Button btnMenuJigo;
        internal Label lblMenuJigo;
        internal Label lblMenuDescription1;
        internal TabControl tabCtrlJigo;
        internal TabPage tabPageJigo1;
        internal Label lblJigoDescription1;
        internal TabPage tabPageHojyo;
        internal Label Label393;
        internal Label Label392;
        internal CheckBox CheckBox51;
        internal Label Label391;
        internal CheckBox CheckBox50;
        internal Panel Panel24;
        internal Label Label390;
        internal Label Label389;
        internal Label Label388;
        internal Label Label387;
        internal Button Button21;
        internal Panel Panel23;
        internal Label Label386;
        internal Label Label385;
        internal Label Label384;
        internal Label Label383;
        internal Button Button20;
        internal Label Label382;
        internal CheckBox CheckBox49;
        internal Panel Panel22;
        internal Label Label381;
        internal Label Label380;
        internal Label Label379;
        internal Label Label378;
        internal Button Button19;
        internal GroupBox grpJigoUserSagyo;
        internal PictureBox PictureBox2;
        internal Label lblHidden4;
        internal Label Label57;
        internal Label Label58;
        internal Label lblJizenPageNum1;
        internal CheckBox chkFBANSERSetuzoku;
        internal CheckBox chkMstANSERArea;
        internal CheckBox chkMstANSERAccpoint;
        internal CheckBox chkSzenRelfile;
        internal Label Label160;
        internal Label Label162;
        internal Label lblKiSqAzBaseCnt;
        internal Panel pnlKiKys;
        internal Panel pnlKiHy;
        internal Panel pnlKiBk;
        internal Panel pnlKiKy;
        internal Panel pnlKiSq;
        internal Panel pnlSzen;
        internal Panel pnlKiClaim;
        internal Panel pnlKiGyCyukaiBase;
        internal Panel pnlKiOw;
        internal Panel pnlKiJisya;
        internal Panel pnlKiSyskanriBase;
        internal Panel pnlKiGySyuzenBase;
        internal Panel pnlKiGyYatinhosyoBase;
        internal Panel pnlKiGyLifelineBase;
        internal Panel pnlKiGyHokenBase;
        internal Panel pnlKiGySisetuBase;
        internal Panel pnlKiGySekoBase;
        internal Panel pnlKiMstBus;
        internal Panel pnlKiMstHokenrui;
        internal Panel pnlKiMstTokuyaku;
        internal Panel pnlKiMstKasyoClaimrui;
        internal Panel pnlKiMstTitle;
        internal Panel pnlKiMstArea;
        internal Panel pnlKiMstSchool;
        internal Panel pnlKiMstHendo;
        internal ComboBox cmbV10Networklib;
        internal Label lblLine3;
        internal Label lblLine0;
        internal Label Label245;
        internal Panel pnlRefresh;
        internal Label Label251;
        internal Button btnRefresh;
        internal Panel pnlJizenListPath;
        internal Label Label254;
        internal TabPage tabPageBase210;
        internal GroupBox grpRendo;
        internal CheckBox chkRendoSosinJisyaweb;
        internal CheckBox chkRendoSosinBase;
        internal Panel pnlRendo;
        internal CheckBox chkKiRendoBase;
        internal Label lblKiRendoBaseCnt;
        internal Label Label314;
        internal Label Label316;
        internal Label lblHFirstLabel;
        internal GroupBox grpHFirstNaiyo;
        internal Label lblHajimeH342;
        internal Label lblHajimeH322;
        internal Label lblHajimeH341;
        internal Label lblHajimeH321;
        internal Label lblHajimeH302;
        internal Label lblHajimeH202;
        internal Label lblHajimeH102;
        internal Label lblHajimeH301;
        internal Label lblHajimeH201;
        internal Label lblHajimeH101;
        internal Label lblHSessionDescription1;
        internal Label lblHMenuDatacv;
        internal Label Label121;
        internal Panel pnlHJigoCmtSyudo;
        internal Label Label136;
        internal Label Label317;
        internal Label Label318;
        internal Panel pnlDcFstCmt07;
        internal Panel pnlDcFstCmt10;
        internal Panel pnlDcFstCmt02;
        internal Panel pnlDcFstCmt08;
        internal Panel pnlDcFstCmt04;
        internal Panel pnlDcFstCmt03;
        internal Panel pnlDcFstCmt01;
        internal Panel pnlDcFstCmt09;
        internal Label lblHDatacvHajimeniLabel;
        internal Label Label124;
        internal Label Label328;
        internal Label lblHDatacvEndLabel;
        internal TabPage tabPageHJizen1;
        internal GroupBox GroupBox5;
        internal Panel pnlHJizenTyukan;
        internal CheckBox chkHJizenTyukan;
        internal Label Label355;
        internal Label lblHajimeH312;
        internal Label lblHajimeH311;
        internal CheckBox chkRendoSosinSuumo;
        internal CheckBox chkRendoSosinAthome;
        internal CheckBox chkRendoSosinHomes;
        internal Label lblKiMstBikotitleCnt;
        internal Label Label236;
        internal Label Label235;
        internal Label lblKiMstGazotitleCnt;
        internal Label Label346;
        internal Label Label347;
        internal CheckBox chkMstGazotitle;
        internal CheckBox chkRendoKokokuSuumo;
        internal CheckBox chkRendoKokokuAthome;
        internal CheckBox chkRendoKokokuHomes;
        internal CheckBox chkRendoKokokuJisyaweb;
        internal CheckBox chkRendoHyrui;
        internal CheckBox chkRendoHysosin;
        internal CheckBox chkRendoBtoBgroup;
        internal CheckBox chkRendoMapdisp;
        internal Panel pnlPrgDCConv;
        internal PictureBox picIcoDCConv2;
        internal PictureBox picIcoDCConv1;
        internal Label lblPrgDCConv;
        internal Panel pnlPrgDCRelation;
        internal PictureBox picIcoDCRelation2;
        internal PictureBox picIcoDCRelation1;
        internal Label lblPrgDCRelation;
        internal Panel pnlPrgDCFileWrite;
        internal PictureBox picIcoDCFileWrite2;
        internal PictureBox picIcoDCFileWrite1;
        internal Label lblPrgDCFileWrite;
        internal Label Label349;
        internal CheckBox chkSyskanriinit;
        internal Label Label83;
        internal TextBox txtRendoID;
        internal Label Label359;
        internal Label Label360;
        internal Button btnExistMidToBaseMid;
        internal GroupBox grpExistMidToBaseMid;
        internal Label Label361;
        internal TextBox txtExistMidToBaseMid;
        internal Button btnExistMidToBaseMidDirSerach;
        internal Panel pnlDcFstCmt06;
        internal Label Label364;
        internal Label Label369;
        internal Panel pnlDcFstCmt05;
        internal Label Label362;
        internal Label Label363;
        internal Panel pnlUserName;
        internal Panel pnlLogPath;
        internal Label lblTitleH;
        internal Panel pnlDcFstCmt11;
        internal Label Label21;
        internal Label Label25;
        internal Panel pnlDcFstCmtH02;
        internal Label Label27;
        internal Label Label28;
        internal Panel pnlDcFstCmtH01;
        internal Label Label158;
        internal Label Label159;
        internal Panel pnlDcFstCmtH99;
        internal Label Label301;
        internal Label Label302;
        internal ProgressBar pgbCheck;
        internal Label lblPgbCheck;
        internal Label lblCheckSituation;
        internal Panel pnlPrgChk;
        internal Label Label308;
        internal Label lblKiSqOwKojoBaseCnt;
        internal Label Label311;
        internal GroupBox grpDevSettingX;
        internal GroupBox GroupBox6;
        internal Label lblHMenuJizen;
        internal GroupBox grpJigoDonyuji;
        internal Panel pnlJigoCmtSoKotiku;
        internal Label Label109;
        internal Label Label108;
        internal Label Label142;
        internal Panel pnlJigoCmtNkNyuryoku;
        internal Label Label99;
        internal Label Label111;
        internal Label Label110;
        internal Panel pnlJigoCmtSqKotiku;
        internal Label Label113;
        internal Label Label112;
        internal Label Label24;
        private TextBox pre_table_connection_string;
        private Label label1;
        private Label label2;
    }

}