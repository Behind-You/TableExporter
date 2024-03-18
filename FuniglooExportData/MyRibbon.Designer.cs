namespace FuniglooExportData
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Group_ExportData = this.Factory.CreateRibbonGroup();
            this.btnOpenSetting = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.Dropdown_ServerType = this.Factory.CreateRibbonDropDown();
            this.btn_ST_OpenSettings = this.Factory.CreateRibbonButton();
            this.Dropdown_LegionType = this.Factory.CreateRibbonDropDown();
            this.btn_LT_OpenSettings = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnExportData = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.ExportDataInit = this.Factory.CreateRibbonGroup();
            this.Initialize_ExportDataBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Group_ExportData.SuspendLayout();
            this.ExportDataInit.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Group_ExportData);
            this.tab1.Groups.Add(this.ExportDataInit);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // Group_ExportData
            // 
            this.Group_ExportData.Items.Add(this.btnOpenSetting);
            this.Group_ExportData.Items.Add(this.separator2);
            this.Group_ExportData.Items.Add(this.Dropdown_ServerType);
            this.Group_ExportData.Items.Add(this.Dropdown_LegionType);
            this.Group_ExportData.Items.Add(this.separator1);
            this.Group_ExportData.Items.Add(this.btnExportData);
            this.Group_ExportData.Items.Add(this.separator3);
            this.Group_ExportData.Label = "ExportData";
            this.Group_ExportData.Name = "Group_ExportData";
            // 
            // btnOpenSetting
            // 
            this.btnOpenSetting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpenSetting.Image = global::FuniglooExportData.Properties.Resources.Settings;
            this.btnOpenSetting.Label = "Open Setting";
            this.btnOpenSetting.Name = "btnOpenSetting";
            this.btnOpenSetting.ShowImage = true;
            this.btnOpenSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenSettings);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // Dropdown_ServerType
            // 
            this.Dropdown_ServerType.Buttons.Add(this.btn_ST_OpenSettings);
            ribbonDropDownItemImpl1.Label = "None";
            this.Dropdown_ServerType.Items.Add(ribbonDropDownItemImpl1);
            this.Dropdown_ServerType.Label = "ServerType";
            this.Dropdown_ServerType.Name = "Dropdown_ServerType";
            // 
            // btn_ST_OpenSettings
            // 
            this.btn_ST_OpenSettings.Label = "OpenSettings";
            this.btn_ST_OpenSettings.Name = "btn_ST_OpenSettings";
            // 
            // Dropdown_LegionType
            // 
            this.Dropdown_LegionType.Buttons.Add(this.btn_LT_OpenSettings);
            ribbonDropDownItemImpl2.Label = "None";
            this.Dropdown_LegionType.Items.Add(ribbonDropDownItemImpl2);
            this.Dropdown_LegionType.Label = "LegionType";
            this.Dropdown_LegionType.Name = "Dropdown_LegionType";
            // 
            // btn_LT_OpenSettings
            // 
            this.btn_LT_OpenSettings.Label = "OpenSettings";
            this.btn_LT_OpenSettings.Name = "btn_LT_OpenSettings";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnExportData
            // 
            this.btnExportData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportData.Image = global::FuniglooExportData.Properties.Resources.arrow_download;
            this.btnExportData.Label = "Export Table";
            this.btnExportData.Name = "btnExportData";
            this.btnExportData.ShowImage = true;
            this.btnExportData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportData_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // ExportDataInit
            // 
            this.ExportDataInit.Items.Add(this.Initialize_ExportDataBtn);
            this.ExportDataInit.Label = "Initialize ExportData";
            this.ExportDataInit.Name = "ExportDataInit";
            // 
            // Initialize_ExportDataBtn
            // 
            this.Initialize_ExportDataBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Initialize_ExportDataBtn.Image = global::FuniglooExportData.Properties.Resources.Settings;
            this.Initialize_ExportDataBtn.Label = "Initialize ExportData";
            this.Initialize_ExportDataBtn.Name = "Initialize_ExportDataBtn";
            this.Initialize_ExportDataBtn.ShowImage = true;
            this.Initialize_ExportDataBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInit);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Group_ExportData.ResumeLayout(false);
            this.Group_ExportData.PerformLayout();
            this.ExportDataInit.ResumeLayout(false);
            this.ExportDataInit.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_ExportData;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown Dropdown_ServerType;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btn_ST_OpenSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown Dropdown_LegionType;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btn_LT_OpenSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ExportDataInit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Initialize_ExportDataBtn;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
