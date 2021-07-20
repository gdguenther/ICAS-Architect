
namespace ICAS_Architect
{
    partial class ICASArchitect : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ICASArchitect()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Tab1 = this.Factory.CreateRibbonTab();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnLinkToRepository = this.Factory.CreateRibbonButton();
            this.btnUploadToSP = this.Factory.CreateRibbonButton();
            this.btnImportFromRepository = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.cboDB = this.Factory.CreateRibbonComboBox();
            this.cboApp = this.Factory.CreateRibbonComboBox();
            this.chkShowViews = this.Factory.CreateRibbonCheckBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnImportJSON = this.Factory.CreateRibbonButton();
            this.btnImportDBm = this.Factory.CreateRibbonButton();
            this.btnDownloadDynamics = this.Factory.CreateRibbonButton();
            this.grpUtils = this.Factory.CreateRibbonGroup();
            this.Administration = this.Factory.CreateRibbonMenu();
            this.btnShowEntities = this.Factory.CreateRibbonButton();
            this.btnCopySPToExcel = this.Factory.CreateRibbonButton();
            this.btnCopyExceltoSP = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.chkDrawRelatedEntities = this.Factory.CreateRibbonCheckBox();
            this.chkShowAttributes = this.Factory.CreateRibbonCheckBox();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btnDefineDataFlow = this.Factory.CreateRibbonButton();
            this.Tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.grpUtils.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab1
            // 
            this.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab1.ControlId.OfficeId = "TabHome";
            this.Tab1.Label = "Home";
            this.Tab1.Name = "Tab1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Groups.Add(this.grpUtils);
            this.tab2.Groups.Add(this.group4);
            this.tab2.Groups.Add(this.group5);
            this.tab2.Label = "ICAS Architect";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnLinkToRepository);
            this.group2.Items.Add(this.btnUploadToSP);
            this.group2.Items.Add(this.btnImportFromRepository);
            this.group2.Label = "ICAS_Architect";
            this.group2.Name = "group2";
            // 
            // btnLinkToRepository
            // 
            this.btnLinkToRepository.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLinkToRepository.Image = global::ICAS_Architect.Properties.Resources.Repository;
            this.btnLinkToRepository.Label = "Link to Repository";
            this.btnLinkToRepository.Name = "btnLinkToRepository";
            this.btnLinkToRepository.ShowImage = true;
            this.btnLinkToRepository.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLinkToRepository_Click);
            // 
            // btnUploadToSP
            // 
            this.btnUploadToSP.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUploadToSP.Image = global::ICAS_Architect.Properties.Resources.Save_Image;
            this.btnUploadToSP.Label = "Save to Repository";
            this.btnUploadToSP.Name = "btnUploadToSP";
            this.btnUploadToSP.ShowImage = true;
            this.btnUploadToSP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUploadToSP_Click);
            // 
            // btnImportFromRepository
            // 
            this.btnImportFromRepository.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportFromRepository.Enabled = false;
            this.btnImportFromRepository.Image = global::ICAS_Architect.Properties.Resources.sharepoint;
            this.btnImportFromRepository.Label = "Import Repository";
            this.btnImportFromRepository.Name = "btnImportFromRepository";
            this.btnImportFromRepository.ShowImage = true;
            this.btnImportFromRepository.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportFromRepository_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.cboDB);
            this.group1.Items.Add(this.cboApp);
            this.group1.Items.Add(this.chkShowViews);
            this.group1.Label = "Filters";
            this.group1.Name = "group1";
            // 
            // cboDB
            // 
            this.cboDB.Enabled = false;
            this.cboDB.Label = "Database";
            this.cboDB.Name = "cboDB";
            this.cboDB.Text = null;
            this.cboDB.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cboDB_TextChanged);
            // 
            // cboApp
            // 
            this.cboApp.Enabled = false;
            this.cboApp.Label = "Application";
            this.cboApp.Name = "cboApp";
            this.cboApp.Text = null;
            // 
            // chkShowViews
            // 
            this.chkShowViews.Enabled = false;
            this.chkShowViews.Label = "Show Views";
            this.chkShowViews.Name = "chkShowViews";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnImportJSON);
            this.group3.Items.Add(this.btnImportDBm);
            this.group3.Items.Add(this.btnDownloadDynamics);
            this.group3.Label = "Import MetaData";
            this.group3.Name = "group3";
            // 
            // btnImportJSON
            // 
            this.btnImportJSON.Image = global::ICAS_Architect.Properties.Resources.JSON_Image;
            this.btnImportJSON.Label = "Import JSON";
            this.btnImportJSON.Name = "btnImportJSON";
            this.btnImportJSON.ShowImage = true;
            this.btnImportJSON.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportJSON_Click);
            // 
            // btnImportDBm
            // 
            this.btnImportDBm.Image = global::ICAS_Architect.Properties.Resources.DB_Image;
            this.btnImportDBm.Label = "SQL DB ";
            this.btnImportDBm.Name = "btnImportDBm";
            this.btnImportDBm.ShowImage = true;
            this.btnImportDBm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportDBm_Click);
            // 
            // btnDownloadDynamics
            // 
            this.btnDownloadDynamics.Image = global::ICAS_Architect.Properties.Resources.download;
            this.btnDownloadDynamics.Label = "Import Dx API";
            this.btnDownloadDynamics.Name = "btnDownloadDynamics";
            this.btnDownloadDynamics.ShowImage = true;
            this.btnDownloadDynamics.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDownloadDynamics_Click);
            // 
            // grpUtils
            // 
            this.grpUtils.Items.Add(this.Administration);
            this.grpUtils.Label = "Utilities";
            this.grpUtils.Name = "grpUtils";
            // 
            // Administration
            // 
            this.Administration.Items.Add(this.btnShowEntities);
            this.Administration.Items.Add(this.btnCopySPToExcel);
            this.Administration.Items.Add(this.btnCopyExceltoSP);
            this.Administration.Label = "Administration";
            this.Administration.Name = "Administration";
            // 
            // btnShowEntities
            // 
            this.btnShowEntities.Label = "Set up Sharepoint";
            this.btnShowEntities.Name = "btnShowEntities";
            this.btnShowEntities.ShowImage = true;
            this.btnShowEntities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowEntities_Click);
            // 
            // btnCopySPToExcel
            // 
            this.btnCopySPToExcel.Label = "Copy SP to Excel";
            this.btnCopySPToExcel.Name = "btnCopySPToExcel";
            this.btnCopySPToExcel.ShowImage = true;
            this.btnCopySPToExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopySPToExcel_Click);
            // 
            // btnCopyExceltoSP
            // 
            this.btnCopyExceltoSP.Label = "Upload XL to SP";
            this.btnCopyExceltoSP.Name = "btnCopyExceltoSP";
            this.btnCopyExceltoSP.ShowImage = true;
            this.btnCopyExceltoSP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyExceltoSP_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.chkDrawRelatedEntities);
            this.group4.Items.Add(this.chkShowAttributes);
            this.group4.Label = "Drawing";
            this.group4.Name = "group4";
            // 
            // chkDrawRelatedEntities
            // 
            this.chkDrawRelatedEntities.Checked = true;
            this.chkDrawRelatedEntities.Label = "Import Related Entities";
            this.chkDrawRelatedEntities.Name = "chkDrawRelatedEntities";
            // 
            // chkShowAttributes
            // 
            this.chkShowAttributes.Label = "Show Relation Columns";
            this.chkShowAttributes.Name = "chkShowAttributes";
            this.chkShowAttributes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkShowAttributes_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.btnDefineDataFlow);
            this.group5.Label = "Data Flows";
            this.group5.Name = "group5";
            // 
            // btnDefineDataFlow
            // 
            this.btnDefineDataFlow.Label = "Define Data Flow";
            this.btnDefineDataFlow.Name = "btnDefineDataFlow";
            this.btnDefineDataFlow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDefineDataFlow_Click);
            // 
            // ICASArchitect
            // 
            this.Name = "ICASArchitect";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.Tab1);
            this.Tabs.Add(this.tab2);
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.grpUtils.ResumeLayout(false);
            this.grpUtils.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLinkToRepository;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowEntities;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDownloadDynamics;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUploadToSP;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpUtils;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu Administration;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportJSON;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportDBm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportFromRepository;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkDrawRelatedEntities;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkShowViews;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cboApp;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cboDB;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkShowAttributes;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDefineDataFlow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopySPToExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyExceltoSP;
    }

    partial class ThisRibbonCollection
    {
        internal ICASArchitect Ribbon
        {
            get { return this.GetRibbon<ICASArchitect>(); }
        }
    }
}
