
namespace ICAS_Architect
{
    partial class frmDataArchitect
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabFilters = new System.Windows.Forms.TabPage();
            this.chkIncludeAPIs = new System.Windows.Forms.CheckBox();
            this.chkIncludeViews = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cboDatabase = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboApplications = new System.Windows.Forms.ComboBox();
            this.tabDrawing = new System.Windows.Forms.TabPage();
            this.chkChangedFlag = new System.Windows.Forms.CheckBox();
            this.chkDrawRelatedEntities = new System.Windows.Forms.CheckBox();
            this.tabControl1.SuspendLayout();
            this.tabFilters.SuspendLayout();
            this.tabDrawing.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabFilters);
            this.tabControl1.Controls.Add(this.tabDrawing);
            this.tabControl1.Location = new System.Drawing.Point(-4, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(283, 309);
            this.tabControl1.TabIndex = 5;
            // 
            // tabFilters
            // 
            this.tabFilters.Controls.Add(this.chkIncludeAPIs);
            this.tabFilters.Controls.Add(this.chkIncludeViews);
            this.tabFilters.Controls.Add(this.label2);
            this.tabFilters.Controls.Add(this.cboDatabase);
            this.tabFilters.Controls.Add(this.label1);
            this.tabFilters.Controls.Add(this.cboApplications);
            this.tabFilters.Location = new System.Drawing.Point(4, 22);
            this.tabFilters.Name = "tabFilters";
            this.tabFilters.Padding = new System.Windows.Forms.Padding(3);
            this.tabFilters.Size = new System.Drawing.Size(275, 283);
            this.tabFilters.TabIndex = 0;
            this.tabFilters.Text = "Filters";
            this.tabFilters.UseVisualStyleBackColor = true;
            // 
            // chkIncludeAPIs
            // 
            this.chkIncludeAPIs.AutoSize = true;
            this.chkIncludeAPIs.Checked = true;
            this.chkIncludeAPIs.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIncludeAPIs.Location = new System.Drawing.Point(65, 79);
            this.chkIncludeAPIs.Name = "chkIncludeAPIs";
            this.chkIncludeAPIs.Size = new System.Drawing.Size(86, 17);
            this.chkIncludeAPIs.TabIndex = 10;
            this.chkIncludeAPIs.Text = "Include APIs";
            this.chkIncludeAPIs.UseVisualStyleBackColor = true;
            // 
            // chkIncludeViews
            // 
            this.chkIncludeViews.AutoSize = true;
            this.chkIncludeViews.Checked = true;
            this.chkIncludeViews.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIncludeViews.Location = new System.Drawing.Point(64, 59);
            this.chkIncludeViews.Name = "chkIncludeViews";
            this.chkIncludeViews.Size = new System.Drawing.Size(92, 17);
            this.chkIncludeViews.TabIndex = 9;
            this.chkIncludeViews.Text = "Include Views";
            this.chkIncludeViews.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Database";
            // 
            // cboDatabase
            // 
            this.cboDatabase.FormattingEnabled = true;
            this.cboDatabase.Location = new System.Drawing.Point(64, 34);
            this.cboDatabase.Name = "cboDatabase";
            this.cboDatabase.Size = new System.Drawing.Size(121, 21);
            this.cboDatabase.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(1, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Application";
            // 
            // cboApplications
            // 
            this.cboApplications.FormattingEnabled = true;
            this.cboApplications.Location = new System.Drawing.Point(64, 10);
            this.cboApplications.Name = "cboApplications";
            this.cboApplications.Size = new System.Drawing.Size(121, 21);
            this.cboApplications.TabIndex = 5;
            // 
            // tabDrawing
            // 
            this.tabDrawing.Controls.Add(this.chkChangedFlag);
            this.tabDrawing.Controls.Add(this.chkDrawRelatedEntities);
            this.tabDrawing.Location = new System.Drawing.Point(4, 22);
            this.tabDrawing.Name = "tabDrawing";
            this.tabDrawing.Padding = new System.Windows.Forms.Padding(3);
            this.tabDrawing.Size = new System.Drawing.Size(275, 283);
            this.tabDrawing.TabIndex = 1;
            this.tabDrawing.Text = "Drawing";
            this.tabDrawing.UseVisualStyleBackColor = true;
            // 
            // chkChangedFlag
            // 
            this.chkChangedFlag.AutoSize = true;
            this.chkChangedFlag.Location = new System.Drawing.Point(24, 69);
            this.chkChangedFlag.Name = "chkChangedFlag";
            this.chkChangedFlag.Size = new System.Drawing.Size(122, 17);
            this.chkChangedFlag.TabIndex = 1;
            this.chkChangedFlag.Text = "Show Changed Flag";
            this.chkChangedFlag.UseVisualStyleBackColor = true;
            // 
            // chkDrawRelatedEntities
            // 
            this.chkDrawRelatedEntities.AutoSize = true;
            this.chkDrawRelatedEntities.Location = new System.Drawing.Point(24, 43);
            this.chkDrawRelatedEntities.Name = "chkDrawRelatedEntities";
            this.chkDrawRelatedEntities.Size = new System.Drawing.Size(128, 17);
            this.chkDrawRelatedEntities.TabIndex = 0;
            this.chkDrawRelatedEntities.Text = "Draw Related Entities";
            this.chkDrawRelatedEntities.UseVisualStyleBackColor = true;
            // 
            // frmDataArchitect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(275, 311);
            this.Controls.Add(this.tabControl1);
            this.Name = "frmDataArchitect";
            this.Text = "Data Architect";
            this.tabControl1.ResumeLayout(false);
            this.tabFilters.ResumeLayout(false);
            this.tabFilters.PerformLayout();
            this.tabDrawing.ResumeLayout(false);
            this.tabDrawing.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabFilters;
        private System.Windows.Forms.CheckBox chkIncludeAPIs;
        private System.Windows.Forms.CheckBox chkIncludeViews;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboDatabase;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboApplications;
        private System.Windows.Forms.TabPage tabDrawing;
        private System.Windows.Forms.CheckBox chkChangedFlag;
        private System.Windows.Forms.CheckBox chkDrawRelatedEntities;
    }
}