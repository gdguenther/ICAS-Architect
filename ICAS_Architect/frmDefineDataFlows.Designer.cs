
namespace ICAS_Architect
{
    partial class frmDataFlow
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
            this.cboFromDB = new System.Windows.Forms.ComboBox();
            this.cboToDB = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnDocumentFlow = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.dgTables = new System.Windows.Forms.DataGridView();
            this.dgColumns = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.lblDestinationColumns = new System.Windows.Forms.Label();
            this.btnAutoMatchTables = new System.Windows.Forms.Button();
            this.btnCreateAllTables = new System.Windows.Forms.Button();
            this.btnAutoMatchColumns = new System.Windows.Forms.Button();
            this.btnCreateColumns = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgTables)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgColumns)).BeginInit();
            this.SuspendLayout();
            // 
            // cboFromDB
            // 
            this.cboFromDB.FormattingEnabled = true;
            this.cboFromDB.Location = new System.Drawing.Point(15, 30);
            this.cboFromDB.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboFromDB.MaxDropDownItems = 100;
            this.cboFromDB.Name = "cboFromDB";
            this.cboFromDB.Size = new System.Drawing.Size(372, 24);
            this.cboFromDB.Sorted = true;
            this.cboFromDB.TabIndex = 0;
            this.cboFromDB.SelectedValueChanged += new System.EventHandler(this.cboFromDB_SelectedValueChanged);
            // 
            // cboToDB
            // 
            this.cboToDB.FormattingEnabled = true;
            this.cboToDB.IntegralHeight = false;
            this.cboToDB.Location = new System.Drawing.Point(417, 30);
            this.cboToDB.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboToDB.MaxDropDownItems = 100;
            this.cboToDB.Name = "cboToDB";
            this.cboToDB.Size = new System.Drawing.Size(371, 24);
            this.cboToDB.Sorted = true;
            this.cboToDB.TabIndex = 1;
            this.cboToDB.SelectedIndexChanged += new System.EventHandler(this.cboToDB_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 10);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(87, 17);
            this.label2.TabIndex = 7;
            this.label2.Text = "Data Source";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(413, 10);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(79, 17);
            this.label4.TabIndex = 8;
            this.label4.Text = "Destination";
            // 
            // btnDocumentFlow
            // 
            this.btnDocumentFlow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDocumentFlow.Location = new System.Drawing.Point(1204, 634);
            this.btnDocumentFlow.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnDocumentFlow.Name = "btnDocumentFlow";
            this.btnDocumentFlow.Size = new System.Drawing.Size(104, 42);
            this.btnDocumentFlow.TabIndex = 19;
            this.btnDocumentFlow.Text = "Document Flow";
            this.btnDocumentFlow.UseVisualStyleBackColor = true;
            this.btnDocumentFlow.Click += new System.EventHandler(this.btnDocumentFlow_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(1204, 6);
            this.btnClose.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(104, 42);
            this.btnClose.TabIndex = 20;
            this.btnClose.Text = "&Cancel";
            this.btnClose.UseVisualStyleBackColor = true;
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Location = new System.Drawing.Point(1092, 6);
            this.btnSave.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(104, 42);
            this.btnSave.TabIndex = 21;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // dgTables
            // 
            this.dgTables.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dgTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgTables.Location = new System.Drawing.Point(16, 100);
            this.dgTables.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dgTables.Name = "dgTables";
            this.dgTables.RowHeadersWidth = 51;
            this.dgTables.Size = new System.Drawing.Size(623, 527);
            this.dgTables.TabIndex = 22;
            this.dgTables.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgTables_CellEnter);
            this.dgTables.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgTables_CellValueChanged);
            this.dgTables.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgTables_CurrentCellDirtyStateChanged);
            this.dgTables.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgTables_RowHeaderMouseClick);
            // 
            // dgColumns
            // 
            this.dgColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgColumns.Location = new System.Drawing.Point(668, 100);
            this.dgColumns.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dgColumns.Name = "dgColumns";
            this.dgColumns.RowHeadersWidth = 51;
            this.dgColumns.Size = new System.Drawing.Size(648, 527);
            this.dgColumns.TabIndex = 23;
            this.dgColumns.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgColumns_CellEnter);
            this.dgColumns.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgColumns_CellValueChanged);
            this.dgColumns.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgColumns_CurrentCellDirtyStateChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 79);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 17);
            this.label1.TabIndex = 24;
            this.label1.Text = "Source Tables";
            // 
            // lblDestinationColumns
            // 
            this.lblDestinationColumns.AutoSize = true;
            this.lblDestinationColumns.Location = new System.Drawing.Point(673, 76);
            this.lblDestinationColumns.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDestinationColumns.Name = "lblDestinationColumns";
            this.lblDestinationColumns.Size = new System.Drawing.Size(257, 17);
            this.lblDestinationColumns.TabIndex = 25;
            this.lblDestinationColumns.Text = "Destination Columns for Table Selected";
            // 
            // btnAutoMatchTables
            // 
            this.btnAutoMatchTables.Location = new System.Drawing.Point(505, 76);
            this.btnAutoMatchTables.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAutoMatchTables.Name = "btnAutoMatchTables";
            this.btnAutoMatchTables.Size = new System.Drawing.Size(133, 23);
            this.btnAutoMatchTables.TabIndex = 26;
            this.btnAutoMatchTables.Text = "Auto Match";
            this.btnAutoMatchTables.UseVisualStyleBackColor = true;
            this.btnAutoMatchTables.Click += new System.EventHandler(this.btnAutoMatchTables_Click);
            // 
            // btnCreateAllTables
            // 
            this.btnCreateAllTables.Location = new System.Drawing.Point(371, 76);
            this.btnCreateAllTables.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCreateAllTables.Name = "btnCreateAllTables";
            this.btnCreateAllTables.Size = new System.Drawing.Size(133, 23);
            this.btnCreateAllTables.TabIndex = 27;
            this.btnCreateAllTables.Text = "Create All Tables";
            this.btnCreateAllTables.UseVisualStyleBackColor = true;
            // 
            // btnAutoMatchColumns
            // 
            this.btnAutoMatchColumns.Location = new System.Drawing.Point(1183, 75);
            this.btnAutoMatchColumns.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAutoMatchColumns.Name = "btnAutoMatchColumns";
            this.btnAutoMatchColumns.Size = new System.Drawing.Size(133, 23);
            this.btnAutoMatchColumns.TabIndex = 28;
            this.btnAutoMatchColumns.Text = "Auto Match";
            this.btnAutoMatchColumns.UseVisualStyleBackColor = true;
            this.btnAutoMatchColumns.Click += new System.EventHandler(this.btnAutoMatchColumns_Click);
            // 
            // btnCreateColumns
            // 
            this.btnCreateColumns.Location = new System.Drawing.Point(1047, 75);
            this.btnCreateColumns.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCreateColumns.Name = "btnCreateColumns";
            this.btnCreateColumns.Size = new System.Drawing.Size(133, 23);
            this.btnCreateColumns.TabIndex = 29;
            this.btnCreateColumns.Text = "Create All Columns";
            this.btnCreateColumns.UseVisualStyleBackColor = true;
            // 
            // frmDataFlow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1324, 678);
            this.Controls.Add(this.btnCreateColumns);
            this.Controls.Add(this.btnAutoMatchColumns);
            this.Controls.Add(this.btnCreateAllTables);
            this.Controls.Add(this.btnAutoMatchTables);
            this.Controls.Add(this.dgColumns);
            this.Controls.Add(this.dgTables);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnDocumentFlow);
            this.Controls.Add(this.cboToDB);
            this.Controls.Add(this.cboFromDB);
            this.Controls.Add(this.lblDestinationColumns);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "frmDataFlow";
            this.Text = "Define Data Flows";
            ((System.ComponentModel.ISupportInitialize)(this.dgTables)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgColumns)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cboFromDB;
        private System.Windows.Forms.ComboBox cboToDB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnDocumentFlow;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.DataGridView dgTables;
        private System.Windows.Forms.DataGridView dgColumns;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblDestinationColumns;
        private System.Windows.Forms.Button btnAutoMatchTables;
        private System.Windows.Forms.Button btnCreateAllTables;
        private System.Windows.Forms.Button btnAutoMatchColumns;
        private System.Windows.Forms.Button btnCreateColumns;
    }
}