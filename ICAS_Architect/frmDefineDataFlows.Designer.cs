
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
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAddAllFlows = new System.Windows.Forms.Button();
            this.btnAddFlow = new System.Windows.Forms.Button();
            this.btnRemoveFlow = new System.Windows.Forms.Button();
            this.lstFromTables = new System.Windows.Forms.ListBox();
            this.lstToTables = new System.Windows.Forms.ListBox();
            this.lstTableFlows = new System.Windows.Forms.ListBox();
            this.btnRemoveAll = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cboFromDB
            // 
            this.cboFromDB.FormattingEnabled = true;
            this.cboFromDB.Location = new System.Drawing.Point(123, 42);
            this.cboFromDB.MaxDropDownItems = 100;
            this.cboFromDB.Name = "cboFromDB";
            this.cboFromDB.Size = new System.Drawing.Size(280, 21);
            this.cboFromDB.Sorted = true;
            this.cboFromDB.TabIndex = 0;
            this.cboFromDB.SelectedValueChanged += new System.EventHandler(this.cboFromDB_SelectedValueChanged);
            // 
            // cboToDB
            // 
            this.cboToDB.FormattingEnabled = true;
            this.cboToDB.IntegralHeight = false;
            this.cboToDB.Location = new System.Drawing.Point(425, 42);
            this.cboToDB.MaxDropDownItems = 100;
            this.cboToDB.Name = "cboToDB";
            this.cboToDB.Size = new System.Drawing.Size(279, 21);
            this.cboToDB.Sorted = true;
            this.cboToDB.TabIndex = 1;
            this.cboToDB.SelectedIndexChanged += new System.EventHandler(this.cboToDB_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Data Store";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 88);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Entities/Tables";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(122, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Data Source";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(422, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Destination";
            // 
            // btnAddAllFlows
            // 
            this.btnAddAllFlows.Location = new System.Drawing.Point(717, 164);
            this.btnAddAllFlows.Name = "btnAddAllFlows";
            this.btnAddAllFlows.Size = new System.Drawing.Size(29, 23);
            this.btnAddAllFlows.TabIndex = 9;
            this.btnAddAllFlows.Text = ">>";
            this.btnAddAllFlows.UseVisualStyleBackColor = true;
            this.btnAddAllFlows.Click += new System.EventHandler(this.btnAddAllFlows_Click);
            // 
            // btnAddFlow
            // 
            this.btnAddFlow.Location = new System.Drawing.Point(717, 195);
            this.btnAddFlow.Name = "btnAddFlow";
            this.btnAddFlow.Size = new System.Drawing.Size(29, 23);
            this.btnAddFlow.TabIndex = 11;
            this.btnAddFlow.Text = ">";
            this.btnAddFlow.UseVisualStyleBackColor = true;
            this.btnAddFlow.Click += new System.EventHandler(this.btnAddFlow_Click);
            // 
            // btnRemoveFlow
            // 
            this.btnRemoveFlow.Location = new System.Drawing.Point(717, 227);
            this.btnRemoveFlow.Name = "btnRemoveFlow";
            this.btnRemoveFlow.Size = new System.Drawing.Size(29, 23);
            this.btnRemoveFlow.TabIndex = 12;
            this.btnRemoveFlow.Text = "<";
            this.btnRemoveFlow.UseVisualStyleBackColor = true;
            this.btnRemoveFlow.Click += new System.EventHandler(this.btnRemoveFlow_Click);
            // 
            // lstFromTables
            // 
            this.lstFromTables.FormattingEnabled = true;
            this.lstFromTables.Location = new System.Drawing.Point(125, 88);
            this.lstFromTables.Name = "lstFromTables";
            this.lstFromTables.Size = new System.Drawing.Size(278, 407);
            this.lstFromTables.Sorted = true;
            this.lstFromTables.TabIndex = 13;
            this.lstFromTables.SelectedIndexChanged += new System.EventHandler(this.lstFromTables_SelectedIndexChanged);
            // 
            // lstToTables
            // 
            this.lstToTables.FormattingEnabled = true;
            this.lstToTables.Location = new System.Drawing.Point(425, 88);
            this.lstToTables.Name = "lstToTables";
            this.lstToTables.Size = new System.Drawing.Size(278, 407);
            this.lstToTables.Sorted = true;
            this.lstToTables.TabIndex = 14;
            this.lstToTables.SelectedIndexChanged += new System.EventHandler(this.lstToTables_SelectedIndexChanged);
            // 
            // lstTableFlows
            // 
            this.lstTableFlows.FormattingEnabled = true;
            this.lstTableFlows.Location = new System.Drawing.Point(761, 88);
            this.lstTableFlows.Name = "lstTableFlows";
            this.lstTableFlows.Size = new System.Drawing.Size(278, 407);
            this.lstTableFlows.Sorted = true;
            this.lstTableFlows.TabIndex = 15;
            this.lstTableFlows.SelectedIndexChanged += new System.EventHandler(this.lstTableFlows_SelectedIndexChanged);
            // 
            // btnRemoveAll
            // 
            this.btnRemoveAll.Location = new System.Drawing.Point(717, 256);
            this.btnRemoveAll.Name = "btnRemoveAll";
            this.btnRemoveAll.Size = new System.Drawing.Size(29, 23);
            this.btnRemoveAll.TabIndex = 16;
            this.btnRemoveAll.Text = "<<";
            this.btnRemoveAll.UseVisualStyleBackColor = true;
            this.btnRemoveAll.Click += new System.EventHandler(this.btnRemoveAll_Click);
            // 
            // frmDataFlow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1400, 847);
            this.Controls.Add(this.btnRemoveAll);
            this.Controls.Add(this.lstTableFlows);
            this.Controls.Add(this.lstToTables);
            this.Controls.Add(this.lstFromTables);
            this.Controls.Add(this.btnRemoveFlow);
            this.Controls.Add(this.btnAddFlow);
            this.Controls.Add(this.btnAddAllFlows);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboToDB);
            this.Controls.Add(this.cboFromDB);
            this.Name = "frmDataFlow";
            this.Text = "Define Data Flows";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cboFromDB;
        private System.Windows.Forms.ComboBox cboToDB;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnAddAllFlows;
        private System.Windows.Forms.Button btnAddFlow;
        private System.Windows.Forms.Button btnRemoveFlow;
        private System.Windows.Forms.ListBox lstFromTables;
        private System.Windows.Forms.ListBox lstToTables;
        private System.Windows.Forms.ListBox lstTableFlows;
        private System.Windows.Forms.Button btnRemoveAll;
    }
}