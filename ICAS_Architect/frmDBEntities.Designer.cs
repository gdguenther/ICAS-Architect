
namespace ICAS_Architect
{
    partial class frmDBEntities
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private void InitializeComponent2()
        {

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listView1 = new System.Windows.Forms.ListView();
            this.Entity = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Description = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.UID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chkAll = new System.Windows.Forms.CheckBox();
            this.chkDrawRelations = new System.Windows.Forms.CheckBox();
            this.btnDrawChart = new System.Windows.Forms.Button();
            this.openJsonDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnLoadSQL = new System.Windows.Forms.Button();
            this.btnUploadSQL = new System.Windows.Forms.Button();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDB = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.CheckBoxes = true;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Entity,
            this.Description,
            this.UID});
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(-3, 104);
            this.listView1.Margin = new System.Windows.Forms.Padding(1);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(304, 369);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // Entity
            // 
            this.Entity.Text = "Entity";
            this.Entity.Width = 120;
            // 
            // Description
            // 
            this.Description.Text = "Description";
            this.Description.Width = 400;
            // 
            // UID
            // 
            this.UID.Text = "UID";
            this.UID.Width = 0;
            // 
            // chkAll
            // 
            this.chkAll.AutoSize = true;
            this.chkAll.Enabled = false;
            this.chkAll.Location = new System.Drawing.Point(12, 67);
            this.chkAll.Name = "chkAll";
            this.chkAll.Size = new System.Drawing.Size(79, 19);
            this.chkAll.TabIndex = 1;
            this.chkAll.Text = "Select All";
            this.chkAll.UseVisualStyleBackColor = true;
            // 
            // chkDrawRelations
            // 
            this.chkDrawRelations.AutoSize = true;
            this.chkDrawRelations.Enabled = false;
            this.chkDrawRelations.Location = new System.Drawing.Point(12, 41);
            this.chkDrawRelations.Name = "chkDrawRelations";
            this.chkDrawRelations.Size = new System.Drawing.Size(105, 19);
            this.chkDrawRelations.TabIndex = 2;
            this.chkDrawRelations.Text = "Add Relations";
            this.chkDrawRelations.UseVisualStyleBackColor = true;
            // 
            // btnDrawChart
            // 
            this.btnDrawChart.Enabled = false;
            this.btnDrawChart.Location = new System.Drawing.Point(12, 12);
            this.btnDrawChart.Name = "btnDrawChart";
            this.btnDrawChart.Size = new System.Drawing.Size(67, 23);
            this.btnDrawChart.TabIndex = 3;
            this.btnDrawChart.Text = "Draw Chart";
            this.btnDrawChart.UseVisualStyleBackColor = true;
            this.btnDrawChart.Click += new System.EventHandler(this.btnDrawChart_Click);
            // 
            // openJsonDialog
            // 
            this.openJsonDialog.Filter = "Json files|*.json";
            this.openJsonDialog.Title = "Open metadata visualizer json file";
            // 
            // btnLoadSQL
            // 
            this.btnLoadSQL.Location = new System.Drawing.Point(146, 67);
            this.btnLoadSQL.Name = "btnLoadSQL";
            this.btnLoadSQL.Size = new System.Drawing.Size(67, 23);
            this.btnLoadSQL.TabIndex = 6;
            this.btnLoadSQL.Text = "Load SQL";
            this.btnLoadSQL.UseVisualStyleBackColor = true;
            this.btnLoadSQL.Click += new System.EventHandler(this.btnLoadSQL_Click);
            // 
            // btnUploadSQL
            // 
            this.btnUploadSQL.Enabled = false;
            this.btnUploadSQL.Location = new System.Drawing.Point(219, 67);
            this.btnUploadSQL.Name = "btnUploadSQL";
            this.btnUploadSQL.Size = new System.Drawing.Size(67, 23);
            this.btnUploadSQL.TabIndex = 7;
            this.btnUploadSQL.Text = "Upload SQL";
            this.btnUploadSQL.UseVisualStyleBackColor = true;
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(172, 10);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(113, 20);
            this.txtServer.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(123, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(42, 15);
            this.label1.TabIndex = 9;
            this.label1.Text = "Server";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(123, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 15);
            this.label2.TabIndex = 11;
            this.label2.Text = "Database";
            // 
            // txtDB
            // 
            this.txtDB.Location = new System.Drawing.Point(172, 36);
            this.txtDB.Name = "txtDB";
            this.txtDB.Size = new System.Drawing.Size(113, 20);
            this.txtDB.TabIndex = 10;
            // 
            // frmDBEntities
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(298, 477);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtDB);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtServer);
            this.Controls.Add(this.btnUploadSQL);
            this.Controls.Add(this.btnLoadSQL);
            this.Controls.Add(this.btnDrawChart);
            this.Controls.Add(this.chkDrawRelations);
            this.Controls.Add(this.chkAll);
            this.Controls.Add(this.listView1);
            this.Name = "frmDBEntities";
            this.Text = "Import DB";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.CheckBox chkAll;
        private System.Windows.Forms.CheckBox chkDrawRelations;
        private System.Windows.Forms.Button btnDrawChart;

        private System.Windows.Forms.OpenFileDialog openJsonDialog;
        private System.Windows.Forms.ColumnHeader Entity;
        private System.Windows.Forms.ColumnHeader Description;
        private System.Windows.Forms.ColumnHeader UID;
        private System.Windows.Forms.Button btnLoadSQL;
        private System.Windows.Forms.Button btnUploadSQL;
        private System.Windows.Forms.TextBox txtServer;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtDB;
    }
}