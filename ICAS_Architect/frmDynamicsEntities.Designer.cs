
namespace ICAS_Architect
{
    partial class frmDynamicsEntities
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
            this.btnTmpLoadData = new System.Windows.Forms.Button();
            this.openJsonDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnLoadSQL = new System.Windows.Forms.Button();
            this.btnUploadSQL = new System.Windows.Forms.Button();
            this.btnXMLLoad = new System.Windows.Forms.Button();
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
            this.listView1.Size = new System.Drawing.Size(278, 489);
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
            this.chkAll.Location = new System.Drawing.Point(26, 83);
            this.chkAll.Name = "chkAll";
            this.chkAll.Size = new System.Drawing.Size(70, 17);
            this.chkAll.TabIndex = 1;
            this.chkAll.Text = "Select All";
            this.chkAll.UseVisualStyleBackColor = true;
            // 
            // chkDrawRelations
            // 
            this.chkDrawRelations.AutoSize = true;
            this.chkDrawRelations.Location = new System.Drawing.Point(151, 83);
            this.chkDrawRelations.Name = "chkDrawRelations";
            this.chkDrawRelations.Size = new System.Drawing.Size(98, 17);
            this.chkDrawRelations.TabIndex = 2;
            this.chkDrawRelations.Text = "Draw Relations";
            this.chkDrawRelations.UseVisualStyleBackColor = true;
            // 
            // btnDrawChart
            // 
            this.btnDrawChart.Location = new System.Drawing.Point(188, 42);
            this.btnDrawChart.Name = "btnDrawChart";
            this.btnDrawChart.Size = new System.Drawing.Size(75, 23);
            this.btnDrawChart.TabIndex = 3;
            this.btnDrawChart.Text = "Draw Chart";
            this.btnDrawChart.UseVisualStyleBackColor = true;
            this.btnDrawChart.Click += new System.EventHandler(this.btnDrawChart_Click);
            // 
            // btnTmpLoadData
            // 
            this.btnTmpLoadData.Location = new System.Drawing.Point(12, 13);
            this.btnTmpLoadData.Name = "btnTmpLoadData";
            this.btnTmpLoadData.Size = new System.Drawing.Size(72, 23);
            this.btnTmpLoadData.TabIndex = 4;
            this.btnTmpLoadData.Text = "Load Data";
            this.btnTmpLoadData.UseVisualStyleBackColor = true;
            this.btnTmpLoadData.Click += new System.EventHandler(this.btnTmpLoadData_Click);
            // 
            // openJsonDialog
            // 
            this.openJsonDialog.Filter = "Json files|*.json";
            this.openJsonDialog.Title = "Open metadata visualizer json file";
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(12, 42);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(72, 23);
            this.btnUpload.TabIndex = 5;
            this.btnUpload.Text = "Upload";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // btnLoadSQL
            // 
            this.btnLoadSQL.Location = new System.Drawing.Point(93, 13);
            this.btnLoadSQL.Name = "btnLoadSQL";
            this.btnLoadSQL.Size = new System.Drawing.Size(67, 23);
            this.btnLoadSQL.TabIndex = 6;
            this.btnLoadSQL.Text = "Load SQL";
            this.btnLoadSQL.UseVisualStyleBackColor = true;
            this.btnLoadSQL.Click += new System.EventHandler(this.btnLoadSQL_Click);
            // 
            // btnUploadSQL
            // 
            this.btnUploadSQL.Location = new System.Drawing.Point(93, 42);
            this.btnUploadSQL.Name = "btnUploadSQL";
            this.btnUploadSQL.Size = new System.Drawing.Size(67, 23);
            this.btnUploadSQL.TabIndex = 7;
            this.btnUploadSQL.Text = "Upload SQL";
            this.btnUploadSQL.UseVisualStyleBackColor = true;
            // 
            // btnXMLLoad
            // 
            this.btnXMLLoad.Location = new System.Drawing.Point(188, 12);
            this.btnXMLLoad.Name = "btnXMLLoad";
            this.btnXMLLoad.Size = new System.Drawing.Size(67, 23);
            this.btnXMLLoad.TabIndex = 8;
            this.btnXMLLoad.Text = "Test xml";
            this.btnXMLLoad.UseVisualStyleBackColor = true;
            this.btnXMLLoad.Click += new System.EventHandler(this.btnXMLLoad_Click);
            // 
            // frmDynamicsEntities
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(272, 597);
            this.Controls.Add(this.btnXMLLoad);
            this.Controls.Add(this.btnUploadSQL);
            this.Controls.Add(this.btnLoadSQL);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.btnTmpLoadData);
            this.Controls.Add(this.btnDrawChart);
            this.Controls.Add(this.chkDrawRelations);
            this.Controls.Add(this.chkAll);
            this.Controls.Add(this.listView1);
            this.Name = "frmDynamicsEntities";
            this.Text = "Entities";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.CheckBox chkAll;
        private System.Windows.Forms.CheckBox chkDrawRelations;
        private System.Windows.Forms.Button btnDrawChart;
        private System.Windows.Forms.Button btnTmpLoadData;

        private System.Windows.Forms.OpenFileDialog openJsonDialog;
        private System.Windows.Forms.ColumnHeader Entity;
        private System.Windows.Forms.ColumnHeader Description;
        private System.Windows.Forms.ColumnHeader UID;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnLoadSQL;
        private System.Windows.Forms.Button btnUploadSQL;
        private System.Windows.Forms.Button btnXMLLoad;
    }
}