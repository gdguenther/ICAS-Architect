using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.SharePoint.Client;

namespace ICAS_Architect
{

    public partial class frmDataFlow : System.Windows.Forms.Form
    {
        Visio.DataRecordset databaseRecordset = null;
        //Visio.DataRecordset tableRecordset = null;
        Visio.Application vApplication = null;
        SharepointManager sharepointManager = null;

        ListItemCollection toTableList;
        ListItemCollection fromTableList;

        public frmDataFlow()
        {
            InitializeComponent();
            vApplication = Globals.ThisAddIn.Application;
            sharepointManager = Globals.ThisAddIn.sharepointManager;

            if (sharepointManager.httpDownloadClient is null) {
                sharepointManager.GetSPAccessToken();
            }
            databaseRecordset = sharepointManager.getDatabaseDataRecordset();
            int dbColNumber = sharepointManager.GetColumnNumber(databaseRecordset.GetRowData(0), "Database");
            int dbAppNumber = sharepointManager.GetColumnNumber(databaseRecordset.GetRowData(0), "Application");
            var records = databaseRecordset.GetDataRowIDs("");

            for (int i = 1; i <= records.GetUpperBound(0)+1; i++)
            {
                var record = databaseRecordset.GetRowData(i);
                this.cboFromDB.Items.Add(record.GetValue(dbAppNumber).ToString() + " - " + record.GetValue(dbColNumber).ToString());
                this.cboToDB.Items.Add(record.GetValue(dbAppNumber).ToString() + " - " + record.GetValue(dbColNumber).ToString());
            }
        }


        private void cboFromDB_SelectedValueChanged(object sender, EventArgs e)
        {
            string dBName = cboFromDB.Text.Substring(cboFromDB.Text.IndexOf(" - ", 0) + 3);
            string appName = cboFromDB.Text.Replace(" - " + dBName, "");

            this.lstFromTables.Items.Clear();
            if (dBName == "") return;
            fromTableList = sharepointManager.getTableDatatable(appName, dBName, new string[] { "Database", "Table_Name", "Application" });

            for (int i = 0; i < fromTableList.Count(); i++)
            {
                var item = fromTableList[i];
                string txt = item["Table_Name"].ToString();
                this.lstFromTables.Items.Add(txt);
            }
        }


        public class ListBoxItem
        {
            public string Text { get; set; }
            public string Value { get; set; }
        }

        private void cboToDB_SelectedIndexChanged(object sender, EventArgs e)
        {
            string dBName = cboToDB.Text.Substring(cboToDB.Text.IndexOf(" - ", 0) + 3);
            string appName = cboToDB.Text.Replace(" - " + dBName, "");

            this.lstToTables.Items.Clear();
            if (dBName == "") return;
            toTableList = sharepointManager.getTableDatatable(appName, dBName, new string[] {"Database", "Table_Name", "Application"});

            for (int i = 0; i < toTableList.Count(); i++)
            {
                var item = toTableList[i];
                string txt = item["Table_Name"].ToString();
                this.lstToTables.Items.Add(txt);
            }
        }

        private void lstToTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstFromTables.SelectedIndex >= 0 & lstToTables.SelectedIndex >= 0)
                btnAddFlow.Enabled = true;
            else
                btnAddFlow.Enabled = false;
        }

        private void lstFromTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstFromTables.SelectedIndex >= 0 & lstToTables.SelectedIndex >= 0)
                btnAddFlow.Enabled = true;
            else
                btnAddFlow.Enabled = false;
        }

        private void lstTableFlows_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstTableFlows.SelectedIndex >= 0)
                btnRemoveFlow.Enabled = true;
            else
                btnRemoveFlow.Enabled = false;
        }

        private void btnAddAllFlows_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lstFromTables.Items.Count; i++)
            {
//                var fromItem = fromTableList[i];
                string fromTable = lstFromTables.Items[i].ToString();// fromItem["Table_Name"].ToString();

                for (int j = 0; j < lstToTables.Items.Count; j++)
                {
                    //                    var toItem = toTableList[j];
                    string toTable = lstToTables.Items[i].ToString();

                    if (fromTable.ToLower()== toTable.ToLower())
                    {
                        lstTableFlows.Items.Add(fromTable + " -> " + toTable);
                        lstToTables.Items.RemoveAt(j);
                        lstFromTables.Items.RemoveAt(i);
                        i--;
                        j = toTableList.Count()+1;
                    }
                }
            }
        }

        private void btnAddFlow_Click(object sender, EventArgs e)
        {

        }

        private void btnRemoveFlow_Click(object sender, EventArgs e)
        {

        }

        private void btnRemoveAll_Click(object sender, EventArgs e)
        {

        }
    }
}
