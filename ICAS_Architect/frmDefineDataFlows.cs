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
//using Microsoft.Office.Interop.Excel;
using XL = Microsoft.Office.Interop.Excel;

namespace ICAS_Architect
{

    public partial class frmDataFlow : System.Windows.Forms.Form
    {
        private delegate void SafeCallDelegate();
        private delegate void SafeCallDelegateLong(long x);
        Visio.Application vApplication = null;
        SharepointManager sharepointManager = null;


        DataTable dBDataTable = null;
        DataTable sourceTableDataTable = null;
        DataTable destTableDataTable = null;

        DataTable sourceColumnDataTable = null;
        DataTable destColumnDataTable = null;
        bool bFormLoadComplete = false;
        //        DataTable mapTableDataTable = null;
        //        DataTable mapColumnDataTable = null;


        public frmDataFlow()
        {
            InitializeComponent();
            vApplication = Globals.ThisAddIn.Application;
            sharepointManager = Globals.ThisAddIn.sharepointManager;

            sharepointManager.GetClientContext(true, false, true);

            GetDBList();
        }

        private void SafeTableDGComboboxUpdater()
        {
            if (dgTables.InvokeRequired)
            {
                var d = new SafeCallDelegate(SafeTableDGComboboxUpdater);
                dgTables.Invoke(d, new object[] { });
            }
            else
            {
                if (destTableDataTable == null) return; // don't bother if we have no tables in the destination table
                DataGridViewComboBoxColumn column = (DataGridViewComboBoxColumn)dgTables.Columns["Destination Table"];
                if (column is null) return;

                List<ComboListInfo> toList = new List<ComboListInfo>();
                foreach (DataRow rec in destTableDataTable.Rows)
                    toList.Add(new ComboListInfo((long)rec["Id"], rec["Table_Name"].ToString(), "", (long)rec["Database_NameId"]));

                toList.Add(new ComboListInfo(0, "<Add New>", "", 0));
                toList.Add(new ComboListInfo(-1, "", "", 0));
                toList.Sort((x, y) => x.Name.CompareTo(y.Name));

                column.DataSource = toList;
                column.DisplayMember = "Name";
                column.ValueMember = "ID";
            }
        }

        private void SafeTableDGUpdate()
        {
            if (dgTables.InvokeRequired)
            {
                var d = new SafeCallDelegate(SafeTableDGUpdate);
                dgTables.Invoke(d, new object[] { });
            }
            else
            {
                dgTables.Columns.Clear();
                if (sourceTableDataTable == null) return;
                dgTables.DataSource = sourceTableDataTable;

                if (dgTables.Columns["Table_Name"].Visible)
                {
                    DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn();
                    column.Name = "Destination Table";
                    column.DisplayIndex = 0;
                    column.Width = 200;
                    column.DataPropertyName = "Destination_Table";
                    var idx = dgTables.Columns.Add(column);
                    //dgTables.Columns["ID"].Visible = false;
                    //dgTables.Columns["Title"].Visible = false;
                    //dgTables.Columns["Database_Name"].Visible = false;
                    //dgTables.Columns["Application_Name"].Visible = false;
                    //dgTables.Columns["Application_NameId"].Visible = false;
                    //dgTables.Columns["Edit_Status"].Visible = false;
                    dgTables.Columns["Table_Name"].Width = 200;
                    dgTables.Columns["Description"].Width = 300;
                    dgTables.Columns["Table_Name"].DisplayIndex = 0;
                    dgTables.Columns["Destination Table"].DisplayIndex = 1;
                    dgTables.Columns["Description"].DisplayIndex = 2;
                    dgTables.Sort(dgTables.Columns["Table_Name"], ListSortDirection.Ascending);
                }
                SafeTableDGComboboxUpdater();
            }
        }

        private void ColumnDGComboboxUpdater()
        {
            if (sourceColumnDataTable == null) return;
            DataGridViewComboBoxColumn column = (DataGridViewComboBoxColumn)dgColumns.Columns["Source Column"];
            if (column is null) return;
            if (sourceColumnDataTable is null) return;

            List<ComboListInfo> toList = new List<ComboListInfo>();
            foreach (DataRow rec in sourceColumnDataTable.Rows)
                if ((long)dgTables["Id", dgTables.CurrentCell.RowIndex].Value == (long)rec["Table_NameId"])
                    toList.Add(new ComboListInfo((long)rec["Id"], rec["Column_Name"].ToString(), rec["Data_Type"].ToString(), (long)rec["Table_NameId"]));

            toList.Sort((x, y) => x.Name.CompareTo(y.Name));
            column.DataSource = toList;
            column.DisplayMember = "Name";
            column.ValueMember = "ID";
        }

        private void SafeColumnDGComboboxUpdater()
        {
            if (dgColumns.InvokeRequired)
            {
                var d = new SafeCallDelegate(SafeColumnDGComboboxUpdater);
                dgTables.Invoke(d, new object[] { });
            }
            else
            {
                ColumnDGComboboxUpdater();
            }
        }

        private void ColumnDGUpdater(long TableID)
        {
            if (destColumnDataTable == null) return;
            if (dgColumns.DataSource is null)
            {
                // Add columns to the underlying datasource
                if(!dtDataFieldExists(destColumnDataTable, "Column_Edit_Status")) destColumnDataTable.Columns.Add("Column_Edit_Status");
                if (!dtDataFieldExists(destColumnDataTable, "Relation_Edit_Status")) destColumnDataTable.Columns.Add("Relation_Edit_Status");
                if (!dtDataFieldExists(destColumnDataTable, "Source_Column")) destColumnDataTable.Columns.Add("Source_Column", typeof(long));
                if (!dtDataFieldExists(destColumnDataTable, "Source_TableId")) destColumnDataTable.Columns.Add("Source_TableId", typeof(long));

                dgColumns.DataSource = destColumnDataTable;
                DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn();
                column.Name = "Source Column";
                column.DataPropertyName = "Source_Column";
                column.Width = 200;
                var idx = dgColumns.Columns.Add(column);

                //dgColumns.Columns["Title"].Visible = false;
                //dgColumns.Columns["ID"].Visible = false;
                //dgColumns.Columns["Table_NameID"].Visible = false;
                //dgColumns.Columns["Edit_Status"].Visible = false;
                dgColumns.Columns["Description"].Width = 200;
                dgColumns.Columns["Data_Type"].Width = 100;
                dgColumns.Columns["Column_Name"].Width = 200;
                dgColumns.Columns["Column_Name"].DisplayIndex = 0;
                dgColumns.Columns["Description"].DisplayIndex = 2;
                dgColumns.Columns["Source Column"].DisplayIndex = 1;


                lblDestinationColumns.Text = "Destination Columns for ... " + "";
            }
            destColumnDataTable.DefaultView.RowFilter = string.Format("[Table_NameId] = '{0}'", TableID);

            ColumnDGComboboxUpdater();
        }


        private void SafeColumnDGUpdater(long TableID)
        {
            if (dgColumns.InvokeRequired)
            {
                //                var d = new SafeCallDelegateLong(ThreadSafeColumnDataGridUpdater(TableID));
                Invoke(new MethodInvoker(delegate ()
                {
                    SafeColumnDGUpdater(TableID);
                }));
                //dgColumns.Invoke(d, new object[] { });
            }
            else
            {
                ColumnDGUpdater(TableID);
            }
        }


        private void ThreadSafeListViewUpdater()
        {
            if (cboToDB.InvokeRequired)
            {
                var d = new SafeCallDelegate(ThreadSafeListViewUpdater);
                dgTables.Invoke(d, new object[] { });
            }
            else
            {
                cboFromDB.Items.Clear();
                cboToDB.Items.Clear();

                List<ComboListInfo> toList = new List<ComboListInfo>();
                List<ComboListInfo> fromList = new List<ComboListInfo>();
                foreach (DataRow rec in dBDataTable.Rows)
                {
                    ComboListInfo item = new ComboListInfo((long)rec["Id"], rec["AppAndDB"].ToString(), "", 0);
                    toList.Add(item);
                    fromList.Add(item);
                }

                cboFromDB.DataSource = fromList;
                cboFromDB.Sorted = true;
                cboFromDB.DisplayMember = "Name";
                cboFromDB.ValueMember = "ID";
                cboFromDB.SelectedIndex = -1;

                cboToDB.DataSource = toList;
                cboFromDB.Sorted = true;
                cboToDB.DisplayMember = "Name";
                cboToDB.ValueMember = "ID";
                cboToDB.SelectedIndex = -1;

                bFormLoadComplete = true;
            }
        }


        /************************************* Retrieve Data *********************************************/

        private async void GetDBList()
        {
            DataTable dbDT = await sharepointManager.GetDBsFromSharepoint("", new string[] { "Id", "Application_NameId", "Title" }, new string[] { "Title" });

            dBDataTable = dbDT;
            ThreadSafeListViewUpdater();
        }

        private async void GetSourceTableList(long DBID)
        {
            try
            {
                sourceTableDataTable = await sharepointManager.GetTablesFromSharepoint("", DBID, new string[] { "ID", "Table_Name", "Database_NameId", "Description" }, new string[] { });
                if (sourceTableDataTable == null) return;
                sourceTableDataTable.Columns.Add("Edit_Status");
                sourceTableDataTable.Columns.Add("Destination_Table", typeof(long));
                SafeTableDGUpdate();
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }

        private async void GetDestTableList(long DBID)
        {
            try
            {
                destTableDataTable = await sharepointManager.GetTablesFromSharepoint("", DBID, new string[] { "ID", "Table_Name", "Database_NameId", "Description" }, new string[] { });
                if (destTableDataTable == null) return;
                SafeTableDGComboboxUpdater();
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }

        private async void GetSourceColumnListAsyncWrapper(long DBID, long TableID, bool updateGrid = true)
        {
            await GetSourceColumnList(DBID, TableID, updateGrid);
        }

        private async Task GetSourceColumnList(long DBID, long TableID, bool updateGrid = true)
        {
            try
            {
                if (sourceColumnDataTable == null)
                    sourceColumnDataTable = await sharepointManager.GetColumnsFromSharepoint(DBID, TableID, new string[] { "ID", "Table_NameId", "Description", "Title", "Data_Type" }, new string[] { "Table_NameID", "Column_Name" });
                else if (sourceColumnDataTable.Select("Table_NameId=" + TableID.ToString()).Length > 0)
                    sourceColumnDataTable.Merge( await sharepointManager.GetColumnsFromSharepoint(DBID, TableID, new string[] { "ID", "Table_NameId", "Description", "Title", "Data_Type" }, new string[] { "Table_NameID", "Column_Name" }));

                if (updateGrid & sourceColumnDataTable != null)
                    SafeColumnDGComboboxUpdater();
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }

        private async void GetDestColumnListAsyncWrapper(long DBID, long TableID, bool updateGrid = true)
        {
            await GetDestColumnList(DBID, TableID, updateGrid);
        }
        private async Task<bool> GetDestColumnList(long DBID, long TableID, bool updateGrid = true)
        {
            try
            {
                if (destColumnDataTable == null)
                    destColumnDataTable = await sharepointManager.GetColumnsFromSharepoint(DBID, TableID, new string[] { "ID", "Table_NameId", "Description", "Title", "Data_Type" }, new string[] { "Table_NameID", "Column_Name" });
                else if (destColumnDataTable.Select("Table_NameId=" + TableID.ToString()).Length > 0)
                    destColumnDataTable.Merge(await sharepointManager.GetColumnsFromSharepoint(DBID, TableID, new string[] { "ID", "Table_NameId", "Description", "Title", "Data_Type" }, new string[] { "Table_NameID", "Column_Name" }));

                if (updateGrid & destColumnDataTable != null)
                    SafeColumnDGUpdater(TableID);
            }
            catch (Exception e) { Console.WriteLine(e.Message); return false; }
            return true;
        }



        /************************ Control Events ***************************/

        private void cboFromDB_SelectedValueChanged(object sender, EventArgs e)
        {
            if (!bFormLoadComplete) return;
            GetSourceTableList((long)cboFromDB.SelectedValue);
            if (cboFromDB.SelectedItem is null) return;
            var db = (cboFromDB.SelectedItem as ComboListInfo).ID;
            //GetSourceColumnListAsyncWrapper(db);
        }


        private void cboToDB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bFormLoadComplete) return;
            GetDestTableList((long)cboToDB.SelectedValue);
            if (cboToDB.SelectedItem is null) return;
            var db = (cboToDB.SelectedItem as ComboListInfo).ID;
            //GetDestColumnList(db);
        }


        /******************** Tables Grid Events ****************************/
        private void dgTables_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            bool validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            var datagridview = sender as DataGridView;

            // Check to make sure the cell clicked is the cell containing the combobox 
            if (datagridview.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && validClick)
            {
                datagridview.BeginEdit(true);
                ((ComboBox)datagridview.EditingControl).DroppedDown = true;
            }
        }

        private void dgTables_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgTables.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgTables_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            long TableID = 0;
            if (dgTables["Destination Table", e.RowIndex].Value.ToString() == "")
                SafeColumnDGUpdater(-1);
            else
                if (long.TryParse(dgTables["Destination Table", e.RowIndex].Value.ToString(), out TableID))
                SafeColumnDGUpdater(TableID);
        }

        private void AddNewDestinationTable(long FromTableID, string FromTableName)
        {
            string input1 = FromTableName, input2 = "-1";
            Utilites.ScreenEvents.ShowInputDialog(ref input1, ref input2, "New Table", "", "Enter New Destination Table Name");
        }


        private void dgTables_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            bool validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            var datagridview = sender as DataGridView;

            // Check to make sure the cell clicked is the cell containing the combobox 
            if (datagridview.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && validClick)
            {
                DataGridViewComboBoxCell DestTable = (DataGridViewComboBoxCell)datagridview[e.ColumnIndex, e.RowIndex];
                if (DestTable is null) return;
                if (DestTable.Value == null)
                {
                    SafeColumnDGUpdater(-1); // User selected blank
                }
                else if ((long)DestTable.Value == 0)
                {
                    AddNewDestinationTable((long)DestTable.Value, datagridview["Table_Name", e.RowIndex].Value.ToString());
                    datagridview["Edit_Status", e.RowIndex].Value = "Add";
                }
                else
                {
                    GetDestColumnListAsyncWrapper((long)(cboToDB.SelectedItem as ComboListInfo).ID, (long)DestTable.Value);
                    datagridview["Edit_Status", e.RowIndex].Value = "Edit";
                    GetSourceColumnListAsyncWrapper((long)cboFromDB.SelectedValue, (long)datagridview["ID", e.RowIndex].Value);
                }
            }
            //if (datagridview.Columns[e.ColumnIndex].Name == "Edit_Status") return;
            //datagridview["Edit_Status", e.RowIndex].Value = "Updated";
        }


        /******************** Columns Grid Events ****************************/

        private void dgColumns_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            bool validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            var datagridview = sender as DataGridView;

            // Check to make sure the cell clicked is the cell containing the combobox 
            if (datagridview.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && validClick)
            {
                datagridview.BeginEdit(true);
                ((ComboBox)datagridview.EditingControl).DroppedDown = true;
            }
        }


        private void dgColumns_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgColumns.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }


        private void dgColumns_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            bool validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            var datagridview = sender as DataGridView;

            // Check to make sure the cell clicked is the cell containing the combobox 
            if (datagridview.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && validClick)
            {
                datagridview["Relation_Edit_Status", e.RowIndex].Value = "Updated";
            }
            else
            {
                //ignore if we are auto-creating these changes.
                if (datagridview.Columns[e.ColumnIndex].Name == "Column_Edit_Status") return;
                datagridview["Column_Edit_Status", e.RowIndex].Value = "Updated";
            }
        }


        private void SaveTableRecordDR(DataRow row)
        {

        }

        private void SaveDataFlowRecordDR(string Entity, string FlowType, DataRow row)
        {

        }

        private void SaveColumnRecordDR(DataRow row)
        {

        }

        internal bool dtDataFieldExists(DataTable dt, string fieldName)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
                if (dt.Columns[i].ColumnName == fieldName)
                    return true;
            return false;
        }

                        /* dtRelation must contain the following case-sensitive fieldnames: ID, Table_OneId, Table_One, Table_ManyId, Table_Many,
                 * Column_OneId, Column_ManyId, Column_ManyId, Column_Many, Relation_Type, Relation_Level_One, Relation_Level_Many
                 * Table_OneId must be a valid table or we have to look up using Table_One
                 * Table_ManyId must be a valid table or we have to look up using Table_Many
                 * Column_OneId can be empty. If it is not empty it must be valid or Column_One must be valid
                 * Column_ManyId can be empty. If it is not empty it must be valid or Column_Many must be valid
                 * Relation_LevelOne and Relation_LevelMany 
                 *          valid types can be "Table", "Application", "Actor", 
                 *          When linking columns, set Relation_Levels to "Table" as we have added secondary fields for columns
                 * Every other field (Description, etc) will be uploaded into Sharepoint if it matches the Sharepoint field name or ignored*/


        private void btnSave_Click(object sender, EventArgs e)
        {
            //Run through all of the columns (we may add tables later)
            DataTable dtRelations = new DataTable();
            dtRelations = destColumnDataTable.Copy();
            if (dtRelations == null)
                MessageBox.Show("No relation records to be saved");
            else
            {
                dtRelations.Columns.Add("Table_One");                       //rename
                dtRelations.Columns.Add("Table_Many");
                dtRelations.Columns.Add("Column_One");                      //blanks are fine here
                dtRelations.Columns.Add("Column_Many");
                if (!dtDataFieldExists(dtRelations, "Table_OneId"))         dtRelations.Columns.Add("Table_OneId", typeof(long)); // <== this requires manual addition
                if (!dtDataFieldExists(dtRelations, "Table_ManyId"))         dtRelations.Columns.Add("Table_ManyId", typeof(long), "Table_NameId"); //rename
                if (!dtDataFieldExists(dtRelations, "Column_OneId"))        dtRelations.Columns.Add("Column_OneId", typeof(long), "Source_Column");
                if (dtDataFieldExists(dtRelations, "Id"))                   dtRelations.Columns["Id"].ColumnName = "Column_ManyId";
                if (!dtDataFieldExists(dtRelations, "Relation_Type"))       dtRelations.Columns.Add("Relation_Type", typeof(string), "'Data Flow'");
                if (!dtDataFieldExists(dtRelations, "Relation_Level_One"))  dtRelations.Columns.Add("Relation_Level_One", typeof(string), "'Column'");
                if (!dtDataFieldExists(dtRelations, "Relation_Level_Many")) dtRelations.Columns.Add("Relation_Level_Many", typeof(string), "'Column'");
                if (!dtDataFieldExists(dtRelations, "ID"))                  dtRelations.Columns.Add("ID",typeof(long));
                if (dtDataFieldExists(dtRelations, "Relation_Edit_Status")) dtRelations.Columns["Relation_Edit_Status"].ColumnName = "Edit_Status";
                foreach (DataRow row in dtRelations.Rows)
                {
                    if (row["Source_Column"].ToString() == "") continue;
                    var ret2 = sourceColumnDataTable.Select("[ID]='" + row["Column_OneId"].ToString() + "'");
                    if (ret2.Length > 0)
                    {
                        row["Table_OneId"] = (long)ret2[0]["Table_NameId"];
                    }
                }

                sharepointManager.saveRelationsToSharepoint(dtRelations, (cboFromDB.SelectedItem as ComboListInfo).ID, (cboToDB.SelectedItem as ComboListInfo).ID);
            }
        }


        private void SafeDGUpdateFocus(DataGridView dg, string colName, int row)
        {
            try
            {
                if (dg.InvokeRequired)
                    Invoke(new MethodInvoker(delegate () { SafeDGUpdateFocus(dg, colName, row); }));
                else
                    dg.CurrentCell = dg[colName, row];
            }
            catch(Exception e) { Console.WriteLine(e.Message); }
                Utilites.ScreenEvents.DoEvents();
        }


        private async void MatchAllTables(long FromDB, long ToDB)
        {
            for (int i = 0; i < dgTables.Rows.Count; i++)
            {
                SafeDGUpdateFocus(dgTables, "Table_Name", i);
                var sourceTable = getTableNameOnly(dgTables["Table_Name", i].Value.ToString());
                foreach (DataRow destRow in destTableDataTable.Rows)
                {
                    var destTable = getTableNameOnly(destRow["Table_Name"].ToString());
                    if (sourceTable.ToLower() == destTable.ToLower())
                    {
                        var sourceRow = sourceTableDataTable.Select("ID=" + dgTables["ID", i].Value.ToString());
                        sourceRow[0]["Destination_Table"] = destRow["ID"];
                        //this is set to new for the relationship.  If the Destination_Table has no ID, we need to add it as well.
                        sourceRow[0]["Edit_Status"] = "New"; 

                        await GetSourceColumnList(FromDB, (long)sourceRow[0]["ID"], false);
                        await GetDestColumnList(ToDB, (long)destRow["ID"], false);

                        Utilites.ScreenEvents.DoEvents();
                        SafeColumnDGUpdater((long)sourceRow[0]["ID"]);
                        Utilites.ScreenEvents.DoEvents();
                        MatchAllColumnsSilently((long)sourceRow[0]["ID"], (long)destRow["ID"]);
                    }
                }
            }
        }


        private void MatchAllColumnsSilently(long SourceTableID, long DestTableID)
        {
            var sourceRowArr = sourceColumnDataTable.Select("Table_NameId=" + SourceTableID.ToString());
            if (sourceRowArr.Length == 0) return;
            DataTable sourceRows = sourceRowArr.CopyToDataTable();

            var destRow = destColumnDataTable.Select("Table_NameId=" + DestTableID.ToString());
            if (destRow.Length == 0) return;

            //destColumnDataTable.DefaultView.RowFilter = "";
            for(int i=0; i<destRow.Length; i++)
            {
                var destColName = destRow[i]["Column_Name"].ToString();
                foreach (DataRow sourceRow in sourceRows.Rows)
                {
                var sourceColName = sourceRow["Column_Name"].ToString();
                    if (sourceColName.ToLower() == destColName.ToLower())
                    {
                        destRow[i]["Source_Column"] = sourceRow["ID"];
                        destRow[i]["Relation_Edit_Status"] = "New";
                        if (destRow[i]["Description"].ToString() == "" & sourceRow["Description"].ToString() != "")
                        {
                            destRow[i]["Description"] = sourceRow["Description"].ToString();
                            destRow[i]["Column_Edit_Status"] = "Updated";
                            destRow[i]["Source_TableId"] = sourceRow["Table_NameId"];
                        }

                        Utilites.ScreenEvents.DoEvents();
                        continue;
                    }
                }
            }
        }


        //private void MatchAllColumns(long SourceTableID)
        //{
        //    //destColumnDataTable.DefaultView.RowFilter = "";
        //    for (int i = 0; i < dgColumns.Rows.Count; i++)
        //    {
        //        SafeDGUpdateFocus(dgColumns, "Column_Name", i);
        //        //var sourceTableID = 
        //        var sourceRowArr = sourceColumnDataTable.Select("Table_NameId=" + SourceTableID.ToString());
        //        if (sourceRowArr.Length == 0) continue;
        //        DataTable source = sourceRowArr.CopyToDataTable();
        //        try
        //        {   // sometimes i is incremented beyond the last row.
        //            var destColName = getTableNameOnly(dgColumns["Column_Name", i].Value.ToString());
        //            foreach (DataRow sourceRow in source.Rows)
        //            {
        //                var sourceColName = getTableNameOnly(sourceRow["Column_Name"].ToString());
        //                if (sourceColName.Equals( destColName,StringComparison.OrdinalIgnoreCase))
        //                {
        //                    var destRow = destColumnDataTable.Select("ID=" + dgColumns["ID", i].Value.ToString());
        //                    destRow[0]["Source_Column"] = sourceRow["ID"];
        //                    destRow[0]["Relation_Edit_Status"] = "New";
        //                    if (destRow[0]["Description"].ToString() == "" & sourceRow["Description"].ToString() != "")
        //                    {
        //                        destRow[0]["Description"] = sourceRow["Description"].ToString();
        //                        destRow[0]["Column_Edit_Status"] = "Updated";
        //                        destRow[0]["Source_TableId"] = sourceRow["Table_NameId"];
        //                    }

        //                    Utilites.ScreenEvents.DoEvents();
        //                    continue;
        //                }
        //            }
        //        }
        //        catch (Exception e){ Console.WriteLine(e.Message); }
        //    }
        //}


        private string getTableNameOnly(string fullName)
        {
            return fullName.Split('.').Last();
        }

        private void btnAutoMatchTables_Click(object sender, EventArgs e)
        {
            MatchAllTables((cboFromDB.SelectedItem as ComboListInfo).ID, (cboToDB.SelectedItem as ComboListInfo).ID);
            dgTables.CurrentCell = dgTables["Table_Name", 0];
            SafeColumnDGUpdater((long)dgTables["ID", 0].Value);

        }

        private void btnAutoMatchColumns_Click(object sender, EventArgs e)
        {
            //            MatchAllColumns((long)dgTables["ID", dgTables.CurrentCell.RowIndex].Value);

            dgColumns.DataSource = null;
            foreach (DataRow row in sourceTableDataTable.Rows)
            {
                if (row["Destination_Table"].ToString() != "")
                    MatchAllColumnsSilently((long)row["ID"] , (long)row["Destination_Table"]);
            }
            dgTables.CurrentCell = dgTables["Table_Name", 0];
            SafeColumnDGUpdater((long)dgTables["ID", 0].Value);
        }


        private void WriteToExcel(DataTable dt)
        {
            XL.Application xl = new XL.Application();
            xl.Visible = true;
            var wb = xl.Workbooks.Add();
            XL.Worksheet ws = wb.ActiveSheet;

            // column headings               
            int ColumnsCount = dt.Columns.Count;
            object[] Header = new object[ColumnsCount];

            for (int i = 0; i < ColumnsCount; i++)
                Header[i] = dt.Columns[i].ColumnName;

            Microsoft.Office.Interop.Excel.Range HeaderRange = ws.get_Range((Microsoft.Office.Interop.Excel.Range)(ws.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(ws.Cells[1, ColumnsCount]));
            HeaderRange.Value = Header;
            HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            HeaderRange.Font.Bold = true;

            // DataCells
            int RowsCount = dt.Rows.Count;
            object[,] Cells = new object[RowsCount, ColumnsCount];

            for (int j = 0; j < RowsCount; j++)
                for (int i = 0; i < dt.Columns.Count; i++)
                    Cells[j, i] = dt.Rows[j][i];

            ws.get_Range((Microsoft.Office.Interop.Excel.Range)(ws.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(ws.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;



        }


        private async void GetMyStuff(long FromDB, long ToDB)
        {
            var fromCols = await sharepointManager.GetColumnsFromSharepointByDB(FromDB);
            var toCols = await sharepointManager.GetColumnsFromSharepointByDB(ToDB);
            var rel = await sharepointManager.GetRelationsFromSharepointByDB(ToDB, true);

            WriteToExcel(toCols);
            WriteToExcel(rel);
        }

        private void btnDocumentFlow_Click(object sender, EventArgs e)
        {
            GetMyStuff((long)cboFromDB.SelectedValue, (long)cboToDB.SelectedValue);
        }
    }

    public class ComboListInfo
    {
        internal ComboListInfo() { }
        internal ComboListInfo(long ID, string Name, string Type, long ParentID)
        {
            this.ID = ID;
            this.Name = Name;
            this.Type = Type;
            this.ParentID = ParentID;
        }
        public long ID { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public long ParentID { get; set; }
    }
}
