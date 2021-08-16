using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
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
            destColumnDataTable.DefaultView.RowFilter = string.Format("[Table_NameId]={0}", TableID);

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
            DataTable dbDT = await sharepointManager.GetDBsFromSharepoint(0);

            dBDataTable = dbDT;
            ThreadSafeListViewUpdater();
        }

        private async void GetSourceTableList(long DBID)
        {
            try
            {
                sourceTableDataTable = await sharepointManager.GetTablesFromSharepoint(DBID);
                if (sourceTableDataTable == null) return;
                sourceTableDataTable.Columns.Add("Edit_Status");
                sourceTableDataTable.Columns.Add("Destination_Table", typeof(long));
                SafeTableDGUpdate();
            }
            catch (Exception e) { Debug.WriteLine(e.Message); }
        }

        private async void GetDestTableList(long DBID)
        {
            try
            {
                destTableDataTable = await sharepointManager.GetTablesFromSharepoint(DBID);
                if (destTableDataTable == null) return;
                destTableDataTable.Columns.Add("Edit_Status");
                SafeTableDGComboboxUpdater();
            }
            catch (Exception e) { Debug.WriteLine(e.Message); }
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
                    sourceColumnDataTable = await sharepointManager.GetColumnsFromSharepoint(DBID, TableID);
                else if (sourceColumnDataTable.Select("Table_NameId=" + TableID.ToString()).Length == 0)
                {
                    DataTable tmp = await sharepointManager.GetColumnsFromSharepoint(DBID, TableID);
                    if (tmp != null)
                        sourceColumnDataTable.Merge(tmp, true, MissingSchemaAction.Ignore);
                }

                if (updateGrid & sourceColumnDataTable != null)
                    SafeColumnDGComboboxUpdater();
            }
            catch (Exception e) { Debug.WriteLine(e.Message); }
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
                    destColumnDataTable = await sharepointManager.GetColumnsFromSharepoint(DBID, TableID);
                else if (destColumnDataTable.Select("Table_NameId=" + TableID.ToString()).Length == 0)
                {
                    DataTable tmp = await sharepointManager.GetColumnsFromSharepoint(DBID, TableID);
                    if (tmp != null) 
                    destColumnDataTable.Merge(tmp,true,MissingSchemaAction.Ignore);
                }
                if (updateGrid & destColumnDataTable != null)
                    SafeColumnDGUpdater(TableID);
            }
            catch (Exception e) { Debug.WriteLine(e.Message); return false; }
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
                else if (DestTable.Value is DBNull)
                {
                    //GG: something's not working here
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

        private void SaveColumnRelations(DataRow row)
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
            long destDB = (long)cboToDB.SelectedValue;           //(cboFromDB.SelectedItem as ComboListInfo).ID
            long sourceDB = (long)cboFromDB.SelectedValue;

            // save each of the column links
            var query = from dest in destColumnDataTable.AsEnumerable()
                        join source in sourceColumnDataTable.AsEnumerable() on dest["Source_Column"].ToString() equals source["ID"].ToString()
                        where dest["Source_Column"].ToString() != ""
                        select new
                        {
                            Database_One = source["Database_Name"].ToString(),
                            Database_Many = dest["Database_Name"].ToString(),
                            Table_One = "",
                            Table_Many = "",
                            Column_One = "",
                            Column_Many = "",
                            Table_OneId = (long?)source["Table_NameId"],
                            Table_ManyId = (long?)dest["Table_NameId"],
                            Column_OneId = (long?)dest["Source_Column"],
                            Column_ManyId = (long?)dest["ID"],
                            Relation_Type = "Data Flow",
                            Relation_Level_One = "Column",
                            Relation_Level_Many = "Column",
                            ID = (long?)null,//DBNull.Value,
                            Edit_Status = "New"
                        };

            DataTable newRelations = DataTableMethods.ConvertToDataTable(query.ToList(), "Relations");// = query.CopyToDataTable();
            sharepointManager.saveRelationsToSharepoint(newRelations, (cboFromDB.SelectedItem as ComboListInfo).ID, (cboToDB.SelectedItem as ComboListInfo).ID);

            // save links at the table level
            var query2 = (from dest in destColumnDataTable.AsEnumerable()
                          join source in sourceColumnDataTable.AsEnumerable() on dest["Source_Column"].ToString() equals source["ID"].ToString()
                          where dest["Source_Column"].ToString() != ""
                          select new
                          {
                              Database_One = source["Database_Name"].ToString(),
                              Database_Many = dest["Database_Name"].ToString(),
                              Table_One = "",
                              Table_Many = "",
                              Column_One = "",
                              Column_Many = "",
                              Table_OneId = (long?)source["Table_NameId"],
                              Table_ManyId = (long?)dest["Table_NameId"],
                              Column_OneId = (long)0,
                              Column_ManyId = (long)0,
                              Relation_Type = "Data Flow",
                              Relation_Level_One = "Table",
                              Relation_Level_Many = "Table",
                              ID = (long?)null,
                              Edit_Status = "New"
                          }).Distinct();

            DataTable newRelations2 = DataTableMethods.ConvertToDataTable(query2.ToList(), "Relations");// = query.CopyToDataTable();
            sharepointManager.saveRelationsToSharepoint(newRelations2, sourceDB, destDB);

            // if the destination table description is null, copy it from the source
            var query3 = from dest in destTableDataTable.AsEnumerable()
                          join source in sourceTableDataTable.AsEnumerable() on dest["ID"].ToString() equals source["Destination_Table"].ToString()
                          where (dest["Description"].ToString().Length == 0) && (source["Description"].ToString().Length > 0) 
                          select new
                          {
                              ID = (long?)dest["ID"],
                              Table_Name = dest["Table_Name"].ToString(),
                              Description = source["Description"].ToString(),
                              Database_Name = dest["Database_Name"].ToString(),
                              Database_NameId = (long) destDB,
                              Edit_Status = "New"
                          };

            DataTable newRelations3 = DataTableMethods.ConvertToDataTable(query3.ToList(), "Relations");// = query.CopyToDataTable();
            sharepointManager.saveTablesToSharepoint(newRelations3);

            // if the destination column description is null, copy it from the source
            var query4 = from dest in destColumnDataTable.AsEnumerable()
                         join source in sourceColumnDataTable.AsEnumerable() on dest["Source_Column"].ToString() equals source["ID"].ToString()
                         where (dest["Description"].ToString().Length == 0) && (source["Description"].ToString().Length > 0)
                         select new
                         {
                             ID = (long?)dest["ID"],
                             Column_Name = dest["Column_Name"].ToString(),
                             Table_NameId = dest["Table_NameId"].ToString(),
                             Table_Name = dest["Table_Name"].ToString(),
                             Description = source["Description"].ToString(),
                             Database_Name = dest["Database_Name"].ToString(),
                             Database_NameId = destDB,
                             Edit_Status = "New"
                         };

            DataTable newRelations4 = DataTableMethods.ConvertToDataTable(query4.ToList(), "Relations");// = query.CopyToDataTable();
            sharepointManager.saveColumnsToSharepoint(newRelations4, destDB);

        }


        internal DataTable CreateDataTable(string tableName, string fieldNames)
        {
            DataTable dataTable = new DataTable();
            var names = fieldNames.Split(',');
            foreach (var field in names)
                if (field.EndsWith("ID", StringComparison.OrdinalIgnoreCase))
                    dataTable.Columns.Add(field, typeof(long));
                else
                    dataTable.Columns.Add(field, typeof(string));
            return dataTable;
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
            catch(Exception e) { Debug.WriteLine(e.Message); }
                Utilites.ScreenEvents.DoEvents();
        }


        private async void MatchAllTables(long FromDB, long ToDB)
        {
            //await GetSourceColumnList(FromDB, 0, false);
            //await GetDestColumnList(ToDB, 0, false);

            for (int i = 0; i < dgTables.Rows.Count; i++)
            {
                SafeDGUpdateFocus(dgTables, "Table_Name", i);
                var sourceTable = getTableNameOnly(dgTables["Table_Name", i].Value.ToString());

                foreach (DataRow destRow in destTableDataTable.Rows)
                {
                    var destTable = getTableNameOnly(destRow["Table_Name"].ToString());
                    if (sourceTable.ToLower() == destTable.ToLower())
                    {
                        var sourceRows = sourceTableDataTable.Select("ID=" + dgTables["ID", i].Value.ToString());
                        var sourceRow = sourceRows[0];
                        sourceRow["Destination_Table"] = destRow["ID"];
                        if (destRow["Description"].ToString() == "")
                        {
                            destRow["Description"] = sourceRow["Description"];
                            destRow["Edit_Status"] = "Update";
                        }
                        //this is set to new for the relationship.  If the Destination_Table has no ID, we need to add it as well.
                        sourceRow["Edit_Status"] = "New";

                        await GetSourceColumnList(FromDB, (long)sourceRow["ID"], false);
                        await GetDestColumnList(ToDB, (long)destRow["ID"], false);

                        Utilites.ScreenEvents.DoEvents();
                        SafeColumnDGUpdater((long)sourceRow["ID"]);
                        Utilites.ScreenEvents.DoEvents();
                        MatchAllColumnsSilently((long)sourceRow["ID"], (long)destRow["ID"]);
                    }
                }
            }
        }


        private void MatchAllColumnsSilently(long SourceTableID, long DestTableID)
        {
            var sourceRowArr = sourceColumnDataTable.Select("Table_NameId=" + SourceTableID.ToString());
            if (sourceRowArr.Length == 0) return;
            //DataTable sourceRows = sourceRowArr.CopyToDataTable();

            var destRow = destColumnDataTable.Select("Table_NameId=" + DestTableID.ToString());
            if (destRow.Length == 0) return;

            //destColumnDataTable.DefaultView.RowFilter = "";
            for(int i=0; i<destRow.Length; i++)
            {
                var destColName = destRow[i]["Column_Name"].ToString();
                foreach (DataRow sourceRow in sourceRowArr)// sourceRows.Rows)
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

            foreach(string s in new string[] {"FileSystemObjectType","Id","ServerRedirectedEmbedUri","ServerRedirectedEmbedUrl","ContentTypeId","Title","ComplianceAssetId","Table_OneId","Table_ManyId"
                ,"Column_OneId","Column_ManyId","Relation_Level_One","Relation_Level_Many", "ExternalID" ,"Intersect_Entity","Modified","Created","AuthorId","EditorId","OData__UIVersionString","Attachments","GUID" })
                if(dtDataFieldExists(dt, s)) dt.Columns.Remove(s);

            // column headings               
            int ColumnsCount = dt.Columns.Count;
            object[] Header = new object[ColumnsCount];

            for (int i = 0; i < ColumnsCount; i++)
                Header[i] = dt.Columns[i].ColumnName;

            dt.Columns["Description2"].SetOrdinal(0);
            dt.Columns["Column_Many"].SetOrdinal(0);
            dt.Columns["Table_Many"].SetOrdinal(0);
            dt.Columns["Database_Many"].SetOrdinal(0);

            Microsoft.Office.Interop.Excel.Range HeaderRange = ws.get_Range((Microsoft.Office.Interop.Excel.Range)(ws.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(ws.Cells[1, ColumnsCount]));
            HeaderRange.Value = Header;
            HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            HeaderRange.Font.Bold = true;

            // DataCells
            int RowsCount = dt.Rows.Count;
            object[,] Cells = new object[RowsCount, ColumnsCount + 1];

            //string lastRow = "";
            for (int j = 0; j < RowsCount; j++)
                for (int i = 0; i < dt.Columns.Count; i++)
                    Cells[j, i] = (dt.Rows[j][i].ToString().StartsWith("=") ? "'" + dt.Rows[j][i] : dt.Rows[j][i]);

            ws.get_Range((Microsoft.Office.Interop.Excel.Range)(ws.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(ws.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;
            ws.Cells[2][2].Select();
            xl.ActiveWindow.FreezePanes = true;
            xl.Selection.AutoFilter();

            XL.Range sel = xl.Selection;
            sel.Subtotal(1, XL.XlConsolidationFunction.xlCount, new int[] { 2, 32 }, true, false,XL.XlSummaryRow.xlSummaryAbove);
            ws.Columns["B:B"].EntireColumn.AutoFit();

//            subGroupTest(ws);
        }


        private void subGroupTest(XL.Worksheet ws)
        {
            XL.Range sRng=null, eRng=null;
            string[,] groupMap = new string[8000, 100];

            XL.Range currRng = ws.Range["E1"];

            while (currRng.Value != "") {
                if (sRng == null)
                {
                    //                    ' If start-range is empty, set start-range to current range
                    sRng = currRng;
                }
                else
                {
                    //' Start-range not empty
                    //            ' If current range and start range match, we've reached the same index & need to terminate
                    if (currRng.Value != (sRng.Value ?? ""))
                        eRng = currRng;

                    if (currRng.Value == sRng.Value | currRng.Offset[1].Value == "")
                    {
                        XL.Range rng = ws.Range[sRng.Offset[1], eRng];
                        rng.EntireRow.Group();
                        sRng = currRng;
                        eRng = null;
                    }
                }
                currRng = currRng.Offset[1];
            }
        }



        private async void DocumentDataFlowInExcel(long FromDB, long ToDB)
        {
            var fromCols = await sharepointManager.GetColumnsFromSharepoint(FromDB, 0);
            var toCols = await sharepointManager.GetColumnsFromSharepoint(ToDB, 0);
            var dtRelation = await sharepointManager.GetRelationsFromSharepointByDB(ToDB, "Data Flow", true);

            // if the destination column description is null, copy it from the source
            var query = from dest in toCols.AsEnumerable()
                        join relation in dtRelation.AsEnumerable() on (long)dest["ID"] equals (long)relation["Column_ManyId"] into relationSub
                        from relation in relationSub.DefaultIfEmpty()
                        join source in toCols.AsEnumerable() on relation["Column_OneId"].ToString() equals source["ID"].ToString() into sourceSub
                        from source in sourceSub.DefaultIfEmpty()
                        select new
                        {
                            Database = dest["Database_Name"],
                            Table = dest["Table_Name"],
                            Column = dest["Column_Many"],

                            From_Database = source["Database_Name"] ?? string.Empty,
                            From_Table = source["Table_Name"] ?? string.Empty,
                            From_Column = source["Column_Many"] ?? string.Empty,

                            RelationType = "Data Flow"
                         };

            DataTable newRelations4 = DataTableMethods.ConvertToDataTable(query.ToList(), "Relations");// = query.CopyToDataTable();
            sharepointManager.saveColumnsToSharepoint(newRelations4, ToDB);

            WriteToExcel(dtRelation.Copy());
            Debug.WriteLine("Done.");


            try
            {
                // Use Linq to add Table and DB fields to each record
                dtRelation.Columns.Add("Database_One", typeof(string));
                dtRelation.Columns.Add("Table_One", typeof(string));
                dtRelation.Columns.Add("Column_One", typeof(string));
                dtRelation.Columns.Add("Description1", typeof(string));

                dtRelation.Columns.Add("Database_Many", typeof(string));
                dtRelation.Columns.Add("Table_Many", typeof(string));
                dtRelation.Columns.Add("Column_Many", typeof(string));
                dtRelation.Columns.Add("Description2", typeof(string));

                dtRelation.Columns.Add("Data_Type1", typeof(string));
                dtRelation.Columns.Add("Character_Maximum_Length1", typeof(string));
                dtRelation.Columns.Add("Character_Octet_Length1", typeof(string));
                dtRelation.Columns.Add("Date_Time_Precision1", typeof(string));
                dtRelation.Columns.Add("Default1", typeof(string));
                dtRelation.Columns.Add("Is_Nullable1", typeof(string));
                dtRelation.Columns.Add("Is_Primary1", typeof(string));
                dtRelation.Columns.Add("Numeric_Precision1", typeof(string));
                dtRelation.Columns.Add("Numeric_Precision_Radix1", typeof(string));
                dtRelation.Columns.Add("Numeric_Scale1", typeof(string));
                dtRelation.Columns.Add("Ordinal_Position1", typeof(string));

                dtRelation.Columns.Add("Data_Type2", typeof(string));
                dtRelation.Columns.Add("Character_Maximum_Length2", typeof(string));
                dtRelation.Columns.Add("Character_Octet_Length2", typeof(string));
                dtRelation.Columns.Add("Date_Time_Precision2", typeof(string));
                dtRelation.Columns.Add("Default2", typeof(string));
                dtRelation.Columns.Add("Is_Nullable2", typeof(string));
                dtRelation.Columns.Add("Is_Primary2", typeof(string));
                dtRelation.Columns.Add("Numeric_Precision2", typeof(string));
                dtRelation.Columns.Add("Numeric_Precision_Radix2", typeof(string));
                dtRelation.Columns.Add("Numeric_Scale2", typeof(string));
                dtRelation.Columns.Add("Ordinal_Position2", typeof(string));


                dtRelation.AsEnumerable().Join(toCols.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Column_ManyId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => {
                            o._dtmater.SetField("Column_Many", o._dtchild["Column_Name"].ToString());
                            //o._dtmater.SetField("Table_ManyId", (long)o._dtchild["Table_NameId"]);
                            o._dtmater.SetField("Table_Many", o._dtchild["Table_Name"].ToString());
                            o._dtmater.SetField("Database_Many", o._dtchild["Database_Name"].ToString());
                            o._dtmater.SetField("Data_Type2", o._dtchild["Data_Type"].ToString());
                            o._dtmater.SetField("Character_Maximum_Length2", o._dtchild["Character_Maximum_Length"].ToString());
                            o._dtmater.SetField("Character_Octet_Length2", o._dtchild["Character_Octet_Length"].ToString());
                            o._dtmater.SetField("Date_Time_Precision2", o._dtchild["Date_Time_Precision"].ToString());
                            o._dtmater.SetField("Default2", o._dtchild["Default"].ToString());
                            o._dtmater.SetField("Is_Nullable2", o._dtchild["Is_Nullable"].ToString());
                            o._dtmater.SetField("Is_Primary2", o._dtchild["Is_Primary"].ToString());
                            o._dtmater.SetField("Numeric_Precision2", o._dtchild["Numeric_Precision"].ToString());
                            o._dtmater.SetField("Numeric_Precision_Radix2", o._dtchild["Numeric_Precision_Radix"].ToString());
                            o._dtmater.SetField("Numeric_Scale2", o._dtchild["Numeric_Scale"].ToString());
                            o._dtmater.SetField("Ordinal_Position2", o._dtchild["Ordinal_Position"].ToString());
                            o._dtmater.SetField("Description2", o._dtchild["Description"].ToString());

                        }
                    );

                dtRelation.AsEnumerable().Join(toCols.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Column_OneId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => {
                            o._dtmater.SetField("Column_One", o._dtchild["Column_Name"].ToString());
                            //o._dtmater.SetField("Table_OneId", (long)o._dtchild["Table_NameId"]);
                            o._dtmater.SetField("Table_One", o._dtchild["Table_Name"].ToString());
                            o._dtmater.SetField("Database_One", o._dtchild["Database_Name"].ToString());
                            o._dtmater.SetField("Data_Type1", o._dtchild["Data_Type"].ToString());
                            o._dtmater.SetField("Character_Maximum_Length1", o._dtchild["Character_Maximum_Length"].ToString());
                            o._dtmater.SetField("Character_Octet_Length1", o._dtchild["Character_Octet_Length"].ToString());
                            o._dtmater.SetField("Date_Time_Precision1", o._dtchild["Date_Time_Precision"].ToString());
                            o._dtmater.SetField("Default1", o._dtchild["Default"].ToString());
                            o._dtmater.SetField("Is_Nullable1", o._dtchild["Is_Nullable"].ToString());
                            o._dtmater.SetField("Is_Primary1", o._dtchild["Is_Primary"].ToString());
                            o._dtmater.SetField("Numeric_Precision1", o._dtchild["Numeric_Precision"].ToString());
                            o._dtmater.SetField("Numeric_Precision_Radix1", o._dtchild["Numeric_Precision_Radix"].ToString());
                            o._dtmater.SetField("Numeric_Scale1", o._dtchild["Numeric_Scale"].ToString());
                            o._dtmater.SetField("Ordinal_Position1", o._dtchild["Ordinal_Position"].ToString());
                            o._dtmater.SetField("Description1", o._dtchild["Description"].ToString());

                        }
                    );

            }
            catch (Exception e) { Debug.WriteLine(e.Message); }


            WriteToExcel(dtRelation.Copy());
            Console.WriteLine("Done.");
        }

        private void btnDocumentFlow_Click(object sender, EventArgs e)
        {
            DocumentDataFlowInExcel((long)cboFromDB.SelectedValue, (long)cboToDB.SelectedValue);
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
