using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using System.Data;
using System.Security.Cryptography;
using System.Windows.Forms;
using VA = VisioAutomation;
using Visio = Microsoft.Office.Interop.Visio;
using SP = Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newton = Newtonsoft.Json;

namespace ICAS_Architect
{

    public class JsonSPClass
    {
        public string property1 { get; set; }
        public string property2 { get; set; }
        public List<Dictionary<int, string>> property3 { get; set; }
    }

    internal class SharepointManager
    {
        internal HttpDownloadClient httpDownloadClient = null;
        string SPRepository = null;
        Visio.Application vApplication = null;
        ClientContext gclientContext = null;

        // These are used to load data into the Visio External Data table. 
        
        private Visio.DataRecordset applicationRecordset = null;
        private Visio.DataRecordset databaseRecordset = null;
        private Visio.DataRecordset tableRecordset = null;
        private Visio.DataRecordset columnRecordset = null;
        private Visio.DataRecordset relationRecordset = null;

        // These are our DataTables, where we can store updates.
        private DataTable _dtApplications = null;
        private DataTable _dtDatabases = null;
        private DataTable _dtTables = null;
        private DataTable _dtColumns = null;
        //private DataTable _dtRelations = null;


        internal SharepointManager()
        {
            vApplication = Globals.ThisAddIn.Application;
        }

      

        internal ClientContext GetClientContext(bool WarnIfTablesDontExist=true, bool AskForSharepointLocation=false, bool Refresh = false)
        {
            string SPRepositoryFromRegistry = Globals.ThisAddIn.registryKey.GetValue("SharepointName", "https://icas1854.sharepoint.com/sites/Architecture").ToString();
            SPRepository = SPRepositoryFromRegistry;
            if (AskForSharepointLocation)
            {
                string ignore = "-1";
                DialogResult result = Utilites.ScreenEvents.ShowInputDialog(ref SPRepository, ref ignore, "Site", "ignore", "Enter SharePoint site URL");
                if (result != DialogResult.OK) return null;

                if (SPRepository != SPRepositoryFromRegistry)
                    Globals.ThisAddIn.registryKey.SetValue("SharepointName", SPRepository);
            }

            if (httpDownloadClient == null | Refresh == true)
            {
                //GG: Toto = strip everything after the .com
                httpDownloadClient = new HttpDownloadClient("https://icas1854.sharepoint.com");
                httpDownloadClient.Connect("https://icas1854.sharepoint.com/_api/web/lists");
            }

            //httpDownloadClient.client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json;odata=verbose"));

            OfficeDevPnP.Core.AuthenticationManager authMgr = new OfficeDevPnP.Core.AuthenticationManager();

            //            string siteUrl = SPRepository;

            gclientContext= authMgr.GetAzureADAccessTokenAuthenticatedContext(SPRepository, httpDownloadClient.accessToken);

            if (WarnIfTablesDontExist) {
 //               gclientContext = this.GetSPAccessToken();
                Web web = gclientContext.Web;
                gclientContext.Load(web);
                gclientContext.Load(web.Lists);
                gclientContext.ExecuteQueryRetry();

                if (!gclientContext.Web.ListExists("Tables") | !gclientContext.Web.ListExists("Columns") | !gclientContext.Web.ListExists("Relations"))
                {
                    System.Windows.Forms.MessageBox.Show("This Sharepoint site does not have the required lists.\nPlease see your administrator for the correct repository.", "Missing Repository");
                    return null;
                }
            }

            return (gclientContext);
        }


        internal void getAllTableDataRecordsets(string appName = "All", string dBName = "All", bool includeViews = true, bool includeAPIs = true)
        {
            try
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                getTableDataRecordset();
                vApplication.ActiveWindow.Windows.ItemFromID[2044].Visible = true;
                Utilites.ScreenEvents.DoEvents();

                getColumnDataRecordset();
                Utilites.ScreenEvents.DoEvents();

                getRelationDataRecordset();
                Utilites.ScreenEvents.DoEvents();

                getApplicationDataRecordset();
                Utilites.ScreenEvents.DoEvents();

                getDatabaseDataRecordset();
                Utilites.ScreenEvents.DoEvents();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }



        internal string CreateSharepointConnectionString(string listName, string viewName = "")
        {
            if (gclientContext is null) gclientContext = this.GetClientContext();
            List list = gclientContext.Web.GetListByTitle(listName);
            if (viewName == "")
                return "PROVIDER=WSS;DATABASE=" + SPRepository + "; LIST={" + list.Id + "};";
            else
            {
                string ViewID = list.GetViewByName(viewName).Id.ToString();
                return "PROVIDER=WSS;DATABASE=" + SPRepository + "; LIST={" + list.Id + "};VIEW={" + ViewID + "};";
            }
        }

        internal void DeleteDataRecordset(string listName)
        {
            for (int i = 1; i <= vApplication.ActiveDocument.DataRecordsets.Count; i++)
                if (vApplication.ActiveDocument.DataRecordsets[i].Name == "Tables")
                    vApplication.ActiveDocument.DataRecordsets[i--].Delete();
        }

        internal async Task<DataTable> retrieveDataTable(string appUrl, int iteration, string[] PrimaryKey)
        {
            const int iterSize = 1000;
            var client = httpDownloadClient.client;
            client.DefaultRequestHeaders.Accept.Clear();
            System.Net.Http.Headers.MediaTypeWithQualityHeaderValue acceptHeader = System.Net.Http.Headers.MediaTypeWithQualityHeaderValue.Parse("application/json;odata.metadata=none");
            client.DefaultRequestHeaders.Accept.Add(acceptHeader);

            string pageURL = appUrl + "&$skiptoken=Paged=TRUE%26p_SortBehavior=0%26p_ID=" + (iteration * iterSize) + "&$top=" + iterSize;// "&$skiptoken=Paged=TRUE";// + " &$top=" + iterSize + "&$skip=" + (iteration * iterSize).ToString();

            var response = await client.GetAsync(pageURL);
            var responseText = await response.Content.ReadAsStringAsync();
            try { response.EnsureSuccessStatusCode(); }
            catch(Exception e)
            {
                Console.WriteLine(e.Message + "\n" + response);
                return null;
            }

            var schemaObject = JObject.Parse(responseText);
            var values = schemaObject["value"].ToArray();
            if (values.Length == 0) return null;

            //GG: I hate serializing only to deserialize, but it seems the quickest way to put everything into a datatable
            DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(values));
            if (values.Length == iterSize)
            {   // if we retrieved iterSize records then recursively call retrieveDataTable until we get ALL of our records
                DataTable tmpDT = await retrieveDataTable(appUrl, iteration + 1, null);
                if (tmpDT != null)
                    dataTable.Merge(tmpDT);
            }

            if (PrimaryKey == null) return dataTable;

            var keys = new DataColumn[PrimaryKey.Length];
            for (int i = 0; i < PrimaryKey.Length; i++)
                keys[i] = dataTable.Columns[PrimaryKey[i]];

//            dataTable.PrimaryKey = keys;

            return dataTable;
        }


        /******************************************************************************************************************************************
         *  Usage:
         *  This function will retrieve all records from the Tables list in Sharepoint. If you are wanting to return everything, the Primary Key
         *  should be set as "Title", which is the GUID/Hash we created for that record. If you are returning only for one database, you can use
         *  "Table_Name" as it will be unique.
         * ******************************************************************************************************************************************/

        internal async Task<DataTable> GetApplicationsFromSharepoint(string AppID, string[] fieldList, string[] PrimaryKey)
        {
            string selectString = (fieldList.Length > 0) ? "$select=" : "";
            string filterString = (AppID != "") ? "&$filter=Id eq '" + AppID + "'" : "";

            for (int i = 0; i < fieldList.Length; i++)
                selectString = selectString + (i > 0 ? "," : "") + (fieldList[i]);

            DataTable tmpdt= await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Applications')/items?" + selectString + "&" + filterString, 0, PrimaryKey);
            tmpdt.Columns["Title"].ColumnName = "Application_Name";
            return tmpdt;
        }

        internal async Task<DataTable> GetDBsFromSharepoint(string AppID, string[] fieldList, string[] PrimaryKey)
        {
            // Retrieve the parent list (Applications)
            DataTable dbApp = _dtApplications;
            if (dbApp is null) dbApp = await GetApplicationsFromSharepoint("", new string[] { }, new string[] { "Id" });
            _dtApplications = dbApp;

            string selectString = (fieldList.Length > 0) ? "$select=" : "";
            string filterString = (AppID != "") ? "$filter=Application_NameId eq '" + AppID + "'" : "";
            string orderbyString = "&$orderby=Application_Name,Title";

            for (int i = 0; i < fieldList.Length; i++)
                selectString = selectString + (i > 0 ? "," : "") + (fieldList[i]);

            DataTable dbDT = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Databases')/items?" + selectString + filterString + orderbyString, 0, PrimaryKey);
            dbDT.Columns["Title"].ColumnName = "Database_Name";

            dbDT.Columns.Add("Application_Name", typeof(string));

            dbDT.AsEnumerable().Join(dbApp.AsEnumerable(),
                _dtmater => Convert.ToString(_dtmater["Application_NameId"]),
                _dtchild => Convert.ToString(_dtchild["id"]),
                (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                    o => o._dtmater.SetField("Application_Name", o._dtchild["Application_Name"].ToString())
                );
            dbDT.Columns.Add("AppAndDB", typeof(string), "Application_Name + ' - ' + Database_Name");
            
            return dbDT;
        }

        // Returns only the TableRelations, nothing else
        internal async Task<DataTable> GetTableRelationsFromSharepoint(string appId, long dBId, string Connection_Type, string[] fieldList, string[] PrimaryKey)
        {
            string selectString = "$select=Connection_Type eq '" + Connection_Type + "'";
            string filterString = (dBId > 0) ? "$filter=Database_NameId eq '" + dBId.ToString() + "'" : "";

            for (int i = 0; i < fieldList.Length; i++)
                selectString = selectString + (i > 0 ? "," : "") + (fieldList[i]);

            DataTable dataTable = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Relations')/items?" + selectString + "&" + filterString, 0, PrimaryKey);
            dataTable.Columns["Title"].ColumnName = "Relation_Name";

            if (dataTable is null) return null;

            return dataTable;
        }

            internal async Task<DataTable> GetTablesFromSharepoint(string appId, long dBId, string[] fieldList, string[] PrimaryKey)
        {
            // Retrieve the parent list (Databases)
            DataTable dtDb = _dtDatabases;
            if (dtDb is null) dtDb = await GetDBsFromSharepoint("", new string[] { }, new string[] { "Id" });
            _dtDatabases = dtDb; 

            string selectString = (fieldList.Length > 0) ? "$select=" : "";
            string filterString = (dBId > 0) ? "$filter=Database_NameId eq '" + dBId.ToString() + "'" : "";

            for (int i = 0; i < fieldList.Length; i++)
                selectString = selectString + (i > 0 ? "," : "") + (fieldList[i]);

            DataTable dataTable = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Tables')/items?" + selectString + "&" + filterString, 0, PrimaryKey);
            if (dataTable is null) return null;

//            dataTable.Columns["Title"].ColumnName = "Table_Name";
            dataTable.Columns.Add("Database_Name", Type.GetType("System.String"));
            dataTable.Columns.Add("Application_NameId", Type.GetType("System.Int64"));
            dataTable.Columns.Add("Application_Name", Type.GetType("System.String"));

            try
            {
                dataTable.AsEnumerable().Join(dtDb.AsEnumerable(),  _dtmater => Convert.ToString(_dtmater["Database_NameId"]),   _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => {
                            o._dtmater.SetField("Database_Name", o._dtchild["Database_Name"].ToString());
                            o._dtmater.SetField("Application_NameId", (long)o._dtchild["Application_NameId"]);
                            o._dtmater.SetField("Application_Name", o._dtchild["Application_Name"].ToString());
                            }
                    ) ;
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return dataTable;
        }

        internal DataTable GetColumnsFromSharepointWrapper(long TableOneID)
        {
            // if the table exists locally, return the data table
            if (_dtColumns.Select("Table_NameId=" + TableOneID.ToString()).Length > 0) 
                return _dtColumns;

            // Otherwise look up the table and add it to the local cache
            DataTable tbl = Task.Run(async () => await GetColumnsFromSharepoint(0, TableOneID, new string[] {  }, new string[] { "Table_Name", "Title" })).Result;
            if (tbl != null)
                if (_dtColumns is null)
                    _dtColumns = tbl;
                else
                    _dtColumns.Merge(tbl);

            return _dtColumns;
        }


        // Get all columns by DB is not working yet. For our huge initial uploads it needs fixing.
        internal async Task<DataTable> GetColumnsFromSharepoint(long dbID, long tableID, string[] fieldList, string[] PrimaryKey)
        {
            // Retrieve the parent list (Tables)
            if(_dtTables == null)
                _dtTables = await GetTablesFromSharepoint("", dbID ,  new string[] { }, new string[] { "Database_NameId", "Title" });
            else if (_dtTables.Select("ID=" + tableID).Length == 0)
            {
                var dtTables = await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" });
                _dtTables.Merge(dtTables);
            }

            string selectString = (fieldList.Length > 0) ? "$select=" : "";
            for (int i = 0; i < fieldList.Length; i++)
                selectString += (i > 0 ? "," : "") + (fieldList[i]);

            DataTable dtColumn = null;
            try
            {
                if (tableID != 0)
                     dtColumn = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Columns')/items?" + selectString + "&$filter=Table_NameId eq '" + tableID + "'&$orderby=Title", 0, PrimaryKey);
            }catch(Exception e)
            { Console.WriteLine(e.Message); }
            
            if (dtColumn is null) return null;

            dtColumn.Columns["Title"].ColumnName = "Column_Name";
            dtColumn.Columns.Add("Table_Name", Type.GetType("System.String"));
            dtColumn.Columns.Add("Database_NameId", Type.GetType("System.Int64"));
            dtColumn.Columns.Add("Database_Name", Type.GetType("System.String"));

            try
            {
                dtColumn.AsEnumerable().Join(_dtTables.AsEnumerable(),   _dtmater => Convert.ToString(_dtmater["Table_NameId"]),   _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => {
                            o._dtmater.SetField("Table_Name", o._dtchild["Table_Name"].ToString());
                            o._dtmater.SetField("Database_NameId", (long)o._dtchild["Database_NameId"]);
                            o._dtmater.SetField("Database_Name", o._dtchild["Database_Name"].ToString());
                        }
                    );
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return dtColumn;
        }



        internal Visio.DataRecordset getTableDataRecordset(string appName = "All", string dBName = "All", bool includeViews = true, bool includeAPIs = true)
        {
            string whereCond = null;
            if (appName != "All")
                whereCond += (whereCond == null ? " WHERE " : " AND ") + "[Application] = '" + appName + "'";
            if (dBName != "All")
                whereCond += (whereCond == null ? " WHERE " : " AND ") + "[Database] = '" + dBName + "'";
            if (!includeViews)
                whereCond += whereCond == null ? " WHERE " : " AND " + "[Table Type] <> 'View'";
            if (!includeAPIs)
                whereCond += whereCond == null ? " WHERE " : " AND " + "[Table Type] <> 'API'";

            whereCond = "Select * from [Tables (All Items)] " + whereCond + ";";

            DeleteDataRecordset("Tables");
            string connString = CreateSharepointConnectionString("Tables", "All Items");
            tableRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, whereCond, 0, "Tables");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return tableRecordset;
        }

        internal Visio.DataRecordset getColumnDataRecordsetByGUID(string tableGUID)
        {
            string whereCond = " WHERE [TableUniqueID] = '" + tableGUID + "'";

            DeleteDataRecordset("Columns");
            string connString = CreateSharepointConnectionString("Columns");
            columnRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Columns] " + whereCond, 0, "Columns");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return columnRecordset;
        }

        internal Visio.DataRecordset getColumnDataRecordset(string tableName = "All", string dBName = "All")
        {
            string whereCond = null;
            if (tableName != "All")
            {
                whereCond = whereCond == null ? " WHERE " : " AND ";
                whereCond = whereCond + "[Table_Name] = '" + tableName + "'";
            }
            if (dBName != "All")
            {
                whereCond = whereCond == null ? " WHERE " : " AND ";
                whereCond = whereCond + "[Database] = '" + dBName + "'";
            }

            DeleteDataRecordset("Columns");
            string connString = CreateSharepointConnectionString("Columns");
            columnRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Columns] " + whereCond, 0, "Columns");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return columnRecordset;
        }

        internal Visio.DataRecordset getRelationDataRecordset(string tableName = "All", string dBName = "All")
        {
            string whereCond = null;
            if (tableName != "All")
            {
                whereCond = whereCond == null ? " WHERE " : " AND ";
                whereCond = whereCond + "([TableOne]='" + tableName + "' | [TableMany]='" + tableName + "')";
            }

            DeleteDataRecordset("Relations");
            string connString = CreateSharepointConnectionString("Relations");
            relationRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Relations];", 0, "Relations");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return relationRecordset;
        }

        internal Visio.DataRecordset getApplicationDataRecordset()
        {
            //string whereCond = null;
            DeleteDataRecordset("Applications");
            string connString = CreateSharepointConnectionString("Application");
            applicationRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Application];", 0, "Applications");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return applicationRecordset;
        }


        internal Visio.DataRecordset getDatabaseDataRecordset(string tableName = "All", string dBName = "All")
        {
            DeleteDataRecordset("Databases");
            string connString = CreateSharepointConnectionString("Database");
            databaseRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Database];", 0, "Databases");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return databaseRecordset;
        }



        internal void CreateSharepointLists()
        {
            using (var clientContext = GetClientContext(false))      //.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password))
            {
                Web web = clientContext.Web;
                clientContext.Load(web);
                clientContext.Load(web.Lists);
                clientContext.ExecuteQueryRetry();

                /*                List colList = web.Lists.GetByTitle("Columns");
                                clientContext.Load(colList);
                                clientContext.ExecuteQuery();

                                Field colField = colList.Fields.GetByInternalNameOrTitle("Table_NameId");
                                colField.EnableIndex();
                                colField.Update();
                                clientContext.ExecuteQuery();*/


                if (!clientContext.Web.ListExists("Applications"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Applications", true, true, string.Empty, true);
                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.Title = "Application_Name";
                    oField.StaticName = "Application_Name";
                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Update();

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Business_Unit' Name='Business_Unit'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='Cost' Name='Cost'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Description' Name='Description'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='DateTime' DisplayName='Go_Live' Name='Go_Live'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Health_Indicator' Name='Health_Indicator'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='IT_Unit' Name='IT_Unit'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Lifecycle' Name='Lifecycle'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Security_Classification' Name='Security_Classification'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Strategic_Importance' Name='Strategic_Importance'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='DateTime' DisplayName='Sunset_Date' Name='Sunset_Date'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Technical_Completeness' Name='Technical_Completeness'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Vendor' Name='Vendor'/>", true, AddFieldOptions.AddToDefaultContentType);
                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }


                if (!clientContext.Web.ListExists("Databases"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Databases", true, true, string.Empty, true);
                    List refList = web.Lists.GetByTitle("Applications");
                    clientContext.Load(refList);
                    clientContext.ExecuteQuery();

                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.Title = "Database_Name";
                    oField.StaticName = "Database_Name";
                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Update();

                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Application_Name' StaticName='Application_Name' DisplayName='Application_Name' List='" + refList.Id + "' ShowField = 'Title' RelationshipDeleteBehaviorType='Restrict' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='ApplicationID' Name='ApplicationID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Application' Name='Application'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='DBMS_Type' Name='DBMS_Type'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Version' Name='Version'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ServerID' Name='ServerID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Server_Name' Name='Server_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ExternalID' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityListURL' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityDefinitionsPath' Name='EntityDefinitionsPath'/>", true, AddFieldOptions.AddToDefaultContentType);
                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }


                if (!clientContext.Web.ListExists("Tables"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Tables", true, true, string.Empty, true);
                    List refList = web.Lists.GetByTitle("Databases");
                    clientContext.Load(refList);
                    clientContext.ExecuteQuery();

                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.Title = "Full_Table_Name";
                    oField.StaticName = "Full_Table_Name";
                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Update();
                    var fielda = Guid.NewGuid().ToString(); 
                    var fieldb = Guid.NewGuid().ToString(); 

//                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Table_Name' Name='Table_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Database_Name' StaticName='Database_Name' DisplayName='Database_Name' List='" + refList.Id + "' ShowField = 'Title' RelationshipDeleteBehaviorType='Restrict' Indexed='TRUE' Required='TRUE'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Table_Name' Name='Table_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Table_Type' Name='Table_Type'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='TableGUID' Name='TableGUID' EnforceUniqueValues='TRUE' Indexed='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Description' Name='Description'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Application' Name='Application'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Display_Name' Name='Display_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Schema' Name='Schema'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Set_Name' Name='Set_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ExternalID' Name='ExternalID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityListURL' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityDefinitionsPath' Name='EntityDefinitionsPath'/>", true, AddFieldOptions.AddToDefaultContentType);
//                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='ApplicationID' Name='ApplicationID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='DatabaseID' Name='DatabaseID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }


                if (!clientContext.Web.ListExists("Columns"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Columns", false, true, string.Empty, true);

                    List refList = web.Lists.GetByTitle("Tables");
                    clientContext.Load(refList);
                    clientContext.ExecuteQuery();

                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.Title = "Column_Name";
                    oField.StaticName = "Column_Name";
                    oField.EnableIndex();
                    //oField.EnforceUniqueValues = true;
                    oField.Update();
                    var fielda = Guid.NewGuid().ToString(); //GenerateHash("My silly first table reference");
                    var fieldb = Guid.NewGuid().ToString(); //GenerateHash("My silly second table reference");

                    //string schemaLookupField = " ";
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_Name' StaticName='Table_Name' DisplayName='Table_Name' List='" + refList.Id + "'  ID='{" + fielda + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Description' Name='Description'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Ordinal_Position' Name='Ordinal_Position'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Default' Name='Default'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Boolean' DisplayName='Is_Nullable' Name='Is_Nullable'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Boolean' DisplayName='Is_Primary' Name='Is_Primary'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Data_Type' Name='Data_Type'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='Character_Maximum_Length' Name='Character_Maximum_Length'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='Character_Octet_Length' Name='Character_Octet_Length'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='Numeric_Precision' Name='Numeric_Precision'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='Numeric_Precision_Radix' Name='Numeric_Precision_Radix'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='Numeric_Scale' Name='Mumeric_Scale'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Integer' DisplayName='Date_Time_Precision' Name='Date_Time_Precision'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Character_Set_Catalog' Name='Character_Set_Catalog'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Character_Set_Schema' Name='Character_Set_Schema'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Character_Set_Name' Name='Character_Set_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Collation_Catalog' Name='Collation_Catalog'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Collation_Schema' Name='Collation_Schema'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Collation_Name' Name='Collation_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ExternalID' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ColumnGUID' Name='ColumnGUID' EnforceUniqueValues='TRUE' Indexed='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);

                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }


                if (!clientContext.Web.ListExists("Relations"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Relations", false, true, string.Empty, true);
                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.Title = "RelationName";
                    oField.StaticName = "RelationName";
                    oField.EnforceUniqueValues = false;
                    oField.Required = false;
                    oField.Update();

                    List refList = web.Lists.GetByTitle("Tables");
                    List refListCol = web.Lists.GetByTitle("Columns");
                    clientContext.Load(refList);
                    clientContext.Load(refListCol);
                    clientContext.ExecuteQuery();

                    // should I make these constants?
                    var fielda = Guid.NewGuid().ToString();
                    var fieldb = Guid.NewGuid().ToString();
                    var fieldc = Guid.NewGuid().ToString();
                    var fieldd = Guid.NewGuid().ToString();
                    var fielde = Guid.NewGuid().ToString();
                    var fieldf = Guid.NewGuid().ToString();
                    var fieldg = Guid.NewGuid().ToString();
                    var fieldh = Guid.NewGuid().ToString();

                    /* dtRelation must contain the following case-sensitive fieldnames: ID, Relation_Name, Table_OneId, Table_One, Table_ManyId, Table_Many,
                     * Column_OneId, Column_ManyId, Column_ManyId, Column_Many, Relation_Type, Relation_Level_One, Relation_Level_Many */

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Relation_Description' Name='Relation_Description'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_One' StaticName='Table_One' DisplayName='Table_One' List='" + refList.Id + "'  ID='{" + fielda + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_Many' StaticName='Table_Many' DisplayName='Table_Many' List='" + refList.Id + "'  ID='{" + fieldc + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_One' StaticName='Column_One' DisplayName='Column_One' List='" + refListCol.Id + "'  ID='{" + fielde + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_Many' StaticName='Column_Many' DisplayName='Column_Many' List='" + refListCol.Id + "'  ID='{" + fieldg + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);

                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_OneId' StaticName='Table_OneId' DisplayName='Table_OneId' List='" + refList.Id + "'  ID='{" + fielda + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_One' StaticName='Table_One' DisplayName='Table_One' List='" + refList.Id + "' ID='{" + fieldb + "}' ShowField='Table_Name' FieldRef='{" + fielda + "}' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_ManyId' StaticName='Table_ManyId' DisplayName='Table_ManyId' List='" + refList.Id + "'  ID='{" + fieldc + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_Many' StaticName='Table_Many' DisplayName='Table_Many' List='" + refList.Id + "' ID='{" + fieldd + "}' ShowField='Table_Name' FieldRef='{" + fieldc + "}' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_OneId' StaticName='Column_OneId' DisplayName='Column_OneId' List='" + refListCol.Id + "'  ID='{" + fielde + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_One' StaticName='Column_One' DisplayName='Column_One' List='" + refListCol.Id + "' ID='{" + fieldf + "}' ShowField='Column_Name' FieldRef='{" + fielde + "}' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_ManyId' StaticName='Column_ManyId' DisplayName='Column_ManyId' List='" + refListCol.Id + "'  ID='{" + fieldg + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_Many' StaticName='Column_Many' DisplayName='Column_Many' List='" + refListCol.Id + "' ID='{" + fieldh + "}' ShowField='Column_Name' FieldRef='{" + fieldg + "}' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ColumnOne' Name='ColumnOne'/>", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ColumnMany' Name='ColumnMany'/>", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='TableOne' Name='TableOne'/>", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='TableMany' Name='TableMany'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Relation_Type' Name='ConnectionType'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Relation_Level_One' Name='EntityOneType'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Relation_Level_Many' Name='EntityToType'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ExternalID' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Intersect_Entity' Name='Intersect_Entity'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Calculation' Name='Calculation'/>", true, AddFieldOptions.AddToDefaultContentType);
                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }
            }
        }


        internal void UploadToSharePoint()
        {
            UploadCompleteDataRecordsetToSharepoint();
        }

        internal void UploadTableChanges()
        {
            return;
        }


        internal void UploadERInformation(ERInformation eRInformation)
        {
            List targetList = null;
            int i = 0;
            int j = 0;

            var clientContext =GetClientContext();

            i = 0;
            j=0;
            targetList = clientContext.Web.Lists.GetByTitle("Columns");
            foreach (EREntityAttribute eREntityAttribute in eRInformation.EREntitieAttributes)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oItem = targetList.AddItem(itemCreateInfo);
                String TableName = eRInformation.FindParentTableID(eREntityAttribute.EntityLogicalName);
                oItem["Title"] = eREntityAttribute.MetadataId;
                oItem["Column"] = eREntityAttribute.AttributeName;
                oItem["Column_x0020_Type"] = eREntityAttribute.AttributeType;
                oItem["Table_Name"] = TableName;
                oItem["Description"] = eREntityAttribute.Description;
                oItem["Is_x0020_Primary"] = eREntityAttribute.IsPrimaryID;
                oItem["Position"] = eREntityAttribute.ColumnNumber;
                oItem["Default"] = "";
                oItem["Nullable"] = false;
                oItem["Data_x0020_Type"] = eREntityAttribute.DataType;
                oItem.Update();


                j++;
                if (++i >= 50)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + j + " of " + eRInformation.EREntitieAttributes.Count().ToString());
                    clientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                clientContext.ExecuteQuery();


            // Include everything that is a relation as well as the entities that get related and any fields that get related
            // Sometimes only entities get related, so this is the left join.
            var r1 = from rel in eRInformation.ERRelations
                        join tOne in eRInformation.EREntities
                        on rel.EntityOne equals tOne.entityLogicalName into tOneMap
                        join tMany in eRInformation.EREntities
                            on rel.EntityMany equals tMany.entityLogicalName into tManyMap
                        join aOne in eRInformation.EREntitieAttributes
                            on new { ent = rel.EntityOne, col = rel.AttributeMany } equals new { ent = aOne.EntityName, col = aOne.AttributeName } into aOneMap
                        join aMany in eRInformation.EREntitieAttributes
                            on new { ent = rel.EntityOne, col = rel.AttributeMany } equals new { ent = aMany.EntityName, col = aMany.AttributeName } into aManyMap
                        from tOne in tOneMap.DefaultIfEmpty()
                        from tMany in tManyMap.DefaultIfEmpty()
                        from aOne in aOneMap.DefaultIfEmpty()
                        from aMany in aManyMap.DefaultIfEmpty()
                        select new
                        {
                            rel.IntersectEntity, 
                            GUIDOne =  tOne==null ? null : tOne.metadataId,         // If there is no field return blank 
                            GUIDMany = tMany == null ? null : tMany.metadataId, 
                            FieldOne = aOne==null ? null : aOne.MetadataId, 
                            FieldMany = aMany==null ? null : aMany.MetadataId
                        };

            i = 0;
            j = 0;
            targetList = clientContext.Web.Lists.GetByTitle("Relations");
            //            foreach (ERRelation eRRelation in eRInformation.ERRelations)
            foreach (var rel in r1)
            {

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oItem = targetList.AddItem(itemCreateInfo);

                var guid = Guid.NewGuid().ToString();
                oItem["Title"] = guid;
                oItem["Relation_x0020_Name"] = rel.IntersectEntity;
                oItem["Primary_x0020_Column"] = rel.FieldOne;
                oItem["Foreign_x0020_Column"] = rel.FieldMany;
                oItem["Primary_x0020_Table"] = rel.GUIDOne;
                oItem["Foreign_x0020_Table"] = rel.GUIDMany;
                oItem["Connection_x0020_Type"] = "";
                oItem.Update();

                j++;
                if (++i >= 50)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + j + " of " + r1.Count().ToString());
                    clientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                clientContext.ExecuteQuery();

            getAllTableDataRecordsets("All", "All", true, true);
        }

        MD5 md5 = null;

        internal string GenerateHash(string entityName)
        {
            if (md5 == null)
                md5 = MD5.Create();

            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(entityName.ToLower());
            byte[] hashBytes = md5.ComputeHash(inputBytes);

            // Convert the byte array to hexadecimal string
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hashBytes.Length; i++)
            {
                sb.Append(hashBytes[i].ToString("X2"));
            }
            return sb.ToString();
        }



        internal Visio.DataRecordset getTableFromName(string tblName)
        {
            try
            {
                for (int i = 1; i <= vApplication.ActiveDocument.DataRecordsets.Count; i++)
                    if (vApplication.ActiveDocument.DataRecordsets[i].Name == tblName)
                        return vApplication.ActiveDocument.DataRecordsets[i];
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return null;
        }


        internal bool AssignDataRecordsets()
        {
            applicationRecordset = getTableFromName("Applications");
            databaseRecordset = getTableFromName("Databases");
            tableRecordset = getTableFromName("Tables");
            columnRecordset = getTableFromName("Columns");
            relationRecordset = getTableFromName("Relations");
            if (tableRecordset is null | columnRecordset is null | relationRecordset is null)
            {
                MessageBox.Show("We need Tables, Columns, and Relations to upload to SharePoint.");
                return false;
            }
            return true;
        }


        internal int GetDRColumnNumber(System.Array record, string fieldName)
        {
            for (int i = 0; i < record.Length; i++)
                if (record.GetValue(i).ToString().ToLower() == fieldName.ToLower())
                    return i;
            return -1;
        }


        private string LookupFromColumnDataset(string SearchForColumnName, string ColumnToReturn)
        {
            var recs = columnRecordset.GetDataRowIDs("Column_Name = '" + SearchForColumnName + "'");
            if (recs.GetUpperBound(0) == -1) return "";
            var rec = columnRecordset.GetRowData((int)recs.GetValue(0));
            var ID = GetDRColumnNumber(columnRecordset.GetRowData(0), ColumnToReturn);
            return rec.GetValue(ID).ToString();
        }

        private string LookupFromTableDataset(string SearchForTable, string ColumnName)
        {
            var recs = tableRecordset.GetDataRowIDs("Table_Name = '" + SearchForTable + "'");
            if (recs.GetUpperBound(0) == -1) return "";
            var rec = tableRecordset.GetRowData((int)recs.GetValue(0));
            var ID = GetDRColumnNumber(tableRecordset.GetRowData(0), ColumnName);
            return rec.GetValue(ID).ToString();
        }


        internal bool isValidColumn(string columnList, string columnName)
        {
            // Some entities (such as SystemUser entity) are not useful for ER diagram. Skip those noisy not useful entities for visualization.
            return (columnList.Contains($",{columnName.ToLower()},"));
        }


        internal async Task<string> GetSharePointListFields(string listName)
        {
            var client = httpDownloadClient.client;
            client.DefaultRequestHeaders.Accept.Clear();
            System.Net.Http.Headers.MediaTypeWithQualityHeaderValue acceptHeader = System.Net.Http.Headers.MediaTypeWithQualityHeaderValue.Parse("application/json;odata.metadata=none");
            client.DefaultRequestHeaders.Accept.Add(acceptHeader);

            var response = await client.GetAsync(SPRepository + "/_api/lists/getbytitle('" + listName + "')//fields?$filter=Hidden eq false and ReadOnlyField eq false ");
            var responseText = await response.Content.ReadAsStringAsync();
            try { response.EnsureSuccessStatusCode(); }
            catch (Exception e)
            {   MessageBox.Show("Cannot find Sharepoint List " + listName + "\n" + e.Message + "\n" + response);
                return null;            }

            var schemaObject = JObject.Parse(responseText);
            var values = schemaObject["value"].ToArray();

            string columnList = ",";
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i].SelectToken("Title").ToString() != "ID")
                    columnList = columnList + values[i].SelectToken("Title").ToString().ToLower() + ",";
            }

            //List targetList = gclientContext.Web.Lists.GetByTitle(listName);
            //FieldCollection oFieldCollection = targetList.Fields;

            //// Here we have loaded Field collection including title of each field in the collection
            //gclientContext.Load(oFieldCollection, oFields => oFields.Include(field => field.Title));
            //gclientContext.ExecuteQuery();

            //foreach (Field oField in oFieldCollection)
            //    if(oField.Title != "ID")
            //        columnList = columnList + oField.Title.ToLower() + ",";
            //            columnList = ",Title,Column_Name,Description,Ordinal_Position,Data_Type,ExternalID,TableGUID,TableID,".ToLower();
            return columnList;
        }


        internal DataTable convertDataRecordsetToDataTable(Visio.DataRecordset DR)
        {
            // convert string to stream
            byte[] byteArray = Encoding.UTF8.GetBytes(System.Text.RegularExpressions.Regex.Replace(DR.DataAsXML.Replace("\r", "").Replace("\n", "").Replace("rs:", "").Replace("s:", "").Replace("dt:", "").Replace("z:", "").Replace("\t", "").Replace("name='_Visio_RowID_'", ""), "[<]xml.*#RowsetSchema'[>]", "<xml>"));
            System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);

            //dataTable.Columns.Add("c0");            //for (int i = 0; i < DR.GetRowData(0).Length; i++)            //    dataTable.Columns.Add(DR.GetRowData(0).GetValue(i).ToString());

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(stream, System.Data.XmlReadMode.Auto);
            DataTable dataTable = dataSet.Tables["row"];

            return dataTable;
        }


        /* SaveTablesToSharepoint requires the following case-sensitive fieldnames for an upload:
            * ID (null for new ones), Table_Name, Database_NameId, Database_Name,
            * Table_Type (Table, View, API)
            * Every other column will be uploaded into Sharepoint if it matches the Sharepoint field name or ignored
            * TestForExistence not yet implemented
            */
        internal void saveTablesToSharepoint(DataTable dTable, bool TestForExistence = false)
        {
            var targetList = gclientContext.Web.Lists.GetByTitle("Tables");
            string sharepointTableFields = Task.Run(async () => await GetSharePointListFields("Tables")).Result;
            int i = 0, j = 0;
            ListItem oItem = null;

            // Get Database Info to reference
            if (_dtDatabases is null) 
                _dtDatabases = Task.Run(async () => await GetDBsFromSharepoint("", new string[] { }, new string[] { "Title" })).Result;

            foreach (DataRow row in dTable.Rows)
            {
                j++;
                if (row["Edit_Status"].ToString() != "New") continue;

                // Add if there's not ID, otherwise edit.
                if (row.Field<long?>("ID") == null) 
                    oItem = targetList.AddItem(new ListItemCreationInformation());
                else
                    oItem = targetList.GetItemById(row.Field<int>("ID"));

                // set defaults 
                oItem["Table_Type"] = "API";
                foreach ( DataColumn column in dTable.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_"); 
                    if (isValidColumn(sharepointTableFields, fieldName))
                        oItem[fieldName] = row.Field<string>(column.ColumnName);
                }

                if(row.Field<long?>("ID") is null)
                { // if we haven't added this before, we need to look up the app name and create the GUID
                    var dbID = row.Field<long>("Database_NameId");
                    DataRow[] dRow;
                    if (dbID == 0) // if no DBID was passed in, find it from the database name.
                        dRow = _dtDatabases.Select("Database_Name = '" + row.Field<string>("Database_Name") + "'");
                    else
                        dRow = _dtDatabases.Select("ID='" + dbID + "'");

                    //dbID = (long) dRow[0]["Id"];
                    oItem["Database_Name"] = (long)dRow[0]["Id"]; // This is a sharepoint weirdism. It links to ID rather than dbrow.
                    oItem["TableGUID"] = GenerateHash(dRow[0].Field<string>("Application_Name") + "." + dRow[0].Field<string>("Database_Name") + "." + row.Field<string>("Table_Name"));
                }

                oItem["Title"] = row.Field<string>("Database_Name") + "." + row.Field<string>("Table_Name"); // Title is "Full_Table_Name
                oItem["Table_Name"] = row.Field<string>("Table_Name");

                oItem.Update();
                row["Edit_Status"] = "Uploading";

                if (i++ >= 19)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + j + " of " + dTable.Rows.Count);
                    gclientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                gclientContext.ExecuteQuery();
        }


        /* dtColumns must contain the following case-sensitive fieldnames: ID, Column_Name (Not Null), non-null Table_NameId or Table_Name, Column_Type
         * Table_NameId or Table_Name must be of a valid table
         * Every other column will be uploaded into Sharepoint if it matches the Sharepoint field name or ignored
         * TestForExistence not yet implemented
        */
        internal bool saveColumnsToSharepoint(DataTable dtColumns, long dbID, bool TestForExistence = false)
        {
            //try
            //{
                DataTable tableDataTable = Task.Run(async () => await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Table_Name", "Database_NameId" })).Result;
                ListItem oItem = null;

                var targetList = gclientContext.Web.Lists.GetByTitle("Columns");
                string sharepointColumnFields = Task.Run(async () => await GetSharePointListFields("Columns")).Result;
                int i = 0, j = 0;

                foreach (DataRow row in dtColumns.Rows)
                {
                    j++;
                    if (row["Edit_Status"].ToString() != "New") continue;

                    // Add if there's no ID, otherwise edit.
                    if (row.Field<long?>("ID") == null)
                        oItem = targetList.AddItem(new ListItemCreationInformation());
                    else
                        oItem = targetList.GetItemById(row.Field<int>("ID"));

                    oItem["Title"] = "";
                    foreach (DataColumn column in dtColumns.Columns)
                    {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                        string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_");
                        if (column.ColumnName == "Column_Name") continue;
                        if (isValidColumn(sharepointColumnFields, fieldName))
                            oItem[fieldName] = row.Field<string>(column.ColumnName);
                    }

                    if (row.Field<long?>("ID") == null)
                    {
                        var dr = tableDataTable.Select("[Table_Name] = '" + row.Field<string>("Table_Name").ToString() + "'");
                        if (dr.Length == 0)
                        { MessageBox.Show("Cannot find table " + row.Field<string>("Table_Name").ToString() + ". \n Skipping insertion of " + row.Field<string>("Column_Name").ToString()); continue; }
                        var tableName = dr[0][tableDataTable.Columns["Table_Name"].Ordinal].ToString();
                        var dbName = dr[0][tableDataTable.Columns["Database_Name"].Ordinal].ToString();
                        var appName = dr[0][tableDataTable.Columns["Application_Name"].Ordinal].ToString();
                        var tableType = dr[0][tableDataTable.Columns["Table_Type"].Ordinal].ToString();

                        oItem["ColumnGUID"] = GenerateHash(appName + "." + dbName + "." + dr[0][tableDataTable.Columns["Table_Name"].Ordinal].ToString() + "." + row.Field<string>("Column_Name"));
                        //                    oItem["dbTable"] = tableDataTable.Rows.Find(tableName)["Id"];
                        //I'm going to try to set this with Id again, but sharepoint sems to want to use the title field
                        oItem["Title"] = row.Field<string>("Column_Name");
                        oItem["Table_Name"] = dr[0][tableDataTable.Columns["ID"].Ordinal].ToString();
                    }
                    oItem.Update();
                    row["Edit_Status"] = "Uploading";
                    if (i++ >= 49)
                    {
                        Utilites.ScreenEvents.DisplayVisioStatus("Uploading Columns: " + j + " of " + dtColumns.Rows.Count);
                        gclientContext.ExecuteQuery();
                        i = 0;
                    }
                }
                if (i > 0)
                    gclientContext.ExecuteQuery();
            //}
            //catch(Exception e)
            //{
            //    MessageBox.Show("Error: " + e.Message + "\n\n Retrying once");
            //    return false;
            //}

            return true;
        }

        internal long getColumnIDFromDatatable(string ColumnID, string ColumnName, string TableOneID)
        {
            // skip if we know the columnID or if there is no column or if we don't know the TableID
            if ((ColumnID ?? "") == "" & TableOneID  != "" & ColumnName  != "")
            {   
                if (_dtColumns != null)
                {   // if we already have the column cached we return the ID.
                    var ret = _dtColumns.Select("[Column_Name]='" + ColumnName + "'");
                    if (ret.Length > 0)
                        return (long)ret[0][_dtColumns.Columns["ID"].Ordinal];
                }

                DataTable tbl = Task.Run(async () => await GetColumnsFromSharepoint(0, long.Parse(TableOneID), new string[] { "Id", "Title, Table_NameId" }, new string[] { "Table_Name", "Title" })).Result;
                if (tbl != null)
                    if (_dtColumns is null) 
                        _dtColumns = tbl;
                    else 
                        _dtColumns.Merge(tbl);

                //try looking again.
                if (_dtColumns is null) return 0;
                var ret2 = _dtColumns.Select("[Column_Name]='" + ColumnName + "'");
                if (ret2.Length > 0)
                    return (long)ret2[0][_dtColumns.Columns["ID"].Ordinal];
            }
            return 0;
        }

                /* dtRelation must contain the following case-sensitive fieldnames: ID, Table_OneId, Table_One, Table_ManyId, Table_Many,
                 * Column_OneId, Column_ManyId, Column_ManyId, Column_Many, Relation_Type, Relation_Level_One, Relation_Level_Many
                 * Table_OneId must be a valid table or we have to look up using Table_One
                 * Table_ManyId must be a valid table or we have to look up using Table_Many
                 * Column_OneId can be empty. If it is not empty it must be valid or Column_One must be valid
                 * Column_ManyId can be empty. If it is not empty it must be valid or Column_Many must be valid
                 * Relation_Type can be null 
                 *          valid types can be "Foreign Key", "Data Flow", "Message", "Contains", 
                 * Relation_LevelOne and Relation_LevelMany 
                 *          valid types can be "Table", "Application", "Actor", 
                 *          When linking columns, set Relation_Levels to "Table" as we have added secondary fields for columns
                 * Every other field (Description, etc) will be uploaded into Sharepoint if it matches the Sharepoint field name or ignored
                 * 
                 * TestForExistence not yet implemented
                */
        internal void saveRelationsToSharepoint(DataTable dtRelation, long dbID, long dbManyID, bool TestForExistence = false)
        {
            DataTable tableDataTable = Task.Run(async () => await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Table_Name", "Database_NameId" })).Result;
            _dtTables = tableDataTable;
            if (dbID != dbManyID)
            {
                DataTable tbl = Task.Run(async () => await GetTablesFromSharepoint("", dbManyID, new string[] { "Id", "Table_Name" }, new string[] { "Table_Name", "Database_NameId" })).Result;
                tableDataTable.Merge(tbl);
            }

            var targetList = gclientContext.Web.Lists.GetByTitle("Relations");
            string sharepointRelationFields = Task.Run(async () => await GetSharePointListFields("Relations")).Result;
            ListItem oItem = null;
            int i = 0, j = 0;

            foreach (DataRow row in dtRelation.Rows)
            {
                j++;
                if (row["Edit_Status"].ToString() == "") continue;

                // Add if there's no ID, otherwise edit.
                if (row.Field<int?>("ID") == null) //change to isnull
                    oItem = targetList.AddItem(new ListItemCreationInformation());
                else
                    oItem = targetList.GetItemById(row.Field<int>("ID"));
                if(row["Relation_Type"].ToString() == "Foreign Key")
                    if (Globals.ThisAddIn.drawingManager.IsEntityInSkipList(row["Table_One"].ToString()))  continue; //we want to keep skip list columns but remove their relations

                var colOneID = getColumnIDFromDatatable(row["Column_OneID"].ToString(), row["Column_One"].ToString(), row["Table_OneID"].ToString());
                var colManyID = getColumnIDFromDatatable(row["Column_ManyID"].ToString(), row["Column_Many"].ToString(), row["Table_ManyID"].ToString());

                if (row["Table_OneId"].ToString() == "")
                {
                    if (colOneID > 0)
                    {
                        var ret2 = _dtColumns.Select("[ID]='" + row["Column_OneID"].ToString() + "'");
                        if (ret2.Length == 0) continue;
                        row["Table_OneId"] = (long)ret2[0]["Table_NameId"];
                    }
                    else
                    {
                        var ret2 = tableDataTable.Select("[Table_Name]='" + row["Table_One"].ToString() + "'");
                        if (ret2.Length == 0) continue;
                        row["Table_OneId"] = (long)ret2[0]["ID"];
                    }
                }
                if (row["Table_ManyId"].ToString() == "")
                {
                    var tableRow = tableDataTable.Select("[Table_Name]='" + row["Table_Many"].ToString() + "'");
                    row["Table_ManyId"] = (long)tableRow[0]["ID"];
                }

                // skip if we can't find the colManyID value

//GG:                if (row["Column_ManyID"].ToString() != "" & colManyID == 0) continue;

                foreach (DataColumn column in dtRelation.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    var fieldName = column.ColumnName.ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (column.ColumnName.EndsWith("Id")) continue;
                    if (isValidColumn(sharepointRelationFields, fieldName))
                        oItem[fieldName] = row.Field<string>(column.ColumnName);
                }

                oItem["Table_One"] = row.Field<string>("Table_OneId");
                oItem["Table_Many"] = row.Field<string>("Table_ManyId");
                oItem["Column_One"] = colOneID;
                oItem["Column_Many"] = colManyID;

                oItem.Update();
                row["Edit_Status"] = "Uploading";

                if (i++ >= 49)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + j + " of " + dtRelation.Rows.Count);
                    gclientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                gclientContext.ExecuteQuery();
        }

        internal bool dtDataFieldExists(DataTable dt, string fieldName)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
                if (dt.Columns[i].ColumnName == fieldName)
                    return true;
            return false;
        }


        /* dtRelation must contain the following case-sensitive fieldnames: ID, 
         * Relation_Type, Relation_Level_One, Relation_Level_Many
        */

        internal void UploadCompleteDataRecordsetToSharepoint()
        {
            GetClientContext();
            if (!AssignDataRecordsets()) return;
            _dtDatabases = Task.Run(async () => await GetDBsFromSharepoint("", new string[] { "Id", "Application_NameId", "Title" }, new string[] { "Title" })).Result;
            string dbName = "";
            long dbID = 0;

            /* SaveTablesToSharepoint requires the following case-sensitive fieldnames for an upload:
                * ID (null for new ones), Table_Name, Database_NameId, Database_Name,
                * Table_Type (Table, View, API)
                */

            DataTable dtTables = convertDataRecordsetToDataTable(tableRecordset);
            if (dtTables == null)
                MessageBox.Show("No table records to be saved");
            else
            {
                if (dtDataFieldExists(dtTables, "Database_Name"))
                    dbName = dtTables.Rows[0].Field<string>("Database_Name");
                if (dbName == "") // Only assign it this name if it doesn't already exist - change this to drop-down inputbox
                    dbName = "Dynamics OData";

                var dbRow = _dtDatabases.Select("Database_Name='" + dbName + "'");
                if(dbRow is null)
                    { MessageBox.Show("There is no database by that name."); return;}
                dbID = dbRow[0].Field<long>("Id");

                if (!dtDataFieldExists(dtTables, "Database_Name")) dtTables.Columns.Add("Database_Name", typeof(string), "'" + dbName + "'");
                if (!dtDataFieldExists(dtTables, "Database_NameId")) dtTables.Columns.Add("Database_NameId", typeof(long), "'" + dbID + "'");
                if (!dtDataFieldExists(dtTables, "Table_Type")) dtTables.Columns.Add("Database_Name", typeof(string), "'API'");
                if (!dtDataFieldExists(dtTables, "ID")) dtTables.Columns.Add("ID");

                if (!dtDataFieldExists(dtTables, "Edit_Status")) dtTables.Columns.Add("Edit_Status", typeof(string));
                foreach (DataRow row in dtTables.Rows)
                    row["Edit_Status"] = "New";

//                saveTablesToSharepoint(dtTables);
            }


            /* dtColumns must contain the following case-sensitive fieldnames: ID, Column_Name (Not Null), non-null Table_NameId or Table_Name, Column_Type
             * Table_NameId or Table_Name must be of a valid table
             */
            DataTable dtColumns = convertDataRecordsetToDataTable(columnRecordset);
            if (dtColumns == null)
                MessageBox.Show("No column records to be saved");
            else
            {
                gclientContext = this.GetClientContext(true,false,true);

                if (!dtDataFieldExists(dtColumns, "Database_Name")) dtColumns.Columns.Add("Database_Name", typeof(string), "'" + dbName + "'");
                if (!dtDataFieldExists(dtColumns, "Database_NameId")) dtColumns.Columns.Add("Database_NameId", typeof(long));
                if (!dtDataFieldExists(dtColumns, "ID")) dtColumns.Columns.Add("ID");

                if (!dtDataFieldExists(dtColumns, "Edit_Status")) dtColumns.Columns.Add("Edit_Status", typeof(string));
                foreach (DataRow row in dtColumns.Rows)
                    row["Edit_Status"] = "New";

                saveColumnsToSharepoint(dtColumns, dbID);
            }

            gclientContext = this.GetClientContext(true,false,true);

            DataTable dtRelations = convertDataRecordsetToDataTable(relationRecordset);
            if (dtRelations == null)
                MessageBox.Show("No relation records to be saved");
            else
            {
                if (dtDataFieldExists(dtRelations, "TableOne")) dtRelations.Columns["TableOne"].ColumnName = "Table_One";
                if (dtDataFieldExists(dtRelations, "TableMany")) dtRelations.Columns["TableMany"].ColumnName = "Table_Many";
                if (dtDataFieldExists(dtRelations, "ColumnOne")) dtRelations.Columns["ColumnOne"].ColumnName = "Column_One";
                if (dtDataFieldExists(dtRelations, "ColumnMany")) dtRelations.Columns["ColumnMany"].ColumnName = "Column_Many";
                //dtRelations.Columns["Description"].ColumnName = "Relation_Description";
                if (!dtDataFieldExists(dtRelations, "Table_OneId")) dtRelations.Columns.Add("Table_OneId");
                if (!dtDataFieldExists(dtRelations, "Table_ManyId")) dtRelations.Columns.Add("Table_ManyId");
                if (!dtDataFieldExists(dtRelations, "Column_OneId")) dtRelations.Columns.Add("Column_OneId");
                if (!dtDataFieldExists(dtRelations, "Column_ManyId")) dtRelations.Columns.Add("Column_ManyId");
                if (!dtDataFieldExists(dtRelations, "Relation_Type")) dtRelations.Columns.Add("Relation_Type", typeof(string), "'Foreign Key'");
                if (!dtDataFieldExists(dtRelations, "Relation_Level_One")) dtRelations.Columns.Add("Relation_Level_One", typeof(string), "'Table'");
                if (!dtDataFieldExists(dtRelations, "Relation_Level_Many")) dtRelations.Columns.Add("Relation_Level_Many", typeof(string), "'Table'");
                if (!dtDataFieldExists(dtRelations, "ID")) dtRelations.Columns.Add("ID");

                if (!dtDataFieldExists(dtRelations, "Edit_Status")) dtRelations.Columns.Add("Edit_Status", typeof(string));
                foreach (DataRow row in dtRelations.Rows)
                    row["Edit_Status"] = "New";

                saveRelationsToSharepoint(dtRelations, dbID, dbID);
            }

        ////    string dbName = "";
        //    string appName = "Dynamics";
        //    int i = 0;

        //    var clientContext = GetSPAccessToken();
        //    if (clientContext is null) return;
        //    string sharepointApplicationFields = Task.Run(async () => await GetSharePointListFields("Applications")).Result;
        //    string sharepointDatabaseFields = Task.Run(async () => await GetSharePointListFields("Databases")).Result;
        //    string sharepointTableFields = Task.Run(async () => await GetSharePointListFields("Tables")).Result;
        //    string sharepointColumnFields = Task.Run(async () => await GetSharePointListFields("Columns")).Result;
        //    string sharepointRelationFields = Task.Run(async () => await GetSharePointListFields("Relations")).Result;

        //    tableRecordset = null;

        //    var AppCol = GetDRColumnNumber(tableRecordset.GetRowData(0), "Application");
        //    if (tableRecordset.GetRowData(1).GetValue(AppCol).ToString() == "")
        //    {
        //        string ignore = "-1";
        //        var result = Utilites.ScreenEvents.ShowInputDialog(ref appName, ref ignore, "Site", "Ignore", "Enter Application Name");
        //        if (result != DialogResult.OK) return;
        //    }

        //    //DataTable dBDataTable = Task.Run(async () => await GetDBsFromSharepoint("", new string[] { "Id", "Application_NameId", "Title" }, new string[] { "Title" })).Result;
        //    var records = tableRecordset.GetDataRowIDs("");
        //    var targetList = clientContext.Web.Lists.GetByTitle("Tables");
        //    var titleRow = tableRecordset.GetRowData(0);

        //    for (int recNo = 1; recNo <= records.GetUpperBound(0)+1; recNo++)
        //    {
        //        var record = tableRecordset.GetRowData(recNo);

        //        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
        //        ListItem oItem = targetList.AddItem(itemCreateInfo);
        //        oItem["Database_Name"] = "";
        //        oItem["Title"] = "";
        //        for (int k = 0; k < record.Length; k++)
        //        {
        //            string fieldName = titleRow.GetValue(k).ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
        //            if (isValidColumn(sharepointTableFields, fieldName))
        //                oItem[fieldName] = record.GetValue(k);
        //        }

        //        //if (oItem["Database_Name"].ToString() == "") 
        //        //oItem["Database_Name"] = "Web";
        //        dbName = oItem["Database_Name"].ToString();
        //        if (dbName == "") dbName = "Dynamics OData";
        //         row = dBDataTable.Rows.Find(dbName);
        //        oItem["Database_Name"] = row["Id"];
        //        if (oItem["Table_Type"].ToString() == "") oItem["Table_Type"] = "API";
        //        if (oItem["Application"].ToString() == "") oItem["Application"] = appName;

        //        // Create a hash if the item doesn't already have a unique identifier.
        //        if (oItem["Title"].ToString() == "") 
        //            oItem["Title"] = GenerateHash(oItem["Application"] + "." + dbName + "." + oItem["Table_Name"]);

        //        oItem.Update();

        //        if (++i >= 20)
        //        {
        //            Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + recNo + " of " + records.GetUpperBound(0));
        //            clientContext.ExecuteQuery();
        //            i = 0;
        //        }
        //    }
        //    if (i > 0)
        //        clientContext.ExecuteQuery();



        //        /*************** Upload Columns *****************/
        //    if (dbName == "") dbName = "Dynamics OData";
        //    var MyDBID = (long)dBDataTable.Rows.Find(dbName)["Id"];

        //    records = columnRecordset.GetDataRowIDs("");
        //    DataTable tableDataTable = Task.Run(async () => await GetTablesFromSharepoint(appName, MyDBID, new string[] { }, new string[] { "Table_Name" })).Result;

        //    targetList = clientContext.Web.Lists.GetByTitle("Columns");
        //    titleRow = columnRecordset.GetRowData(0);
        //    int tableNameID = GetDRColumnNumber(titleRow, "table_name"); //GG: Changed to Table_Name
        //    int columnNameID = GetDRColumnNumber(titleRow, "column_Name"); //GG: Changed to Table_Name
        //    int tabletypeID = GetDRColumnNumber(titleRow, "Table_Type");
        //    i = 0;
        //    for (int recNo =1; recNo <= records.GetUpperBound(0)+1; recNo++)
        //    {
        //        var record = columnRecordset.GetRowData(recNo);

        //        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
        //        ListItem oItem = targetList.AddItem(itemCreateInfo);

        //        oItem["Title"] = "";
        //        for (int k = 0; k < record.Length; k++)
        //        {
        //            string fieldName = titleRow.GetValue(k).ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
        //            if (fieldName == "Table_Name") continue;
        //            if (record.GetValue(k) is null) continue;
        //            if (isValidColumn(sharepointColumnFields, fieldName))
        //               oItem[fieldName] = record.GetValue(k);
        //        }
        //        var tableName = record.GetValue(tableNameID).ToString();
        //        //                oItem["Is_Primary"] = (record.GetValue(columnNameID).ToString() == LookupFromTableDataset(record.GetValue(tableNameID).ToString(), "HideprimaryIdAttribute")) ? true : false;
        //        if (oItem["Title"].ToString() == ""){
        //            oItem["TableGUID"] = GenerateHash(appName + "." + dbName + "." + LookupFromTableDataset(tableName, "Table_Type") + "." + tableName + "." + record.GetValue(columnNameID).ToString());
        //            oItem["Title"] = dbName + "." + tableName;
        //            //                    oItem["TableGUID"] = GenerateHash(appName + "." + dbName + "." + record.GetValue(tableNameID).ToString());// LookupFromTableDataset(record.GetValue(tableNameID).ToString(), "UniqueID");
        //            oItem["dbTable"] = tableDataTable.Rows.Find(tableName)["Id"];
        //        }
        //        oItem.Update();

        //        if (++i >= 50)
        //        {
        //            Utilites.ScreenEvents.DisplayVisioStatus("Uploading Columns: " + recNo + " of " + records.GetUpperBound(0));
        //            clientContext.ExecuteQuery();
        //            i = 0;
        //        }
        //    }
        //    if (i > 0)
        //        clientContext.ExecuteQuery();
            


        //    /*************** Upload Relations *********************/
        //    records = relationRecordset.GetDataRowIDs("");
        //    records = columnRecordset.GetDataRowIDs("");

        //    DataTable columnDataTable = Task.Run(async () => await GetColumnsFromSharepoint(MyDBID, 0, new string[] {"Id", "Title", "Table_Name", "Column_Name"}, new string[] { "Table_Name", "Column_Name" })).Result;

        //    targetList = clientContext.Web.Lists.GetByTitle("Relations");
        //    titleRow = relationRecordset.GetRowData(0);
        //    int DatabaseCol = GetDRColumnNumber(titleRow, "Database");
        //    int EntityOneCol = GetDRColumnNumber(titleRow, "TableOne");
        //    int EntityManyCol = GetDRColumnNumber(titleRow, "TableMany"); 
        //    int columnOneCol = GetDRColumnNumber(titleRow, "ColumnMany");
        //    int columnManyCol = GetDRColumnNumber(titleRow, "ColumnMany");
        //    int UniqueCol = GetDRColumnNumber(titleRow, "Title"); // storing the GUID in the title field
        //    int ConnectionTypeCol = GetDRColumnNumber(titleRow, "Connection_Type");
        //    for (int recNo = 1; recNo < records.GetUpperBound(0); recNo++)
        //    {
        //        var record = relationRecordset.GetRowData(recNo);

        //        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
        //        ListItem oItem = targetList.AddItem(itemCreateInfo);

        //        oItem["Title"] = "";
        //        for (int k = 0; k < record.Length; k++)
        //        {
        //            string fieldName = titleRow.GetValue(k).ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
        //            if (isValidColumn(sharepointRelationFields, fieldName))
        //                oItem[fieldName] = record.GetValue(k);
        //        }
        //        string sEntityOne = EntityOneCol < 0 ? "" : record.GetValue(EntityOneCol).ToString();
        //        string sEntityMany = EntityManyCol < 0 ? "" : record.GetValue(EntityManyCol).ToString();
        //        string sColumnOne = columnOneCol < 0 ? "" : record.GetValue(columnOneCol).ToString();
        //        string sColumnMany = columnManyCol < 0 ? "" : record.GetValue(columnManyCol).ToString();
        //        // If this hasn't been uploaded yet then we generate IDs ourself
        //        if (UniqueCol == -1)
        //        {
        //            dbName = DatabaseCol == -1 ? "Dynamics OData" : dbName = record.GetValue(DatabaseCol).ToString();

        //            if (sEntityOne != "") 
        //                if(record.GetValue(EntityOneCol).ToString() == "")
        //                    oItem["TableGUID"] = tableDataTable.Rows.Find(sEntityOne)["Id"];
        //            if (sEntityOne != "")
        //                if (record.GetValue(EntityManyCol).ToString() == "")
        //                    oItem["TableGUID"] = tableDataTable.Rows.Find(sEntityOne)["Id"];
        //            if (sColumnMany != "")
        //                if (record.GetValue(columnManyCol).ToString() == "")
        //                    oItem["TableGUID"] = columnDataTable.Rows.Find(new object[] { sEntityMany, sColumnMany })["Id"];

        //            if (ConnectionTypeCol ==-1) 
        //                    oItem["Connection_Type"] = "Foreign Key";

        //            oItem["Title"] =  GenerateHash(appName + "." + dbName + "." + record.GetValue(EntityOneCol).ToString() + "." + record.GetValue(columnOneCol).ToString() + "." + record.GetValue(EntityManyCol).ToString() + "." + record.GetValue(columnManyCol).ToString());
        //        }
        //        oItem.Update();

        //        if (++i >= 50)
        //        {
        //            Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + recNo + " of " + records.GetUpperBound(0));
        //            clientContext.ExecuteQuery(); 
        //            i = 0;
        //        }
        //    }
        //    if (i > 0)
        //        clientContext.ExecuteQuery();
        }

    }
}

