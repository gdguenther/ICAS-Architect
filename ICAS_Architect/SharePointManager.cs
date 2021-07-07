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



        /* Sets up the http register. To check if the requisite tables exist, pass in WarnIfTablesDontExist
         *  To move to a different Sharepoint Location pass in AskForSharepointLocation. 
         *  Sharepoint connections also timeout, so ask the client to refresh from time to time.
         */
        internal ClientContext GetClientContext(bool WarnIfTablesDontExist = true, bool AskForSharepointLocation = false, bool Refresh = false)
        {
            string SPRepositoryFromRegistry = Globals.ThisAddIn.registryKey.GetValue("SharepointName", "https://icas1854.sharepoint.com/sites/Architecture/Test").ToString();
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
                string loginSite = SPRepository.Substring(0, SPRepository.IndexOf("/sites/"));
                httpDownloadClient = new HttpDownloadClient(loginSite);
                httpDownloadClient.Connect(loginSite + "/_api/web/lists");
            }

            OfficeDevPnP.Core.AuthenticationManager authMgr = new OfficeDevPnP.Core.AuthenticationManager();
            gclientContext = authMgr.GetAzureADAccessTokenAuthenticatedContext(SPRepository, httpDownloadClient.accessToken);

            if (WarnIfTablesDontExist)
            {
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



        /************************************************************************************
         * 
         * Use CSOM to create the tables. Can do it with the REST API, but it's not
         * documented.
         * 
         * *********************************************************************************/


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

                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Application_Name' StaticName='Application_Name' DisplayName='Application_Name' List='" + refList.Id + "' ShowField = 'Title' RelationshipDeleteBehaviorType='Restrict' IsRelationship='TRUE' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
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
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Database_Name' StaticName='Database_Name' DisplayName='Database_Name' List='" + refList.Id + "' ShowField = 'Title' RelationshipDeleteBehaviorType='Restrict' IsRelationship='TRUE'  Indexed='TRUE' Required='TRUE'/>", true, AddFieldOptions.AddToDefaultContentType);
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
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_Name' StaticName='Table_Name' DisplayName='Table_Name' List='" + refList.Id + "'  ID='{" + fielda + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' IsRelationship='TRUE'  Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
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
                    var fieldc = Guid.NewGuid().ToString();
                    var fielde = Guid.NewGuid().ToString();
                    var fieldg = Guid.NewGuid().ToString();

                    /* dtRelation must contain the following case-sensitive fieldnames: ID, Relation_Name, Table_OneId, Table_One, Table_ManyId, Table_Many,
                     * Column_OneId, Column_ManyId, Column_ManyId, Column_Many, Relation_Type, Relation_Level_One, Relation_Level_Many */

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Relation_Description' Name='Relation_Description'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_One' StaticName='Table_One' DisplayName='Table_One' List='" + refList.Id + "'  ID='{" + fielda + "}' ShowField = 'Table_Name'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_Many' StaticName='Table_Many' DisplayName='Table_Many' List='" + refList.Id + "'  ID='{" + fieldc + "}' ShowField = 'Table_Name'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_One' StaticName='Column_One' DisplayName='Column_One' List='" + refListCol.Id + "'  ID='{" + fielde + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='FALSE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Column_Many' StaticName='Column_Many' DisplayName='Column_Many' List='" + refListCol.Id + "'  ID='{" + fieldg + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='FALSE' />", true, AddFieldOptions.AddToDefaultContentType);

                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_OneId' StaticName='Table_OneId' DisplayName='Table_OneId' List='" + refList.Id + "'  ID='{" + fielda + "}' ShowField = 'Title'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    //list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_One' StaticName='Table_One' DisplayName='Table_One' List='" + refList.Id + "' ID='{" + fieldb + "}' ShowField='Table_Name' FieldRef='{" + fielda + "}' />", true, AddFieldOptions.AddToDefaultContentType);
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



        /**************************************************************************************************************************
         * 
         * Visio has a default internal "DataRecordset" structure that pulls in external data into the "External Data" window.
         * It's neat in that you can link this data to shapes and when you refresh the data, you can report on any differences
         * between the external data and the shape data you've drawn with. The data also gets saved with the Visio sheet, which is
         * great if you are in the middle of a drawing, or a huge waste of space otherwise.  It's also ugly to work with. 
         * Row zero of a datarecordset is the fieldnames that were returned, and row 1 to n are the returned item.
         * 
         * ************************************************************************************************************************/


        /******* Pull each of our five main datarecordsets in. *******/
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


        /********************* Sharepoint connection string for our datarecordsets **********************/
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

        /**************** deletes a named datarecordset from Visio *********************************/
        internal void DeleteDataRecordset(string listName)
        {
            for (int i = 1; i <= vApplication.ActiveDocument.DataRecordsets.Count; i++)
                if (vApplication.ActiveDocument.DataRecordsets[i].Name == "Tables")
                    vApplication.ActiveDocument.DataRecordsets[i--].Delete();
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


        internal Visio.DataRecordset getColumnDataRecordset(string tableName = "All", string dBName = "All")
        {
            string whereCond = null;
            if (tableName != "All")
            {
                whereCond = whereCond == null ? " WHERE " : " AND ";
                whereCond = whereCond + "[Table_Name] = '" + tableName + "'";
            }
            //if (dBName != "All")
            //{
            //    whereCond = whereCond == null ? " WHERE " : " AND ";
            //    whereCond = whereCond + "[Database_Name] = '" + dBName + "'";
            //}

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
            applicationRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Applications];", 0, "Applications");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return applicationRecordset;
        }


        internal Visio.DataRecordset getDatabaseDataRecordset(string tableName = "All", string dBName = "All")
        {
            DeleteDataRecordset("Databases");
            string connString = CreateSharepointConnectionString("Database");
            databaseRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Databases];", 0, "Databases");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
            //GG: todo: Add primary key, add important column ids
            return databaseRecordset;
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





 /**********************************************************************************************************
 * 
 *  The following is code uses Sharepoint's OData Interface to pull information. This is the better way to 
 *  write data to Sharepoint, but keeps us from generating queries.
 * 
   *********************************************************************************************************/



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

            // Use Linq to add DB and Application fields to each record
            dataTable.Columns.Add("Database_Name", Type.GetType("System.String"));
            dataTable.Columns.Add("Application_NameId", Type.GetType("System.Int64"));
            dataTable.Columns.Add("Application_Name", Type.GetType("System.String"));
            dataTable.AsEnumerable().Join(dtDb.AsEnumerable(),  _dtmater => Convert.ToString(_dtmater["Database_NameId"]),   _dtchild => Convert.ToString(_dtchild["id"]),
                (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                    o => {
                        o._dtmater.SetField("Database_Name", o._dtchild["Database_Name"].ToString());
                        o._dtmater.SetField("Application_NameId", (long)o._dtchild["Application_NameId"]);
                        o._dtmater.SetField("Application_Name", o._dtchild["Application_Name"].ToString());
                        }
                ) ;

            return dataTable;
        }



        internal DataTable GetColumnsFromSharepointWrapper(long TableOneID)
        {
            // if the table exists locally, return the data table
            if (_dtColumns != null)
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
            // Retrieve the parent list, but only if needed (Tables)
            if(_dtTables == null)
                _dtTables = await GetTablesFromSharepoint("", dbID ,  new string[] { }, new string[] { "Database_NameId", "Title" });
            else if (_dtTables.Select("ID=" + tableID).Length == 0)
                _dtTables.Merge(await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" }));

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

            try
            {
                // Use Linq to add Table and DB fields to each record
                dtColumn.Columns["Title"].ColumnName = "Column_Name";
                dtColumn.Columns.Add("Table_Name", Type.GetType("System.String"));
                dtColumn.Columns.Add("Database_NameId", Type.GetType("System.Int64"));
                dtColumn.Columns.Add("Database_Name", Type.GetType("System.String"));
                dtColumn.AsEnumerable().Join(_dtTables.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Table_NameId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o =>
                        {
                            o._dtmater.SetField("Table_Name", o._dtchild["Table_Name"].ToString());
                            o._dtmater.SetField("Database_NameId", (long)o._dtchild["Database_NameId"]);
                            o._dtmater.SetField("Database_Name", o._dtchild["Database_Name"].ToString());
                        }
                    );
            }
            catch (Exception e) { Console.WriteLine(e.Message); }

            return dtColumn;
        }


        internal async Task<DataTable> GetColumnsFromSharepointByDB(long dbID)
        {
            DataTable dt = null, dataTable = null;
            // Retrieve the parent list (Tables)
            if (_dtTables == null)
                _dtTables = await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" });
            else if (_dtTables.Select("Database_NameId=" + dbID).Length == 0)
            {
                dt = await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" });
                if (dt != null)
                    _dtTables.Merge(dt);
            }
            if (_dtTables == null) return null;

            foreach (DataRow row in _dtTables.Rows)
            {
                if (dataTable == null)
                    dataTable = await GetColumnsFromSharepoint(0, (long)row["ID"], new string[] { }, new string[] { "Id", "Title, Table_NameId" });
                else if (dataTable.Select("Table_NameId=" + row["ID"]).Length == 0)
                    try
                    {
                        dt = await GetColumnsFromSharepoint(0, (long)row["ID"], new string[] { }, new string[] { "Id", "Title, Table_NameId" });
                        if(dt != null)
                            dataTable.Merge(dt, true, MissingSchemaAction.Ignore);
                    } catch (Exception e) {    Console.WriteLine(e.Message); }
            }
            return dataTable;
        }


        internal async Task<DataTable> GetRelationsFromSharepointByDB(long dbID, bool ManyDirection = false)
        {
            // Retrieve the parent list (Tables)
            if (_dtTables == null)
                _dtTables = await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" });
            else if (_dtTables.Select("Database_NameId=" + dbID).Length == 0)
                _dtTables.Merge(await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" }));

            DataTable dtRelations = null;
            foreach(DataRow row in _dtTables.Rows)
            {
                if(dtRelations == null)
                    dtRelations = await GetRelationsFromSharepoint(0, (long)row["ID"], new string[] { }, new string[] { }, ManyDirection);
                else if (dtRelations.Select((ManyDirection ? "Table_ManyId=" : "Table_OneId") + row["ID"]).Length == 0)
                    try
                    {
                        DataTable dt = await GetRelationsFromSharepoint(0, (long)row["ID"], new string[] { }, new string[] { }, ManyDirection);
                        if (dt != null)
                            dtRelations.Merge(dt, true, MissingSchemaAction.Ignore);
                    }
                    catch (Exception e) { Console.WriteLine(e.Message); }
            }
            return dtRelations;
        }


        // Get all columns by DB is not working yet. For our huge initial uploads it needs fixing.
        internal async Task<DataTable> GetRelationsFromSharepoint(long dbID, long tableID, string[] fieldList, string[] PrimaryKey, bool ManyDirection)
        {
            // Retrieve the parent list (Tables)
            if (_dtTables == null)
                _dtTables = await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" });
            else if (_dtTables.Select("ID=" + tableID).Length == 0)
                _dtTables.Merge(await GetTablesFromSharepoint("", dbID, new string[] { }, new string[] { "Database_NameId", "Title" }));

            string selectString = (fieldList.Length > 0) ? "$select=" : "";
            for (int i = 0; i < fieldList.Length; i++)
                selectString += (i > 0 ? "," : "") + (fieldList[i]);

            DataTable dtRelation = null;
            try
            {
                if (tableID != 0)
                    if(ManyDirection)
                        dtRelation = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Relations')/items?" + selectString + "&$filter=Table_ManyId eq '" + tableID + "'&$orderby=Title", 0, PrimaryKey);
                    else    
                        dtRelation = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Relations')/items?" + selectString + "&$filter=Table_OneId eq '" + tableID + "'&$orderby=Title", 0, PrimaryKey);
            }
            catch (Exception e)
            { Console.WriteLine(e.Message); }

            if (dtRelation is null) return null;

            try
            {
                // Use Linq to add Table and DB fields to each record
                dtRelation.Columns.Add("Column_One", typeof(string));
                dtRelation.Columns.Add("Table_OneId", typeof(long));
                dtRelation.Columns.Add("Table_One", typeof(string));
                dtRelation.Columns["Title"].ColumnName = "Relation_Name";
                dtRelation.Columns.Add("Column_Many", typeof(string));
                dtRelation.Columns.Add("Table_ManyId", typeof(long));
                dtRelation.Columns.Add("Table_Many", typeof(string));
                dtRelation.AsEnumerable().Join(_dtTables.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Table_NameId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => {
                            o._dtmater.SetField("Table_Name", o._dtchild["Table_Name"].ToString());
                            o._dtmater.SetField("Database_NameId", (long)o._dtchild["Database_NameId"]);
                            o._dtmater.SetField("Database_Name", o._dtchild["Database_Name"].ToString());
                        }
                    );
            }
            catch (Exception e) { Console.WriteLine(e.Message); }

            return dtRelation;
        }



        // wrapper that ensures that we can retrieve as many rows as we like.
        internal async Task<DataTable> retrieveDataTable(string appUrl, int iteration, string[] PrimaryKey)
        {
            const int iterSize = 1000;
            var client = httpDownloadClient.client;
            client.DefaultRequestHeaders.Accept.Clear();
            System.Net.Http.Headers.MediaTypeWithQualityHeaderValue acceptHeader = System.Net.Http.Headers.MediaTypeWithQualityHeaderValue.Parse("application/json;odata.metadata=none");
//            System.Net.Http.Headers.MediaTypeWithQualityHeaderValue acceptHeader = System.Net.Http.Headers.MediaTypeWithQualityHeaderValue.Parse("application/json");
            client.DefaultRequestHeaders.Accept.Add(acceptHeader);

            string pageURL = appUrl + "&$skiptoken=Paged=TRUE%26p_SortBehavior=0%26p_ID=" + (iteration * iterSize) + "&$top=" + iterSize;// "&$skiptoken=Paged=TRUE";// + " &$top=" + iterSize + "&$skip=" + (iteration * iterSize).ToString();

            var response = await client.GetAsync(pageURL);
            var responseText = await response.Content.ReadAsStringAsync();
            try { response.EnsureSuccessStatusCode(); }
            catch (Exception e)
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

            var keys = new DataColumn[PrimaryKey.Length];
            for (int i = 0; i < PrimaryKey.Length; i++)
                keys[i] = dataTable.Columns[PrimaryKey[i]];
            // I may add the Primary Key field back in later.   dataTable.PrimaryKey = keys;

            return dataTable;
        }


        internal string GenerateHash(string entityName)
        {
            MD5 md5 = MD5.Create();

            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(entityName.ToLower());
            byte[] hashBytes = md5.ComputeHash(inputBytes);

            // Convert the byte array to hexadecimal string
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hashBytes.Length; i++)
                sb.Append(hashBytes[i].ToString("X2"));

            return sb.ToString();
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
                if (values[i].SelectToken("Title").ToString() != "ID")
                    columnList = columnList + values[i].SelectToken("Title").ToString().ToLower() + ",";

            return columnList;
        }


        internal DataTable convertDataRecordsetToDataTable(Visio.DataRecordset DR)
        {
            // convert string to stream
            byte[] byteArray = Encoding.UTF8.GetBytes(System.Text.RegularExpressions.Regex.Replace(DR.DataAsXML.Replace("\r", "").Replace("\n", "").Replace("rs:", "").Replace("s:", "").Replace("dt:", "").Replace("z:", "").Replace("\t", "").Replace("name='_Visio_RowID_'", ""), "[<]xml.*#RowsetSchema'[>]", "<xml>"));
            System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);

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

                // Read all Sharepoint-ok'd columns
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

                if (i++ >= 0)
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
                bool AddNew = (row.Field<long?>("ID") == null);
                    // Add if there's no ID, otherwise edit.
                    if (AddNew)
                        oItem = targetList.AddItem(new ListItemCreationInformation());
                    else
                        oItem = targetList.GetItemById(row.Field<int>("ID"));

                    foreach (DataColumn column in dtColumns.Columns)
                    {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                        string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_");
                        if (column.ColumnName == "Column_Name") continue;
                        if (isValidColumn(sharepointColumnFields, fieldName)) // if the column is not in sharepoint, skip the column
                            if(!AddNew | row[column.ColumnName].ToString() == "" ) // if we're adding something new and the column is null, skip the column
                                oItem[fieldName] = row.Field<string>(column.ColumnName);
                    }

                    if (row.Field<long?>("ID") == null)
                    {
                        var dr = tableDataTable.Select("[Table_Name] = '" + row.Field<string>("Table_Name").ToString() + "'");
                        if (dr.Length == 0)
                        { MessageBox.Show("Cannot find table " + row.Field<string>("Table_Name").ToString() + ". \n Skipping insertion of " + row.Field<string>("Column_Name").ToString()); continue; }
                        var tableName = dr[0]["Table_Name"].ToString();
                        var dbName = dr[0]["Database_Name"].ToString();
                        var appName = dr[0]["Application_Name"].ToString();
                        var tableType = dr[0]["Table_Type"].ToString();

                        oItem["ColumnGUID"] = GenerateHash(appName + "." + dbName + "." + dr[0]["Table_Name"].ToString() + "." + row.Field<string>("Column_Name"));
                        //                    oItem["dbTable"] = tableDataTable.Rows.Find(tableName)["Id"];
                        //I'm going to try to set this with Id again, but sharepoint sems to want to use the title field
                        oItem["Title"] = row.Field<string>("Column_Name");
                        oItem["Table_Name"] = dr[0]["ID"].ToString();
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

        internal bool getColumnIDFromDatatable(ref long ColumnID, string ColumnName, ref long TableID, string TableName)
        {
            // Get the TableID
            var tableRet = _dtTables.Select("Table_Name='" + TableName.ToString() + "'");
            TableID = (TableID > 0) ? TableID : (long)tableRet[0]["ID"];
            long tablePass = TableID; 

            if (TableID == 0)                   return false; // If we didn't find the Table above, just return false;
            if (ColumnID > 0 & TableID > 0)     return true; // everything is fine, so we can continue;
            if (ColumnName == "") return true; // we're not looking for a column

            // Get the columnID
            if (ColumnID == 0)
            {
                if (_dtColumns is null) // if _dtColumns is null, let's initialise it just for the sake of it
                    _dtColumns = Task.Run(async () => await GetColumnsFromSharepoint(0, tablePass, new string[] { "Id", "Title, Table_NameId" }, new string[] { "Table_Name", "Title" })).Result;

                var ret = _dtColumns.Select("[Column_Name]='" + ColumnName + "' and [Table_NameId]='" + TableID + "'");
                if (ret.Length > 0)
                {
                    ColumnID = (long)ret[0]["ID"];
                    return true;
                }

                //If we can't find the ColumnID, let's try importing all columns in that table from Sharepoint
                   _dtColumns.Merge(Task.Run(async () => await GetColumnsFromSharepoint(0, tablePass, new string[] { "Id", "Title, Table_NameId" }, new string[] { "Table_Name", "Title" })).Result);

                    ret = _dtColumns.Select("[Column_Name]='" + ColumnName + "' and [Table_NameId]='" + TableID + "'");
                    if (ret.Length > 0)
                    {
                        ColumnID = (long)ret[0]["ID"];
                        return true;
                    }
                }
            return true;
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
                 *          valid types can be "Column", "Table", "Application", "Actor", 
                 *          When linking columns, set Relation_Levels to "Column", but be sure to include the TableID as well
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
                if (row["ID"].ToString() == "") //change to isnull
                    oItem = targetList.AddItem(new ListItemCreationInformation());
                else
                    oItem = targetList.GetItemById(row["ID"].ToString());

                if(row["Relation_Type"].ToString() == "Foreign Key")
                    if (Globals.ThisAddIn.drawingManager.IsEntityInSkipList(row["Table_One"].ToString()))  continue; //we want to keep skip list columns but remove their relations

                long colOneID = row["Column_OneId"].ToString() == "" ? (long) 0 : (long) row["Column_OneID"];
                long colManyID = row["Column_ManyID"].ToString() == "" ? 0 : (long) row["Column_ManyID"];
                long tabOneID = row["Table_OneID"].ToString() == "" ? 0 : (long) row["Table_OneID"];
                long tabManyID = row["Table_ManyID"].ToString() == "" ? 0 : (long) row["Table_ManyID"];
                if (! getColumnIDFromDatatable(ref colOneID, row["Column_One"].ToString(), ref tabOneID, row["Table_One"].ToString()) ) continue;
                if (! getColumnIDFromDatatable(ref colManyID, row["Column_Many"].ToString(), ref tabManyID, row["Table_Many"].ToString()) ) continue;

                if (row["Table_OneId"].ToString() == "")
                {
                    var ret = tableDataTable.Select("[Table_Name]='" + row["Table_One"].ToString() + "'");
                    if (ret.Length == 0) continue;
                    tabOneID = (long)ret[0]["ID"];
                }
                if (row["Table_ManyId"].ToString() == "")
                {
                    var ret = tableDataTable.Select("[Table_Name]='" + row["Table_Many"].ToString() + "'");
                    if (ret.Length == 0) continue;
                    tabManyID = (long)ret[0]["ID"];
                }


                foreach (DataColumn column in dtRelation.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    var fieldName = column.ColumnName.ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (column.ColumnName.EndsWith("Id")) continue;
                    if (isValidColumn(sharepointRelationFields, fieldName))
                        oItem[fieldName] = row.Field<string>(column.ColumnName);
                }

                oItem["Table_One"] = tabOneID;
                oItem["Table_Many"] = tabManyID;
                if (colOneID > 0) oItem["Column_One"] = colOneID;
                if (colManyID > 0) oItem["Column_Many"] = colManyID;

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
                if (!dtDataFieldExists(dtTables, "Table_Type")) dtTables.Columns.Add("Table_Type", typeof(string), "'API'");
                if (!dtDataFieldExists(dtTables, "ID")) dtTables.Columns.Add("ID");

                if (!dtDataFieldExists(dtTables, "Edit_Status")) dtTables.Columns.Add("Edit_Status", typeof(string));

                foreach (DataRow row in dtTables.Rows)
                {
                    if (row["Database_Name"].ToString() == "")
                    {
                        row["Database_Name"] = dbName;
                        row["Table_Type"] = "API";
                    }
                    row["Edit_Status"] = "New";
                }
                saveTablesToSharepoint(dtTables);
            }


            /* dtColumns must contain the following case-sensitive fieldnames: ID, Column_Name (Not Null), non-null Table_NameId or Table_Name, Column_Type
             * Table_NameId or Table_Name must be of a valid table
             */
            DataTable dtColumns = convertDataRecordsetToDataTable(columnRecordset);
            if (dtColumns == null)
                MessageBox.Show("No column records to be saved");
            else
            {
                gclientContext = this.GetClientContext(true, false, true);

                if (!dtDataFieldExists(dtColumns, "Database_Name")) dtColumns.Columns.Add("Database_Name", typeof(string), "'" + dbName + "'");
                if (!dtDataFieldExists(dtColumns, "Database_NameId")) dtColumns.Columns.Add("Database_NameId", typeof(long));
                if (!dtDataFieldExists(dtColumns, "ID")) dtColumns.Columns.Add("ID");

                if (!dtDataFieldExists(dtColumns, "Edit_Status")) dtColumns.Columns.Add("Edit_Status", typeof(string));
                foreach (DataRow row in dtColumns.Rows)
                    row["Edit_Status"] = "New";

                saveColumnsToSharepoint(dtColumns, dbID);
            }

            gclientContext = this.GetClientContext(true, false, true);

            /* dtRelation must contain the following case-sensitive fieldnames: ID, 
             * Relation_Type, Relation_Level_One, Relation_Level_Many
            */

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
            MessageBox.Show("Complete");
        }
    }
}

