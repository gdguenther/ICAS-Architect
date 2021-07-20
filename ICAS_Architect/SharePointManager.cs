using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Data;
using System.Security.Cryptography;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using XL = Microsoft.Office.Interop.Excel;

namespace ICAS_Architect
{
    internal class DTData
    {
        public DataTable Applications = null;
        public DataTable Databases = null;
        public DataTable Tables = null;
        public DataTable Columns = null;
        public DataTable Relations = null;
    }

    public enum RecordStatus
    {
        DoesNotExist,
        ExistsAndHasChanged,
        ExistsAndHasNotChanged
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
        private DTData repoData = null;
        private DTData editData = null;

        List<ComboListInfo> DatabaseList = null;

        internal SharepointManager()
        {
            vApplication = Globals.ThisAddIn.Application;
            repoData = new DTData();
            editData = new DTData();
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
            //    gclientContext.ExecuteQueryRetry();

                if (!gclientContext.Web.ListExists("Tables") | !gclientContext.Web.ListExists("Columns") | !gclientContext.Web.ListExists("Relations"))
                {
                    System.Windows.Forms.MessageBox.Show("This Sharepoint site does not have the required lists.\nPlease see your administrator for the correct repository.", "Missing Repository");
                    return null;
                }
            }
            return (gclientContext);
        }



        public static bool ViewExists(List list, string viewName)
        {
            if (string.IsNullOrEmpty(viewName))
                return false;

            foreach (var view in list.Views)
                if (view.Title.ToLowerInvariant() == viewName.ToLowerInvariant())
                    return true;

            return false;
        }


        internal void CreateBackendView(string listName)
        {
            //These are the sharepoint field we would like to ignore in our view.
            const string IgnoreFieldList = "FileSystemObjectType,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ContentTypeId,ComplianceAssetId,Created,AuthorId,OData__UIVersionString,Attachments,GUID,content type,item child count,folder child count,label setting,retention label applied,label applied by,item is a record,app created by,app modified by";

            List targetList = gclientContext.Web.Lists.GetByTitle(listName);
            ViewCollection viewCollection = targetList.Views;
            gclientContext.Load(viewCollection);
            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = "ICAS Backend";
            viewCreationInformation.ViewTypeKind = ViewType.Html;

            string tmp = Task.Run(async () => await GetSharePointListFields(listName, "System")).Result;
            var spListFields = tmp.Split(',');

            string wantedFields = "";
            foreach (var fieldName in spListFields)
                if (!IgnoreFieldList.ToLower().Contains("," + fieldName.ToLower() + ",") & fieldName != "")
                    wantedFields += fieldName + ",";

            if(wantedFields.Length > 0)
                wantedFields = wantedFields.Substring(0, wantedFields.Length - 1);


            viewCreationInformation.ViewFields = wantedFields.Split(',');

            Microsoft.SharePoint.Client.View listView = viewCollection.Add(viewCreationInformation);
            gclientContext.ExecuteQuery();

            // Code to update the display name for the view.
            listView.Title = "ICAS Backend";

            listView.Update();
            gclientContext.ExecuteQuery();
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


                //CreateBackendView("Applications");
                


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

                //CreateBackendView("Databases");


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


                // For tables, we can also add a user view which would use listView.Aggregations = "<FieldRef Name='Title' Type='COUNT'/>";
                //CreateBackendView("Tables");


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

                //CreateBackendView("Columns");


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
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_One' StaticName='Table_One' DisplayName='Table_One' List='" + refList.Id + "'  ID='{" + fielda + "}' ShowField = 'Full_Table_Name'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Lookup' Name='Table_Many' StaticName='Table_Many' DisplayName='Table_Many' List='" + refList.Id + "'  ID='{" + fieldc + "}' ShowField = 'Full_Table_Name'  RelationshipDeleteBehaviorType='Cascade' Indexed='TRUE' Required='TRUE' />", true, AddFieldOptions.AddToDefaultContentType);
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

                //CreateBackendView("Relations");

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
                Debug.WriteLine(e.Message);
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
                Debug.WriteLine(e.Message);
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

        internal async Task<DataTable> GetApplicationsFromSharepoint(string AppID, string[] fieldList)
        {
            string selectString = (fieldList.Length > 0) ? "$select=" : "";
            string filterString = (AppID != "") ? "&$filter=Id eq '" + AppID + "'" : "";

            for (int i = 0; i < fieldList.Length; i++)
                selectString = selectString + (i > 0 ? "," : "") + (fieldList[i]);

            DataTable tmpdt= await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Applications')/items?" + selectString + "&" + filterString, 0,  new string[] { "ID" });
            if (tmpdt != null)
                tmpdt.Columns["Title"].ColumnName = "Application_Name";
            return tmpdt;
        }

        internal async Task<DataTable> GetDBsFromSharepoint(long AppID)
        {
            // Retrieve the parent list (Applications)
            if (repoData.Applications is null) repoData.Applications = await GetApplicationsFromSharepoint("", new string[] { });

            string filterString = (AppID > 0) ? "$filter=Application_NameId eq '" + AppID + "'" : "";

            DataTable dbDT = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Databases')/items?" + filterString, 0, new string[] { "ID" });
            if (dbDT == null) return null;
            dbDT.Columns["Title"].ColumnName = "Database_Name";

            dbDT.Columns.Add("Application_Name", typeof(string));
            dbDT.AsEnumerable().Join(repoData.Applications.AsEnumerable(),
                _dtmater => Convert.ToString(_dtmater["Application_NameId"]),
                _dtchild => Convert.ToString(_dtchild["id"]),
                (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                    o => o._dtmater.SetField("Application_Name", o._dtchild["Application_Name"].ToString())
                );
            dbDT.Columns.Add("AppAndDB", typeof(string), "Application_Name + ' - ' + Database_Name");


            DatabaseList = new List<ComboListInfo>();
            foreach (DataRow row in dbDT.Rows)
                DatabaseList.Add(new ComboListInfo((long) row["ID"], row["AppAndDB"].ToString(), "DB", (long) row["Application_NameId"]));

            return dbDT;
        }


        internal async Task<DataTable> GetTablesFromSharepoint(long dBId)//, string[] fieldList)
        {
            // Retrieve the parent list (Databases)
            DataTable dtDb = repoData.Databases;
            if (dtDb is null) dtDb = await GetDBsFromSharepoint( 0);
            repoData.Databases = dtDb;

            string filterString = (dBId > 0) ? "$filter=Database_NameId eq '" + dBId.ToString() + "'" : "";

            DataTable dataTable = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Tables')/items?" + filterString, 0, new string[] { "ID" });
            if (dataTable is null) return null;

            // Use Linq to add DB and Application fields to each record
            dataTable.Columns["Title"].ColumnName = "Full_Table_Name";
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


        internal async Task<DataTable> GetColumnsFromSharepoint(long dbID, long tableID = 0)
         {
            int j = 0;
            string filter = "";
            DataTable dt = null, dtColumn = null;

            // pull all tables from this particular database into our DataTable
            if (repoData.Tables == null)
                repoData.Tables = await GetTablesFromSharepoint(dbID);
            else if (repoData.Tables.Select("Database_NameId=" + dbID).Length == 0)
            {   // Otherwise we will pull all tables in the database in.
                dt = await GetTablesFromSharepoint(dbID);
                if (dt == null) return null;
                repoData.Tables.Merge(dt, true, MissingSchemaAction.Ignore);
            }

            if (tableID > 0)
            {
                // if we are asking for a specific table, this is all we have to do
                dtColumn = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Columns')/items?" + "&$filter=Table_NameId eq " + tableID + "&$orderby=Title", 0, new string[] { "ID" });
            }
            else
            {   // otherwise we get all columns from within the database by cycling through the rows.
                DataRow[] dbTables = repoData.Tables.Select("Database_NameId=" + dbID);
                for (int i = 0; i < dbTables.Length; i++)//  DataRow row in _dtTables.Rows)
                {
                    DataRow row = dbTables[i];
                    filter += " Table_NameId eq " + row["ID"] + " or "; // add the tableID to the filter

                    if (j++ >= 19 | i >= dbTables.Length-1) // we retrieve 20 sets of columns at a time for speed.
                    {
                        filter = "&$filter=" + filter.Substring(0, filter.Length - 3);
                        try
                        {
                            dt = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Columns')/items?" + filter, 0, new string[] { "ID" });
                            if (dtColumn == null) dtColumn = dt;
                            else if (dt != null) dtColumn.Merge(dt, true, MissingSchemaAction.Ignore);
                        }
                        catch (Exception e) { Debug.WriteLine(e.Message); }
                        filter = "";
                        j = 0;
                    }
                }
            }

            // Use Linq to add Table and DB parent fields to each record
            try
            {
                if (dtColumn == null) return null;
                dtColumn.Columns["Title"].ColumnName = "Column_Name";
                dtColumn.Columns.Add("Table_Name", Type.GetType("System.String"));
                dtColumn.Columns.Add("Database_NameId", Type.GetType("System.Int64"));
                dtColumn.Columns.Add("Database_Name", Type.GetType("System.String"));
                dtColumn.AsEnumerable().Join(repoData.Tables.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Table_NameId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o =>
                        {
                            o._dtmater.SetField("Table_Name", o._dtchild["Table_Name"].ToString());
                            o._dtmater.SetField("Database_NameId", (long)o._dtchild["Database_NameId"]);
                            o._dtmater.SetField("Database_Name", o._dtchild["Database_Name"].ToString());
                        }
                    );
            }
            catch (Exception e) { Debug.WriteLine(e.Message); }

            return dtColumn;
        }


        internal async Task<DataTable> GetRelationsFromSharepointByDB(long dbID, string Relation_Type = "All", bool ManyDirection = false)
        {
            DataTable dt = null, dtRelation = null;
            int j = 0; 
            string filter = "";

            // pull all tables from this particular database into our DataTable
            if (repoData.Tables == null)
                repoData.Tables = await GetTablesFromSharepoint(dbID);
            else if (repoData.Tables.Select("Database_NameId=" + dbID).Length == 0)
            {  
                dt = await GetTablesFromSharepoint(dbID);
                if (dt == null) return null;
                repoData.Tables.Merge(dt, true, MissingSchemaAction.Ignore);
            }
            DataRow[] dbTables = repoData.Tables.Select("Database_NameId=" + dbID);

            for (int i = 0; i < dbTables.Length; i++)//  DataRow row in _dtTables.Rows)
            {
                DataRow row = dbTables[i];
                filter = filter + (ManyDirection ? "Table_ManyId eq " : "Table_OneId eq ") + row["ID"] + " or ";

                if (j++ >= 19 | i >= dbTables.Length-1) // retrieve data for 20 tables at a time
                {
                    if( Relation_Type == "All")
                        filter = "&$filter=(" + filter.Substring(0, filter.Length - 3) + ") ";
                    else
                        filter = "&$filter=(" +  filter.Substring(0, filter.Length - 3) + ") and (Relation_Type eq '" + Relation_Type + "')";
                    try
                    {
                        dt = await retrieveDataTable(SPRepository + "/_api/lists/getbytitle('Relations')/items?" + filter, 0, new string[] { });
                        if (dtRelation == null)    dtRelation = dt;
                        else if (dt != null)       dtRelation.Merge(dt, true, MissingSchemaAction.Ignore);
                    }
                    catch (Exception e) { Debug.WriteLine(e.Message); }

                    filter = "";
                    j = 0;
                }
            }
            try
            {
                if (dtRelation == null) return null;
                // Use Linq to add Table and DB fields to each record
//                dtRelation.Columns["Title"].ColumnName = "Relation_Name";
                dtRelation.Columns.Add("Column_One", typeof(string));
                dtRelation.Columns.Add("Table_One", typeof(string));
                dtRelation.Columns.Add("Database_One", typeof(string));
                dtRelation.Columns.Add("Column_Many", typeof(string));
                dtRelation.Columns.Add("Table_Many", typeof(string));
                dtRelation.Columns.Add("Database_Many", typeof(string));


                dtRelation.AsEnumerable().Join(repoData.Tables.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Table_OneId"]), _dtchild => Convert.ToString(_dtchild["ID"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => { o._dtmater.SetField("Table_One", o._dtchild["Table_Name"].ToString()); o._dtmater.SetField("Database_One", o._dtchild["Database_Name"]); }
                    );

                dtRelation.AsEnumerable().Join(repoData.Tables.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Table_ManyId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => { o._dtmater.SetField("Table_Many", o._dtchild["Table_Name"].ToString()); o._dtmater.SetField("Database_Many", o._dtchild["Database_Name"]); }
                    );

                if(repoData.Columns != null)
                dtRelation.AsEnumerable().Join(repoData.Columns.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Column_OneId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => { o._dtmater.SetField("Column_One", o._dtchild["Column_Name"].ToString()); }
                    );

                if(repoData.Columns != null)
                dtRelation.AsEnumerable().Join(repoData.Columns.AsEnumerable(), _dtmater => Convert.ToString(_dtmater["Column_ManyId"]), _dtchild => Convert.ToString(_dtchild["id"]),
                    (_dtmater, _dtchild) => new { _dtmater, _dtchild }).ToList().ForEach(
                        o => { o._dtmater.SetField("Column_Many", o._dtchild["Column_Name"].ToString()); }
                    );

            }
            catch (Exception e) { Debug.WriteLine(e.Message); }

            return dtRelation;
        }




        // wrapper that ensures that we can retrieve as many rows as we like, given Sharepoint's 5000 record limit
        internal async Task<DataTable> retrieveDataTable(string appUrl, int iteration, string[] PrimaryKey)
        {
            const int iterSize = 5000;
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
                Debug.WriteLine(e.Message + "\n" + response);
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
            if (PrimaryKey != null)
            {
                var keys = new DataColumn[PrimaryKey.Length];
                for (int i = 0; i < PrimaryKey.Length; i++)
                    keys[i] = dataTable.Columns[PrimaryKey[i]];
                // I may add the Primary Key field back in later.   dataTable.PrimaryKey = keys;
            }
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

        
        
        internal bool isSharepointColumn(string columnList, string columnName)
        {
            // Some entities (such as SystemUser entity) are not useful for ER diagram. Skip those noisy not useful entities for visualization.
            return (columnList.Contains($",{columnName.ToLower()},"));
        }


        internal async Task<string> GetSharePointListFields(string listName, string nameType = "User")
        {
            var client = httpDownloadClient.client;
            client.DefaultRequestHeaders.Accept.Clear();
            System.Net.Http.Headers.MediaTypeWithQualityHeaderValue acceptHeader = System.Net.Http.Headers.MediaTypeWithQualityHeaderValue.Parse("application/json;odata.metadata=none");
            client.DefaultRequestHeaders.Accept.Add(acceptHeader);

            var response = await client.GetAsync(SPRepository + "/_api/lists/getbytitle('" + listName + "')//fields?$filter=Hidden eq false " + (nameType == "User" ? "and ReadOnlyField eq false" : ""));

            var responseText = await response.Content.ReadAsStringAsync();
            try { response.EnsureSuccessStatusCode(); }
            catch (Exception e)
            {   MessageBox.Show("Cannot find Sharepoint List " + listName + "\n" + e.Message + "\n" + response);
                return null;            }

            var schemaObject = JObject.Parse(responseText);
            var values = schemaObject["value"].ToArray();

            string columnList = ",";
            for (int i = 0; i < values.Length; i++)
                    if(nameType=="User")
                        if (values[i].SelectToken("InternalName").ToString() != "Title") // ignore the Title as we have renamed it in the UI
                            if (values[i].SelectToken("TypeAsString").ToString() != "Lookup") // ignore lookup fields as we have to do extra work
            if (values[i].SelectToken("Title").ToString() != "ID") // ignore the ID field
                                columnList = columnList + values[i].SelectToken("Title").ToString().ToLower() + ",";
                    else
                        columnList = columnList + values[i].SelectToken("InternalName").ToString() + ",";

            return columnList;
        }


        // Because Visio.DataRecordsets are not Linq compliant, we have to convert them into an XML stream
        // to move them into a more useful datatable structure.
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


        internal RecordStatus RowHasChanged(DataRow editRow, DataTable dataTable)
        {
            var repoRow = dataTable.Select("ID=" + editRow["ID"].ToString());
            if (repoRow.Length == 0) return RecordStatus.DoesNotExist;

            // compare the edited row with the original row. If Sharepoint column doesn't exist, ignore it as
            // that column won't be uploaded anyway.
            for (int i = 0; i<editRow.ItemArray.Length; i++)
            {
                string colName = editRow.Table.Columns[i].ColumnName;
                if (dtDataFieldExists(dataTable, colName))
                    if (editRow[i] != repoRow[0][colName])
                        return RecordStatus.ExistsAndHasChanged;
            }
            return RecordStatus.ExistsAndHasNotChanged;
        }

        internal bool FieldHasChanged(object editField, object repoField)
        {
            if (repoField == null)
                if (editField == DBNull.Value)
                    return false;
                else
                    return true;
            if (editField.ToString() == repoField.ToString())
                return false;
            else
                return true;
        }


        /* SaveApplicationToSharepoint requires the following case-sensitive fieldnames for an upload:
* ID (null for new ones), Database_Name, Application_NameId, Application_Name
*/
        internal void saveApplicationToSharepoint(DataTable dtApp, bool TestForExistence = false)
        {
            GetClientContext();
            var targetList = gclientContext.Web.Lists.GetByTitle("Applications");
            string sharepointTableFields = Task.Run(async () => await GetSharePointListFields("Applications")).Result;
            int j = 0;
            ListItem oItem = null;
            DataRow repoRow = null;

            DataTable repoTable = repoData.Applications;

            if (dtApp is null) return; // nothing to save
            if (repoTable is null) repoTable = Task.Run(async () => await GetApplicationsFromSharepoint("", new string[] { })).Result;

            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Tables: " + j + " of " + dtApp.Rows.Count, "ScopeStart");

            foreach (DataRow row in dtApp.Rows)
            {
                j++;
                if (row["Edit_Status"].ToString() != "New") continue;

                // Choose whether to add or edit the data row
                if (row["ID"] == DBNull.Value)
                {   // verify and choose add
                    if (row["Application_Name"].ToString() == "") continue;
                    if (repoTable != null)
                        if(repoTable.Select("Application_Name = '" + row["Application_Name"].ToString() + "'").Length > 0) continue;  // application exists, skip it
                    oItem = targetList.AddItem(new ListItemCreationInformation());
                    repoRow = null;

                    oItem["Title"] = row["Application_Name"].ToString();
                }
                else
                {   // we are editing and the record doesn't exist or has not changed, skip
                    if (RowHasChanged(row, repoData.Tables) != RecordStatus.ExistsAndHasChanged) continue;
                    repoRow = repoTable.Select("ID=" + row["ID"].ToString())[0];
                    // Continue, as the record exists and has changed.
                    oItem = targetList.GetItemById((int)row["ID"]);
                }

                // Copy all columns from that are Sharepoint-valid 
                foreach (DataColumn column in dtApp.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (isSharepointColumn(sharepointTableFields, fieldName))
                        if (repoRow == null)
                        {
                            if (row[column.ColumnName] != DBNull.Value)
                                oItem[fieldName] = row[column.ColumnName].ToString();
                        }
                        else if (FieldHasChanged(row[column.ColumnName], repoRow[column.ColumnName]))
                            oItem[fieldName] = row[column.ColumnName].ToString();
                    //if (isSharepointColumn(sharepointTableFields, fieldName))
                    //    if (FieldHasChanged(row[column.ColumnName], dtDataFieldExists(repoTable, column.ColumnName) ? repoRow[column.ColumnName] : DBNull.Value))
                    //        oItem[fieldName] = row[column.ColumnName].ToString();
                }

                // update sharepoint
                oItem.Update();
                row["Edit_Status"] = "Uploading";

                Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Application : " + j + " of " + dtApp.Rows.Count);
                gclientContext.ExecuteQuery();
            }
            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading " + j + " of " + dtApp.Rows.Count, "ScopeEnd");
        }


        /* SaveDatabaseToSharepoint requires the following case-sensitive fieldnames for an upload:
    * ID (null for new ones), Database_Name, Application_NameId, Application_Name
    */
        internal void saveDBToSharepoint(DataTable dbTable, bool TestForExistence = false)
        {
            var targetList = gclientContext.Web.Lists.GetByTitle("Databases");
            if (repoData.Databases == null) repoData.Databases = Task.Run(async () => await GetDBsFromSharepoint(0)).Result;
            if (repoData.Applications == null) repoData.Applications = Task.Run(async () => await GetApplicationsFromSharepoint("", new string[] { })).Result;

            string sharepointTableFields = Task.Run(async () => await GetSharePointListFields("Databases")).Result;
            int i = 0, j = 0;
            ListItem oItem = null;
            DataRow repoRow = null;
            DataTable repoTable = repoData.Databases;

            if (dbTable is null) return; // nothing to save
            if (repoTable is null) repoTable = Task.Run(async () => await GetDBsFromSharepoint(0)).Result;

            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Tables: " + j + " of " + dbTable.Rows.Count, "ScopeStart");

            foreach (DataRow row in dbTable.Rows)
            {
                j++;
                if (row["Edit_Status"].ToString() != "New") continue;

                if (row["ID"] == DBNull.Value)
                {   // Add new record
                    if (row["Database_Name"].ToString() == "") continue;
                    if (repoTable != null)
                        if (repoTable.Select("Database_Name = '" + row["Database_Name"].ToString() + "'").Length > 0) continue;  // application exists, skip it
                    repoRow = null;
                    oItem = targetList.AddItem(new ListItemCreationInformation());

                    oItem["Title"] = row["Database_Name"].ToString();
                    FieldLookupValue fk = new FieldLookupValue();
                    if (row["Application_NameId"] == DBNull.Value)
                        fk.LookupId = Convert.ToInt32( repoData.Applications.Select("Application_Name='" + row["Application_Name"] + "'")[0]["ID"]);
                    else
                        fk.LookupId = Convert.ToInt32(row["Application_NameId"]);
                    oItem["Application_Name"] = fk;// row["Application_Name"].ToString();
                }
                else 
                {   // edit record
                    if (RowHasChanged(row, repoTable) != RecordStatus.ExistsAndHasChanged) continue; // nothing to change
                    repoRow = repoTable.Select("ID=" + row["ID"].ToString())[0];
                    oItem = targetList.GetItemById((int)row["ID"]);
                }

                // Copy all columns from that are Sharepoint-valid 
                foreach (DataColumn column in dbTable.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (fieldName == "Database_Name") continue;
                    if (isSharepointColumn(sharepointTableFields, fieldName))
                        if (repoRow == null)
                        {
                            if (row[column.ColumnName] != DBNull.Value)
                                oItem[fieldName] = row[column.ColumnName].ToString();
                        }
                        else if (FieldHasChanged(row[column.ColumnName], repoRow[column.ColumnName]))
                            oItem[fieldName] = row[column.ColumnName].ToString();
                }

                // update sharepoint
                oItem.Update();
                row["Edit_Status"] = "Uploading";

                if (i++ >= 0) { // write every line individually
                    Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Database: " + j + " of " + dbTable.Rows.Count);
                    gclientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                gclientContext.ExecuteQuery();
            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading " + j + " of " + dbTable.Rows.Count, "ScopeEnd");
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
            DataRow repoRow = null;

            DataTable repoTable = repoData.Tables;

            if (dTable is null) return; // nothing to save
            if (repoData.Databases is null) repoData.Databases = Task.Run(async () => await GetDBsFromSharepoint(0)).Result;

            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Tables: " + j + " of " + dTable.Rows.Count, "ScopeStart");

            foreach (DataRow row in dTable.Rows)
            {
                j++;
                if (row["Edit_Status"].ToString() != "New") continue;

                // this line is just in case there is no repository data information
                if (repoData.Tables == null)
                {
                    var dRow = repoData.Databases.Select("Database_Name='" + row["Database_Name"] + "'");
                    repoData.Tables = Task.Run(async () => await GetTablesFromSharepoint((long)dRow[0]["Id"])).Result;
                }

                // Choose whether to add or edit the data row
                if (row["ID"] == DBNull.Value)
                {   // verify and choose add
                    if (row["Table_Name"].ToString() == "") continue;
                    if (repoData.Tables != null)
                        if (repoData.Tables.Select("Table_Name='" + row["Table_Name"] + "' and Database_Name = '" + row["Database_Name"] + "'").Length>0) continue;  // application exists, skip it

                    // Add the GUID
                    DataRow[] dRow = repoData.Databases.Select("[Database_Name] = '" + row["Database_Name"].ToString() + "'");// +  "' OR ID=" + row["Database_NameId"].ToString());
                    oItem = targetList.AddItem(new ListItemCreationInformation());
                    oItem["TableGUID"] = GenerateHash(dRow[0]["Application_Name"].ToString() + "." + dRow[0]["Database_Name"].ToString() + "." + row["Table_Name"].ToString());
                    repoRow = null;

                    // get the foreign (Database) key
                    FieldLookupValue fk = new FieldLookupValue();
                    fk.LookupId = Convert.ToInt32( dRow[0]["Id"]);
                    oItem["Database_Name"] = fk;// row["Application_Name"].ToString();
                }
                else
                {   // we are editing and the record doesn't exist or has not changed, skip
                    if (RowHasChanged(row, repoTable) != RecordStatus.ExistsAndHasChanged) continue;
                    repoRow = repoTable.Select("ID=" + row["ID"].ToString())[0];
                    // Continue, as the record exists and has changed.
                    oItem = targetList.GetItemById(Convert.ToInt32( row["ID"]));
                }

                // Copy all columns from that are Sharepoint-valid 
                foreach ( DataColumn column in dTable.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_"); 
                    if (isSharepointColumn(sharepointTableFields, fieldName))
                        if (repoRow == null)
                        {
                            if (row[column.ColumnName] != DBNull.Value)
                                oItem[fieldName] = row[column.ColumnName].ToString();
                        }
                        else if (FieldHasChanged(row[column.ColumnName], repoRow[column.ColumnName]))
                            oItem[fieldName] = row[column.ColumnName].ToString();
                }

                oItem["Title"] = row["Database_Name"].ToString() + "." + row["Table_Name"].ToString(); // Title is "Full_Table_Name

                // update sharepoint
                oItem.Update();
                row["Edit_Status"] = "Uploading";
                
                if (i++ >= 0)
                {
                    Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Tables: " + j + " of " + dTable.Rows.Count);
                    gclientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                gclientContext.ExecuteQuery();
            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading " + j + " of " + dTable.Rows.Count, "ScopeEnd");
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
            DataTable tableDataTable = null;
            if (dbID > 0)
                tableDataTable = Task.Run(async () => await GetTablesFromSharepoint(dbID)).Result;
            else
                tableDataTable = repoData.Tables;

                ListItem oItem = null;

                var targetList = gclientContext.Web.Lists.GetByTitle("Columns");
                string sharepointColumnFields = Task.Run(async () => await GetSharePointListFields("Columns")).Result;
                int i = 0, j = 0;
            DataRow repoRow = null;

            // get all columns for this particular database
            if (repoData.Columns == null & dbID > 0)
                repoData.Columns = Task.Run(async () => await GetColumnsFromSharepoint(dbID, 0)).Result;
            DataTable repoTable = repoData.Columns;

            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Columns: " + j + " of " + dtColumns.Rows.Count, "ScopeStart");
            foreach (DataRow row in dtColumns.Rows)
            {
                j++;
                if (row["Edit_Status"].ToString() != "New") continue;
                bool AddNew = (row["ID"].ToString() == "");

                var dRow = repoData.Databases.Select("Database_Name='" + row["Database_Name"] + "'");

                // Add if there's no ID, otherwise edit
                if (row["ID"] == DBNull.Value)
                {   // If we are trying to add and the record already exists, skip
                    if (row["Column_Name"].ToString() == "") continue;
                    if (repoTable != null)   
                        if(repoTable.Select("Database_Name='" + row["Database_Name"] + "' and Table_Name='" + row["Table_Name"] + "' and Column_Name='" + row["Column_Name"] + "'").Length > 0) continue;
                    repoRow = null;
                    oItem = targetList.AddItem(new ListItemCreationInformation());
                    
                    // set our GUID
                    var dr = tableDataTable.Select("Full_Table_Name = '" + row["Database_Name"] + "." + row["Table_Name"] + "'");
                    if (dr.Length == 0) { MessageBox.Show("Cannot find table " + row["Table_Name"].ToString() + ". \n Skipping insertion of " + row["Column_Name"].ToString()); continue; }
                    oItem["ColumnGUID"] = GenerateHash(dr[0]["Application_Name"].ToString() + "." + dr[0]["Database_Name"].ToString() + "." + dr[0]["Table_Name"].ToString() + "." + row["Column_Name"].ToString());

                    // get the foreign (Table) key
                    FieldLookupValue fk = new FieldLookupValue();
                    fk.LookupId = Convert.ToInt32(dr[0]["Id"]);
                    oItem["Table_Name"] = fk;// row["Application_Name"].ToString();
                    oItem["Title"] = row["Column_Name"].ToString();
                }
                else
                {   // we are editing and the record doesn't exist or has not changed, skip
                    if (RowHasChanged(row, repoTable) != RecordStatus.ExistsAndHasChanged) continue;
                    repoRow = repoTable.Select("ID=" + row["ID"].ToString())[0];
                    // Continue, as the record exists and has changed.
                    oItem = targetList.GetItemById((int)row["ID"]);
                }

                // Copy all columns from that are Sharepoint-valid 
                foreach (DataColumn column in dtColumns.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (isSharepointColumn(sharepointColumnFields, fieldName))
                        if (repoRow == null)
                        {
                            if (row[column.ColumnName] != DBNull.Value)
                                oItem[fieldName] = row[column.ColumnName].ToString();
                        }
                        else if (FieldHasChanged(row[column.ColumnName], repoRow[column.ColumnName]))
                            oItem[fieldName] = row[column.ColumnName].ToString();
                }

                // update sharepoint
                oItem.Update();
                row["Edit_Status"] = "Uploading";
                if (i++ >= 49)
                {
                    Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Columns: " + j + " of " + dtColumns.Rows.Count);
                    gclientContext.ExecuteQuery();
                    i = 0;
                }
            }
                if (i > 0)
                    gclientContext.ExecuteQuery();
            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Columns: " + j + " of " + dtColumns.Rows.Count, "ScopeEnd");
            //}
            //catch(Exception e)
            //{
            //    MessageBox.Show("Error: " + e.Message + "\n\n Retrying once");
            //    return false;
            //}

            return true;
        }

        internal bool getColumnIDFromDatatable(ref long? ColumnID, string ColumnName, ref long? TableID, string TableName)
        {
            // Get the TableID
            if (TableID == 0) {
                var tableRet = repoData.Tables.Select("Table_Name='" + TableName.ToString() + "'");
                if (tableRet == null) return false; // there is no table by that name
                TableID = (long)tableRet[0]["ID"];
            }
            long tableNonRef = TableID ?? 0; 

            if (TableID == 0)                   return false; // If we didn't find the Table above, just return false;
            if (ColumnID > 0 & TableID > 0)     return true; // everything is fine, so we can continue;
            if (ColumnName == "")               return true; // we're not looking for a column

            // Get the columnID
            if (ColumnID == 0)
            {
                if (repoData.Columns is null) // if DTData.Columns is null, let's initialise it just for the sake of it
                    repoData.Columns = Task.Run(async () => await GetColumnsFromSharepoint(0, tableNonRef)).Result;

                var ret = repoData.Columns.Select("[Column_Name]='" + ColumnName + "' and [Table_NameId]='" + TableID + "'");
                if (ret.Length > 0)
                {
                    ColumnID = (long)ret[0]["ID"];
                    return true;
                }

                //If we can't find the ColumnID, let's try importing all columns in that table from Sharepoint
                   repoData.Columns.Merge(Task.Run(async () => await GetColumnsFromSharepoint(0, tableNonRef)).Result);

                    ret = repoData.Columns.Select("[Column_Name]='" + ColumnName + "' and [Table_NameId]='" + TableID + "'");
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
            DataTable tableDataTable = Task.Run(async () => await GetTablesFromSharepoint(dbID)).Result;
            repoData.Relations = Task.Run(async () => await GetRelationsFromSharepointByDB((long)tableDataTable.Rows[0]["Database_NameId"])).Result;
            if (dbID != dbManyID)
            {
                DataTable tbl = Task.Run(async () => await GetTablesFromSharepoint(dbManyID)).Result;
                tableDataTable.Merge(tbl,true, MissingSchemaAction.Ignore);
                repoData.Relations = Task.Run(async () => await GetRelationsFromSharepointByDB((long)tbl.Rows[0]["Database_NameId"])).Result;
            }

            var targetList = gclientContext.Web.Lists.GetByTitle("Relations");
            string sharepointRelationFields = Task.Run(async () => await GetSharePointListFields("Relations")).Result;
            ListItem oItem = null;
            int i = 0, j = 0;
            long? colOneID=0, colManyID=0, tabOneID=0, tabManyID=0;
            DataRow repoRow = null;

            DataTable repoTable = repoData.Relations;

            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Relations: " + j + " of " + dtRelation.Rows.Count, "ScopeStart");
            foreach (DataRow row in dtRelation.Rows)
            {
                j++;
                if (row["Edit_Status"].ToString() == "") continue;
                // Add if there's no ID, otherwise edit
                if ( row["ID"] == DBNull.Value)
                {   // If we are trying to add and the record already exists, skip
                    //var repoRow = repoData.Relations.Select("Column_One='" + row["Column_One"] + "'and Column_Many='" + row["Column_Many"] + "' and Table_One_Full='" + row["Database_One"] + "." + row["Table_One"] + "'and Table_Many_Full='" + row["Database_Many"] + "." + row["Table_Many"] + "'");
                    //if (repoRow.Length > 0) continue; // the item exists and should not be re-uploaded
                    oItem = targetList.AddItem(new ListItemCreationInformation());
                    repoRow = null;
                    colOneID = (long?) row["Column_OneId"];
                    colManyID = (long?)row["Column_ManyId"];
                    tabOneID = (long?)row["Table_OneId"];
                    tabManyID = (long?)row["Table_ManyId"];
                    // assign the column and table IDs
                    if (!getColumnIDFromDatatable(ref colOneID, row["Column_One"].ToString(), ref tabOneID, row["Table_One"].ToString())) continue;
                    if (!getColumnIDFromDatatable(ref colManyID, row["Column_Many"].ToString(), ref tabManyID, row["Table_Many"].ToString())) continue;
                    oItem["Table_One"] = tabOneID;
                    oItem["Table_Many"] = tabManyID;
                    if (colOneID > 0) oItem["Column_One"] = colOneID;
                    if (colManyID > 0) oItem["Column_Many"] = colManyID;
                }
                else
                {   // we are editing and the record doesn't exist or has not changed, skip
                    if (RowHasChanged(row, repoTable) != RecordStatus.ExistsAndHasChanged) continue;
                    repoRow = repoTable.Select("ID=" + row["ID"].ToString())[0];
                    // Continue, as the record exists and has changed.
                    oItem = targetList.GetItemById((int)row["ID"]);
                }

                if(row["Relation_Type"].ToString() == "Foreign Key")
                    if (Globals.ThisAddIn.drawingManager.IsEntityInSkipList(row["Table_One"].ToString()))  continue; //we want to keep skip list columns but remove their relations

                // Copy all columns from that are Sharepoint-valid 
                foreach (DataColumn column in dtRelation.Columns)
                {   // replace any Visio- or Sharepoint-styled strings and check to see if it's still valid
                    string fieldName = column.ColumnName.Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (isSharepointColumn(sharepointRelationFields, fieldName))
                        if (repoRow == null)
                        {
                            if (row[column.ColumnName] != DBNull.Value)
                                oItem[fieldName] = row[column.ColumnName].ToString();
                        }
                        else if (FieldHasChanged(row[column.ColumnName], repoRow[column.ColumnName]))
                            oItem[fieldName] = row[column.ColumnName].ToString();
                }

                // update sharepoint
                oItem.Update();
                row["Edit_Status"] = "Uploading";
                if (i++ >= 39)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading Relations: " + j + " of " + dtRelation.Rows.Count);
                    gclientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                gclientContext.ExecuteQuery();
            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Uploading Relations: " + j + " of " + dtRelation.Rows.Count, "ScopeEnd");
        }

        internal bool dtDataFieldExists(DataTable dt, string fieldName)
        {   if (dt == null) return false;
            for (int i = 0; i < dt.Columns.Count; i++)
                if (dt.Columns[i].ColumnName == fieldName)
                    return true;
            return false;
        }

        internal DataTable convertDataRecToDataTable(Visio.DataRecordset dr)
        {
            if (dr == null) return null;

            int rows = dr.GetDataRowIDs("").Length;
            int cols = dr.DataColumns.Count;
            DataTable dataTable = new DataTable();

            for (int col = 0; col <cols; col++)
                dataTable.Columns.Add(dr.GetRowData(0).GetValue(col).ToString());


            for (int row = 1; row <= rows; row++)
            {
                DataRow newRow = dataTable.NewRow();
                for (int col = 0; col < cols; col++)
                    newRow[col] = dr.GetRowData(row).GetValue(col).ToString();

                dataTable.Rows.Add(newRow);
            }
            return dataTable;
        }


        internal void UploadCompleteDataRecordsetToSharepoint()
        {
            GetClientContext();
            if (!AssignDataRecordsets()) return;
            repoData.Databases = Task.Run(async () => await GetDBsFromSharepoint(0)).Result;
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
                    dbName = dtTables.Rows[0]["Database_Name"].ToString();
                if (dbName == "") // Only assign it this name if it doesn't already exist - change this to drop-down inputbox
                    dbName = "Dynamics OData";

                var dbRow = repoData.Databases.Select("Database_Name='" + dbName + "'");
                if(dbRow is null)
                    { MessageBox.Show("There is no database by that name."); return;}
                dbID = (long) dbRow[0]["Id"];

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

        internal void CopySharepointToExcel()
        {
            GetClientContext();

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath) || !System.IO.Directory.Exists(folderBrowserDialog.SelectedPath))
                if (DialogResult.OK != folderBrowserDialog.ShowDialog()) return;

            if (!System.IO.Directory.Exists(folderBrowserDialog.SelectedPath))
            {
                MessageBox.Show( $"{folderBrowserDialog.SelectedPath} does not exist.");
                return;
            }

            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Downloading Repository to Excel", "ScopeStart");
            int i = 0;
            repoData.Databases = Task.Run(async () => await (GetDBsFromSharepoint(0))).Result;
            foreach(DataRow row in repoData.Databases.Rows)
            {
                i++;
                Utilites.ScreenEvents.DisplayVisioStatusPersistent("Retrieving Tables from " + row["Database_Name"].ToString() + "  (" + i + " of " + repoData.Databases.Rows.Count + ")");
                // get tables, columns, and relations by DatabaseID
                DataTable tmp = null;
                tmp = Task.Run(async () => await (GetTablesFromSharepoint((long)row["ID"]))).Result;
                if (repoData.Tables == null) repoData.Tables = tmp;
                else if (tmp != null) repoData.Tables.Merge(tmp, true, MissingSchemaAction.Ignore);

                Utilites.ScreenEvents.DisplayVisioStatusPersistent("Retrieving Columns from " + row["Database_Name"].ToString() + "  (" + i + " of " + repoData.Databases.Rows.Count + ")");
                tmp = Task.Run(async () => await (GetColumnsFromSharepoint((long)row["ID"], 0))).Result;
                if (repoData.Columns == null) repoData.Columns = tmp;
                else if (tmp != null) repoData.Columns.Merge(tmp, true, MissingSchemaAction.Ignore);

                Utilites.ScreenEvents.DisplayVisioStatusPersistent("Retrieving Relations from " + row["Database_Name"].ToString() + "  (" + i + " of " + repoData.Databases.Rows.Count + ")");
                tmp = Task.Run(async () => await (GetRelationsFromSharepointByDB((long) row["ID"],"All", true))).Result;
                if (repoData.Relations == null) repoData.Relations = tmp;
                else if (tmp != null) repoData.Relations.Merge(tmp, true, MissingSchemaAction.Ignore);
            }


            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Writing Data to Excel");

            XL.Application xl = new XL.Application();
            xl.Visible = true;
            var wb = xl.Workbooks.Add();

            WriteTableToExcel(wb, repoData.Applications, "Applications");
            WriteTableToExcel(wb, repoData.Databases, "Databases");
            WriteTableToExcel(wb, repoData.Tables, "Tables");
            WriteTableToExcel(wb, repoData.Columns, "Columns");
            WriteTableToExcel(wb, repoData.Relations, "Relations");

            Utilites.ScreenEvents.DisplayVisioStatusPersistent("Complete", "ScopeEnd");

        }

        internal void UploadExceltoSharepoint()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            DialogResult result = openFileDialog.ShowDialog();
            if (result != DialogResult.OK) return;


            XL.Application xl = new XL.Application();
            xl.Visible = true;
            XL.Workbook wb =  xl.Workbooks.Open(openFileDialog.FileName);

            editData = new DTData();

            editData.Applications = ImportAndUpload(wb, "Applications");
            repoData.Applications = null;
            saveApplicationToSharepoint(editData.Applications);

            editData.Databases = ImportAndUpload(wb, "Databases");
            repoData.Applications = null;
            repoData.Databases = null;
            saveDBToSharepoint(editData.Databases);


            // refresh all Tables and Databases before trying to upload more tables
            repoData.Databases =  Task.Run(async () => await (GetDBsFromSharepoint(0))).Result;
            repoData.Tables = null;
            foreach (DataRow row in repoData.Databases.Rows)
            {
                // get tables, columns, and relations by DatabaseID
                DataTable tmp = null;
                tmp = Task.Run(async () => await (GetTablesFromSharepoint((long)row["ID"]))).Result;
                if (repoData.Tables == null) repoData.Tables = tmp;
                else if (tmp != null) repoData.Tables.Merge(tmp, true, MissingSchemaAction.Ignore);
            }

            editData.Tables = ImportAndUpload(wb, "Tables");
            saveTablesToSharepoint(editData.Tables);

            // refresh all Tables and Columns before trying to upload more columns
            repoData.Tables = null;
            repoData.Columns = null;
            foreach (DataRow row in repoData.Databases.Rows)
            {
                // get tables, columns, and relations by DatabaseID
                DataTable tmp = null;
                tmp = Task.Run(async () => await (GetTablesFromSharepoint((long)row["ID"]))).Result;
                if (repoData.Tables == null) repoData.Tables = tmp;
                else if (tmp != null) repoData.Tables.Merge(tmp, true, MissingSchemaAction.Ignore);
                tmp = Task.Run(async () => await (GetColumnsFromSharepoint((long)row["ID"], 0))).Result;
                if (repoData.Columns == null) repoData.Columns = tmp;
                else if (tmp != null) repoData.Columns.Merge(tmp, true, MissingSchemaAction.Ignore);
            }

            editData.Columns = ImportAndUpload(wb,"Columns");
            saveColumnsToSharepoint(editData.Columns,0);

            // get all column information for upload
            repoData.Columns = null;
            foreach (DataRow row in repoData.Databases.Rows)
            {
                // get tables, columns, and relations by DatabaseID
                DataTable tmp = null;
                tmp = Task.Run(async () => await (GetColumnsFromSharepoint((long)row["ID"], 0))).Result;
                if (repoData.Columns == null) repoData.Columns = tmp;
                else if (tmp != null) repoData.Columns.Merge(tmp, true, MissingSchemaAction.Ignore);
            }

            editData.Relations = ImportAndUpload(wb,"Relations");
            repoData.Relations = null;
            saveRelationsToSharepoint(editData.Relations,0,0);
        }



        private DataTable ImportAndUpload(XL.Workbook wb, string SheetName)
        {
            XL.Worksheet ws = null;
            try
            {
                ws = wb.Sheets[SheetName];
            }
            catch
            {
                // it's ok if the sheetname doesn't exist as we may only be uploading one type of entity.
                Debug.WriteLine(SheetName + " does not exist in workbook " + wb.Name.ToString());
                return null;
            }

            ws.Activate();
            ws.Cells[2][2].Select();
            XL.Range rng = ws.Application.Selection.CurrentRegion;
            var MyArray = rng.Value;

            if (MyArray == null) return null;

            int rows = MyArray.GetLength(0);
            int cols = MyArray.GetLength(1);
            DataTable dataTable = new DataTable();

            for (int col = 1; col <= cols; col++)
                dataTable.Columns.Add(MyArray[1,col].ToString());
            if (!dtDataFieldExists(dataTable, "Edit_Status")) dataTable.Columns.Add("Edit_Status", typeof(string));


            for (int row = 2; row <= rows; row++)
            {
                DataRow newRow = dataTable.NewRow();
                for (int col = 2; col <= cols; col++)
                    if (MyArray[row, col] != null)
                        newRow[col-1] = MyArray[row, col].ToString();

                newRow["Edit_Status"] = "New";
                dataTable.Rows.Add(newRow);
            }
            return dataTable;
        }



        private void WriteTableToExcel(XL.Workbook wb, DataTable dt, string sheetName)
        {
            XL.Worksheet ws = wb.Worksheets.Add();
            ws.Name = sheetName;
            if (dt == null) return;
            // Remove the nuisance Sharepoint columns
            foreach (string s in new string[] {"FileSystemObjectType","ServerRedirectedEmbedUri","ServerRedirectedEmbedUrl","ContentTypeId","ComplianceAssetId","Created","AuthorId","OData__UIVersionString","Attachments", "GUID" })
                if (dtDataFieldExists(dt, s)) dt.Columns.Remove(s);

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
            object[,] Cells = new object[RowsCount, ColumnsCount + 1];

            //string lastRow = "";
            for (int j = 0; j < RowsCount; j++)
                for (int i = 0; i < dt.Columns.Count; i++)
                    Cells[j, i] = (dt.Rows[j][i].ToString().StartsWith("=") ? "'" + dt.Rows[j][i] : dt.Rows[j][i]);

            ws.get_Range((Microsoft.Office.Interop.Excel.Range)(ws.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(ws.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;
            ws.Cells[2][2].Select();
            wb.Application.ActiveWindow.FreezePanes = true;
            wb.Application.Selection.AutoFilter();

            //XL.Range sel = xl.Selection;
            //sel.Subtotal(1, XL.XlConsolidationFunction.xlCount, new int[] { 2, 32 }, true, false, XL.XlSummaryRow.xlSummaryAbove);
            //ws.Columns["B:B"].EntireColumn.AutoFit();

        }

    }
}

