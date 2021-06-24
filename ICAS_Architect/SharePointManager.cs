using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Data;
using System.Security.Cryptography;
using System.Windows.Forms;
using VA = VisioAutomation;
using Visio = Microsoft.Office.Interop.Visio;
using SP = Microsoft.SharePoint.Client;

namespace ICAS_Architect
{
    internal class SharepointManager
    {
        internal HttpDownloadClient httpDownloadClient = null;
        string SPRepository = null;
        Visio.Application vApplication = null;
        ClientContext gclientContext = null;

        private Visio.DataRecordset applicationRecordset = null;
        private Visio.DataRecordset databaseRecordset = null;
        private Visio.DataRecordset tableRecordset = null;
        private Visio.DataRecordset columnRecordset = null;
        private Visio.DataRecordset relationRecordset = null;
        string applicationColumns = "";
        string databaseColumns = "";
        string tableColumns = "";
        string columnColumns = "";
        string relationColumns = "";


        internal SharepointManager()
        {
            vApplication = Globals.ThisAddIn.Application;
        }

      

        internal ClientContext GetSPAccessToken(bool WarnIfTablesDontExist=true)
        {
            string SPRepositoryFromRegistry = Globals.ThisAddIn.registryKey.GetValue("SharepointName", "https://icas1854.sharepoint.com/sites/Architecture").ToString();
            SPRepository = SPRepositoryFromRegistry;
            string ignore = "-1";
            DialogResult result = Utilites.ScreenEvents.ShowInputDialog(ref SPRepository, ref ignore,  "Site", "ignore", "Enter SharePoint site URL");
            if (result != DialogResult.OK) return null;

            if (SPRepository != SPRepositoryFromRegistry)
            {
                Globals.ThisAddIn.registryKey.SetValue("SharepointName", SPRepository);
            }

            if (httpDownloadClient == null)
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


        internal void LoadTablesList(string appName = "All", string dBName = "All", bool includeViews = true, bool includeAPIs = true)
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
            if (gclientContext is null) gclientContext = this.GetSPAccessToken();
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


        internal ListItemCollection getTableDatatable(string appName, string dBName, string[] fieldList)
        {
            string connString = CreateSharepointConnectionString("Tables", "All Items");
            SP.List oList = gclientContext.Web.Lists.GetByTitle("Tables");
            SP.View oView = oList.GetViewByName("All Items");

            CamlQuery camlQuery = new CamlQuery();
//            camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>     "<Value Type='Number'>10</Value></Geq></Where></Query><RowLimit>100</RowLimit></View>";
            string camlQ = "<Query><Where><Eq><FieldRef Name='Database'/><Value Type='Text'>" + dBName + "</Value></Eq></Where></Query>";

            string camlFields = "";
            for (int i = 0; i < fieldList.Length; i++)
                camlFields = camlFields + "<FieldRef Name='" + fieldList[i] + "'/>";
            if (camlFields.Length > 1) camlFields = "<ViewFields>" + camlFields + "</ViewFields>";

            camlQuery.ViewXml = "<View>" + camlQ + camlFields + "</View>";

            ListItemCollection collListItem = oList.GetItems(camlQuery);

            gclientContext.Load(collListItem);
            gclientContext.ExecuteQuery();
            return collListItem;
        }


        internal Visio.DataRecordset getTableDataRecordset(string appName = "All", string dBName = "All", bool includeViews = true, bool includeAPIs = true)
        {
            string whereCond = null;
            if (appName != "All")
            {
                whereCond = whereCond + (whereCond == null ? " WHERE " : " AND ");
                whereCond = whereCond + "[Application] = '" + appName + "'";
            }
            if (dBName != "All")
            {
                whereCond = whereCond + (whereCond == null ? " WHERE " : " AND ");
                whereCond = whereCond + "[Database] = '" + dBName + "'";
            }
            if (!includeViews)
            {
//                whereCond = whereCond + whereCond == null ? " WHERE " : " AND ";
                whereCond = whereCond + "[Table Type] <> 'View'";
            }
            if (!includeAPIs)
            {
//                whereCond = whereCond + whereCond == null ? " WHERE " : " AND ";
                whereCond = whereCond + "[Table Type] <> 'API'";
            }
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
            if (dBName != "All")
            {
                whereCond = whereCond == null ? " WHERE " : " AND ";
                whereCond = whereCond + "[Database] = '" + dBName + "'";
            }

            DeleteDataRecordset("Columns");
            string connString = CreateSharepointConnectionString("Columns");
            columnRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connString, "Select * from [Columns];", 0, "Columns");  // Select * from [Tables (All Items)] where [Application] = 'EDW'
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
            using (var clientContext = GetSPAccessToken(false))      //.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password))
            {
                Web web = clientContext.Web;
                clientContext.Load(web);
                clientContext.Load(web.Lists);
                clientContext.ExecuteQueryRetry();

                if (!clientContext.Web.ListExists("Application"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Applications", true, true, string.Empty, true);
                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Title = "UniqueID";
                    oField.StaticName = "UniqueID";
                    oField.Update();

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Application' Name='Application'/>", true, AddFieldOptions.AddToDefaultContentType);
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


                if (!clientContext.Web.ListExists("Database"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Databases", true, true, string.Empty, true);
                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Title = "UniqueID";
                    oField.StaticName = "UniqueID";
                    oField.Update();

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Database' Name='Database'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ApplicationUniqueID' Name='ApplicationUniqueID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Application' Name='Application'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='DBMS_Type' Name='DBMS_Type'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Version' Name='Version'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ServerUniqueID' Name='ServerUniqueID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Server_Name' Name='Server_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ExternalID' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityListURL' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityDefinitionsPath' Name='EntityDefinitionsPath'/>", true, AddFieldOptions.AddToDefaultContentType);
                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }

                
                if (!clientContext.Web.ListExists("Tables"))
                {
                    List list = clientContext.Web.CreateList( ListTemplateType.GenericList, "Tables", true, true, string.Empty,  true);
                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Title = "UniqueID";
                    oField.StaticName = "UniqueID";
                    oField.Update();

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Table_Name' Name='TableName'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Database' Name='Database'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Table_Type' Name='TableType'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Description' Name='Description'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Application' Name='Application'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Display_Name' Name='DisplayName'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Schema' Name='Schema'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Set_Name' Name='SetName'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ExternalID' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityListURL' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityDefinitionsPath' Name='EntityDefinitionsPath'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ApplicationUniqueID' Name='ApplicationUniqueID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='DatabaseUniqueID' Name='DatabaseUniqueID'/>", true, AddFieldOptions.AddToDefaultContentType);
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
                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Title = "UniqueID";
                    oField.Update();

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Column_Name' Name='Column_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Description' Name='Description'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Ordinal_Position' Name='Ordinal_Position'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Table_Name' Name='TableName'/>", true, AddFieldOptions.AddToDefaultContentType);
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
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='TableUniqueID' Name='Collation_Name'/>", true, AddFieldOptions.AddToDefaultContentType);
                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }


                if (!clientContext.Web.ListExists("Relations"))
                {
                    List list = clientContext.Web.CreateList(ListTemplateType.GenericList, "Relations", false, true, string.Empty, true);
                    Field oField = list.Fields.GetByInternalNameOrTitle("Title");
//                    oField.EnforceUniqueValues = true;
                    oField.EnableIndex();
                    oField.Title = "UniqueID";
                    oField.Update();

                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Relation_Name' Name='RelationName'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ColumnOne' Name='ColumnOne'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ColumnMany' Name='ColumnMany'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityOne' Name='TableOne'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityMany' Name='TableMany'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Connection_Type' Name='ConnectionType'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='Intersect_Entity' Name='Intersect_Entity'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityOneType' Name='EntityOneType'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EntityToType' Name='EntityToType'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ColumnOneID' Name='PrimaryColumnGUID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ColumnManyID' Name='ForeignColumnGUID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='TableOneID' Name='PrimaryTableGUID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='TableManyID' Name='ForeignTableGUID'/>", true, AddFieldOptions.AddToDefaultContentType);
                    list.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='ExternalID' Name='EntityListURL'/>", true, AddFieldOptions.AddToDefaultContentType);
                    clientContext.Load(list.Fields);
                    clientContext.ExecuteQuery();
                }
            }
        }


        internal void UploadToSharePoint()
        {
            AssignDataRecordsets();
            //GG: Not working yet. Please fix
            //if (tableRecordset.DataConnection.ConnectionString.Contains("SharePoint"))
            //    UploadTableChanges();
            //else
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

            var clientContext =GetSPAccessToken();

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

            LoadTablesList("All", "All", true, true);
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


        internal void AssignDataRecordsets()
        {
            applicationRecordset = getTableFromName("Application");
            databaseRecordset = getTableFromName("Database");
            tableRecordset = getTableFromName("Tables");
            columnRecordset = getTableFromName("Columns");
            relationRecordset = getTableFromName("Relations");
            if (tableRecordset is null | columnRecordset is null | relationRecordset is null)
            {
                MessageBox.Show("We need Tables, Columns, and Relations to upload to SharePoint.");
                return;
            }
        }

        private void InsertGenericSharePointRecord()
        {

        }

        internal int GetColumnNumber(System.Array record, string colName)
        {
            for (int i = 0; i < record.Length; i++)
                if (record.GetValue(i).ToString().ToLower() == colName.ToLower())
                    return i;
            return -1;
        }


        private string LookupFromColumnDataset(string SearchForColumnName, string ColumnToReturn)
        {
            var recs = columnRecordset.GetDataRowIDs("Column_Name = '" + SearchForColumnName + "'");
            if (recs.GetUpperBound(0) == -1) return "";
            var rec = columnRecordset.GetRowData((int)recs.GetValue(0));
            var ID = GetColumnNumber(columnRecordset.GetRowData(0), ColumnToReturn);
            return rec.GetValue(ID).ToString();
        }

        private string LookupFromTableDataset(string SearchForTable, string ColumnName)
        {
            var recs = tableRecordset.GetDataRowIDs("Table_Name = '" + SearchForTable + "'");
            if (recs.GetUpperBound(0) == -1) return "";
            var rec = tableRecordset.GetRowData((int) recs.GetValue(0));
            var ID = GetColumnNumber(tableRecordset.GetRowData(0), ColumnName);
            return rec.GetValue(ID).ToString();
        }


        internal bool isValidColumn(string columnList, string columnName)
        {
            // Some entities (such as SystemUser entity) are not useful for ER diagram. Skip those noisy not useful entities for visualization.
            return (columnList.Contains($",{columnName.ToLower()},"));
        }


        internal string GetSharePointListColumns(string listName)
        {
            if (gclientContext is null) gclientContext = GetSPAccessToken(true);
            ClientContext clientContext =  gclientContext;
            // clientcontext.Web.Lists.GetById - This option also can be used to get the list using List GUID
            // This value is NOT List internal name
            List targetList = clientContext.Web.Lists.GetByTitle(listName);
            FieldCollection oFieldCollection = targetList.Fields;

            // Here we have loaded Field collection including title of each field in the collection
            clientContext.Load(oFieldCollection, oFields => oFields.Include(field => field.Title));
            clientContext.ExecuteQuery();
            
            string columnList = ",";
            foreach (Field oField in oFieldCollection)
                columnList = columnList + oField.Title.ToLower() + ",";

            return columnList;
        }




        internal void UploadCompleteDataRecordsetToSharepoint()
        {
            string dbName = "";
            string appName = "Input Application Name";
            int i = 0;

            var clientContext = GetSPAccessToken();

            applicationColumns = GetSharePointListColumns("Application");
            databaseColumns = GetSharePointListColumns("Database");
            tableColumns = GetSharePointListColumns("Tables");
            columnColumns = GetSharePointListColumns("Columns");
            relationColumns = GetSharePointListColumns("Relations");

            tableRecordset = null;
            AssignDataRecordsets();

            var AppCol = GetColumnNumber(tableRecordset.GetRowData(0), "Application");
            if (tableRecordset.GetRowData(1).GetValue(AppCol).ToString() == "")
            {
                string ignore = "-1";
                var result = Utilites.ScreenEvents.ShowInputDialog(ref appName, ref ignore, "Site", "Ignore", "Enter Application Name");
                if (result != DialogResult.OK) return;
            }



            var records = tableRecordset.GetDataRowIDs("");
            var targetList = clientContext.Web.Lists.GetByTitle("Tables");
            var titleRow = tableRecordset.GetRowData(0);

            for (int recNo = 1; recNo < records.GetUpperBound(0); recNo++)
            {
                var record = tableRecordset.GetRowData(recNo);

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oItem = targetList.AddItem(itemCreateInfo);

                oItem["Title"] = "";
                for (int k = 0; k < record.Length; k++)
                {
                    string colName = titleRow.GetValue(k).ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (colName == "UniqueID") colName = "Title";
                    if (isValidColumn(tableColumns, colName))
                        oItem[colName] = record.GetValue(k);
                }

                if (oItem["Database"].ToString() == "") oItem["Database"] = "Web";
                if (oItem["Table_Type"].ToString() == "") oItem["Table_Type"] = "API";
                if (oItem["Application"].ToString() == "") oItem["Application"] = appName;

                // Create a hash if the item doesn't already have a unique identifier.
                if (oItem["Title"].ToString() == "") oItem["Title"] = GenerateHash(oItem["Application"] + "." + oItem["Database"] + "." + oItem["Table_Name"]);

                oItem.Update();

                if (++i >= 20)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + recNo + " of " + records.GetUpperBound(0));
                    clientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                clientContext.ExecuteQuery();


            records = columnRecordset.GetDataRowIDs("");

            targetList = clientContext.Web.Lists.GetByTitle("Columns");
            titleRow = columnRecordset.GetRowData(0);
            int tableNameID = GetColumnNumber(titleRow, "table_name"); //GG: Changed to Table_Name
            int columnNameID = GetColumnNumber(titleRow, "column_Name"); //GG: Changed to Table_Name
            int tabletypeID = GetColumnNumber(titleRow, "Table_Type");
            dbName = LookupFromTableDataset("account", "database");

            for (int recNo =1; recNo < records.GetUpperBound(0); recNo++)
            {
                var record = columnRecordset.GetRowData(recNo);

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oItem = targetList.AddItem(itemCreateInfo);

                oItem["Title"] = "";
                oItem["TableUniqueID"] = "";
                for (int k = 0; k < record.Length; k++)
                {
                    string colName = titleRow.GetValue(k).ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (colName == "UniqueID") colName = "Title";
                    if (isValidColumn(columnColumns, colName))
                        oItem[colName] = record.GetValue(k);
                }
                var tableName = record.GetValue(tableNameID).ToString();
//                oItem["Is_Primary"] = (record.GetValue(columnNameID).ToString() == LookupFromTableDataset(record.GetValue(tableNameID).ToString(), "HideprimaryIdAttribute")) ? true : false;
                if (oItem["Title"].ToString() == "")
                    oItem["Title"] = GenerateHash(appName + "." + dbName + "." + LookupFromTableDataset(tableName, "Table_Type") + "." + tableName + "." + record.GetValue(columnNameID).ToString());
                if (oItem["TableUniqueID"].ToString() == "") 
                    oItem["TableUniqueID"] = GenerateHash(appName + "." + dbName + "." + record.GetValue(tableNameID).ToString());// LookupFromTableDataset(record.GetValue(tableNameID).ToString(), "UniqueID");

                oItem.Update();

                if (++i >= 50)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + recNo + " of " + records.GetUpperBound(0));
                    clientContext.ExecuteQuery();
                    i = 0;
                }
            }
            if (i > 0)
                clientContext.ExecuteQuery();



            records = relationRecordset.GetDataRowIDs("");
            
            targetList = clientContext.Web.Lists.GetByTitle("Relations");
            titleRow = relationRecordset.GetRowData(0);
            int DatabaseCol = GetColumnNumber(titleRow, "Database");
            int EntityOneCol = GetColumnNumber(titleRow, "TableOne");
            int EntityManyCol = GetColumnNumber(titleRow, "TableMany"); 
            int columnOneCol = GetColumnNumber(titleRow, "ColumnMany");
            int columnManyCol = GetColumnNumber(titleRow, "ColumnMany");
            int UniqueCol = GetColumnNumber(titleRow, "UniqueID");
            int ConnectionTypeCol = GetColumnNumber(titleRow, "Connection_Type");
            for (int recNo = 1; recNo < records.GetUpperBound(0); recNo++)
            {
                var record = relationRecordset.GetRowData(recNo);

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oItem = targetList.AddItem(itemCreateInfo);

                oItem["Title"] = "";
                for (int k = 0; k < record.Length; k++)
                {
                    string colName = titleRow.GetValue(k).ToString().Replace("_VisDM_", "").Replace(" ", "_x0020_");
                    if (colName == "UniqueID") colName = "Title";
                    if (isValidColumn(relationColumns, colName))
                        oItem[colName] = record.GetValue(k);
                }

                // If this hasn't been uploaded yet then we generate IDs ourself
                if (UniqueCol == -1)
                {
                    dbName = DatabaseCol == -1 ? "Web" : dbName = record.GetValue(DatabaseCol).ToString();

                    if (EntityOneCol >= 0) 
                        if(record.GetValue(EntityOneCol).ToString() == "") 
                            oItem["TableOneID"] = GenerateHash(appName + "." + dbName + "." + record.GetValue(EntityOneCol).ToString());// LookupFromTableDataset(record.GetValue(EntityOneCol).ToString(), "UniqueID");
                    if (EntityManyCol >= 0)
                        if (record.GetValue(EntityManyCol).ToString() == "")
                        oItem["TableManyID"] = GenerateHash(appName + "." + dbName + "." + record.GetValue(EntityManyCol).ToString());// LookupFromTableDataset(record.GetValue(EntityOneCol).ToString(), "UniqueID");
                    if (columnOneCol >= 0)
                        if (record.GetValue(columnOneCol).ToString() == "")
                        oItem["ColumnOneID"] = GenerateHash(appName + "." + dbName + "." + record.GetValue(columnOneCol).ToString());// LookupFromTableDataset(record.GetValue(EntityOneCol).ToString(), "UniqueID");
                    if (columnManyCol >= 0)
                        if (record.GetValue(columnManyCol).ToString() == "")
                        oItem["ColumnManyID"] = GenerateHash(appName + "." + dbName + "." + record.GetValue(columnManyCol).ToString());// LookupFromTableDataset(record.GetValue(EntityOneCol).ToString(), "UniqueID");

                    if(ConnectionTypeCol ==-1) 
                            oItem["Connection_Type"] = "Foreign Key";
                    if(oItem["Title"].ToString() == "")
                        oItem["Title"] = GenerateHash(appName + "." + dbName + "." + record.GetValue(EntityOneCol).ToString() + "." + record.GetValue(columnOneCol).ToString() + "." + record.GetValue(EntityManyCol).ToString() + "." + record.GetValue(columnManyCol).ToString());
                }
                oItem.Update();

                if (++i >= 50)
                {
                    Utilites.ScreenEvents.DisplayVisioStatus("Uploading " + recNo + " of " + records.GetUpperBound(0));
                    clientContext.ExecuteQuery(); 
                    i = 0;
                }
            }
            if (i > 0)
                clientContext.ExecuteQuery();

        }


    }
}

