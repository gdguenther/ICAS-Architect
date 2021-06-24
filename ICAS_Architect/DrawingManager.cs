using System;
using System.Security;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using System.Data.SqlClient;
using Microsoft.Office.Tools.Ribbon;
using Dapper;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;
using System.Web.Script.Serialization;

namespace ICAS_Architect
{
    class DrawingManager
    {
        const double SHDW_PATTERN = 0;
        const double BEGIN_ARROW_MANY = 29;
        const double BEGIN_ARROW = 24;  // One end
        const double END_ARROW = 29;    // Many end
        const double LINE_COLOR_MANY = 10;
        const double LINE_COLOR = 8;
        const double LINE_PATTERN_MANY = 2;
        const double LINE_PATTERN = 1;
        const string LINE_WEIGHT = "2pt";
        const double ROUNDING = 0.0625;
        const double HEIGHT = 0.25;
        const short NAME_CHARACTER_SIZE = 12;
        const short FONT_STYLE = 225;
        const short VISIO_SECTION_OJBECT_INDEX = 1;
        private const string SKIP_ENTITIES_FOR_CONNECTION = @",businessunit,organization,systemuser,team,transactioncurrency,workflow,webresource,webwizard,uom,transformationparametermapping,transformationmapping,traceregarding,tracelog,topicmodelconfiguration,topicmodel,timezonedefinition,template,teamtemplate,syncattributemappingprofile,syncattributemapping,subscription,solution,dependencynode,solution,solutioncomponent,sdkmessage,sdkmessagefilter,sdkmessagepair,sdkmessageprocessingstep,sdkmessageprocessingstepsecureconfig,sdkmessagerequest,sdkmessageresponse,role,roletemplate,rollupproperties,routingrule,ribboncustomization,ribbondiff,report,relationshiprole,recurringappointmentmaster,recurrencerule,recommendationmodel,recommendationmodelmapping,recommendationmodelversion,ratingmodel,ratingvalue,principalsyncattributemap,privilege,pluginassembly,plugintype,Empty,asyncoperation,processtrigger,systemform,attributemap,entitymap,knowledgesearchmodel,azureserviceconnection,advancedsimilarityrule,teamprofiles,fieldsecurityprofile,systemuserprofiles,fieldpermission,mobileofflineprofileitemassociation,mobileofflineprofileitem,mobileofflineprofile,savedquery,publisher,publisheraddress,appmodule,appmoduleroles,appmodulecomponent,importlog,importdata,importfile,import,ownermapping,importentitymapping,importmap,columnmapping,lookupmapping,picklistmapping,";

        Visio.Application vApplication = null;
        private frmGraphViewer graphViewer = null;              // ER viewer form
        public ERInformation ERI = null;
        ICASArchitect ribbon = null;

        private Visio.DataRecordset tableRecordset = null;
        private Visio.DataRecordset columnRecordset = null;
        private Visio.DataRecordset relationRecordset = null;


        internal DrawingManager()
        {
            vApplication = Globals.ThisAddIn.Application;
            ribbon = Globals.Ribbons.Ribbon;
        }

        private Visio.Shape GetShapeOnPage(string ShapeName)
        {
            foreach (Visio.Shape shp in Globals.ThisAddIn.Application.ActivePage.Shapes)
                if (shp.NameU == ShapeName)
                    return shp;
            return null;
        }

        private Visio.Shape InsertMyShape(string shapeName, int shapeNo)
        {
            Visio.Document docStencil = vApplication.Documents.OpenEx("ICAS Data Architect.vssx", (short)6); //   Visio.Document docStencle = vApplication.Documents.OpenEx("Basic_U.VSS", (short)6);//var vsoMaster = docStencle.Masters["Rectangle"];
            Visio.Page Pg = vApplication.ActivePage;
            Visio.Shape shape = null;
            var dsRowList = tableRecordset.GetDataRowIDs("[Table_Name] = '" + shapeName + "'");
            int dsRow = (int)dsRowList.GetValue(0);
            var rowData = tableRecordset.GetRowData(dsRow);
            var targetName = rowData.GetValue(0).ToString();
            int i = 1;


            if (!ribbon.chkShowAttributes.Checked)
            {
                var vsoMaster = docStencil.Masters["Entity"];
                shape = Pg.DropLinked(vsoMaster, 4, 4, tableRecordset.ID, dsRow, false);
                shape.NameU = targetName;
                shape.Text = targetName;
            }
            else
            {
                // use attributes information
                var vsoMaster = docStencil.Masters["Entity And Attributes"];
                var vsoAttribute = docStencil.Masters["Attribute"];
                var vsoPKAttribute = docStencil.Masters["Primary Key Attribute"];
                var vsoPKSeparator = docStencil.Masters["Primary Key Separator"];

                shape = Pg.DropLinked(vsoMaster, 2.0 + shapeNo / 20.0, 2.0 + shapeNo / 20.0, tableRecordset.ID, dsRow, false);
                shape.NameU = targetName;

                Visio.Shape attShape = null;

                var rowIDs = columnRecordset.GetDataRowIDs("[Table_Name]='" + targetName + "' and Is_Primary=true");
                int intRowIDs =0;
                int j = 0;
                intRowIDs = rowIDs.GetUpperBound(0);
//                string primaryKeyName = LookupFromTableDataset(targetName, "IsPrimaryID");
/*                if(primaryKeyName != "")
                {
                    int rowID = (int)rowIDs.GetValue(i);
                    var row = columnRecordset.GetRowData(i);
                        attShape = vApplication.ActivePage.DropLinked(vsoPKAttribute, 5, 5, columnRecordset.ID, rowID, false);
                    attShape.Text = row.GetValue(1).ToString();
                    shape.ContainerProperties.InsertListMember(attShape, j++);
                    Utilites.ScreenEvents.DoEvents();
                }
                for (i=0; i < intRowIDs; i++)
                {
                    int rowID = (int) rowIDs.GetValue(i);
                    var row = columnRecordset.GetRowData(i);
                    attShape = vApplication.ActivePage.DropLinked(vsoPKAttribute, 5,5,columnRecordset.ID, rowID, false);
                    attShape.Text = row.GetValue(1).ToString();
                    shape.ContainerProperties.InsertListMember(attShape, j++);
                    Utilites.ScreenEvents.DoEvents();
                } */

                attShape = vApplication.ActivePage.Drop(vsoPKSeparator, 5, 5);
                shape.ContainerProperties.InsertListMember(attShape, j++);

                rowIDs = columnRecordset.GetDataRowIDs("[Table_Name]='" + targetName + "' and Is_Primary<>true");
                intRowIDs = rowIDs.GetUpperBound(0);
                for (i = 0; i < intRowIDs; i++)
                {
                    int rowID = (int)rowIDs.GetValue(i);
                    var row = columnRecordset.GetRowData(rowID);
                    var foreignKeyName = row.GetValue(1).ToString();
                    if (relationRecordset.GetDataRowIDs("[TableMany] = '" + foreignKeyName + "'").GetUpperBound(0) > 0 )
                    {
                        attShape = vApplication.ActivePage.DropLinked(vsoAttribute, 5, 5, columnRecordset.ID, rowID, false);
                        attShape.Text = row.GetValue(1).ToString();
                        shape.ContainerProperties.InsertListMember(attShape, j++);
                        Utilites.ScreenEvents.DoEvents();
                    }
                }

            }
            return (shape);
        }

        internal Visio.DataRecordset getTableFromName(string tblName)
        {
            try{
                for (int i = 1; i <= vApplication.ActiveDocument.DataRecordsets.Count; i++)
                    if (vApplication.ActiveDocument.DataRecordsets[i].Name == tblName)
                        return vApplication.ActiveDocument.DataRecordsets[i];
            }catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return null;
        }


        internal void AddShapeFromTableData(Visio.Shape shape)
        {
            int i = 1;
            int DSRow = 0;
            Visio.Page Pg = vApplication.ActivePage;
            
            // If the shape is not linked to a tableRecordset exit function, otherwise replace the shape
            try
            {
                if (tableRecordset is null | columnRecordset is null | relationRecordset is null)
                {
                    tableRecordset = getTableFromName("Tables");
                    columnRecordset = getTableFromName("Columns");
                    relationRecordset = getTableFromName("Relations");
                    if (tableRecordset is null | columnRecordset is null | relationRecordset is null)
                        return;
                }

                DSRow = shape.GetLinkedDataRow(tableRecordset.ID);
                var shapeName = tableRecordset.GetRowData(DSRow).GetValue(0).ToString();
                shape.Delete();
                shape = InsertMyShape(shapeName, 0);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return;
            }

            // Find all the related entities, even if we aren't going to insert them.
/*            List<string> relatedEntities = null;
            try
            {
//                relatedEntities = ERI.FindRelatedEntities(shape.get_Cells("Prop._VisDM_Table_Name").ResultStr[0]);
                relatedEntities = ERI.FindRelatedEntities(shape.NameU);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return;
            } */

            string qstring = "([TableOne] = '" + shape.NameU + "' OR [TableMany] = '" + shape.NameU + "')";
            var relRecords = relationRecordset.GetDataRowIDs(qstring);
            var recs = relRecords.GetUpperBound(0);

            if (ribbon.chkDrawRelatedEntities.Checked)
            {
                for (int j = 0; j < recs; j++)
                {
                    int rec = (int) relRecords.GetValue(j);
                    var record = relationRecordset.GetRowData(rec);
                    // choose EntityMany if EntityOne is the same as the shape we're adding
                    string EntityToAdd = (shape.NameU == record.GetValue(1).ToString()) ? record.GetValue(2).ToString() : record.GetValue(1).ToString();

                    if (IsEntityInSkipList(EntityToAdd))   continue;

                    // Only add if the shape doesn't exist yet.
                    if (GetShapeOnPage(EntityToAdd) is null)
                        InsertMyShape(EntityToAdd, ++i);
//                    InsertMyShape((int)relRecords.GetValue(j), ++i);

                    Utilites.ScreenEvents.DoEvents();
                }
            }


            // now add the connectors
            for (int j = 0; j < recs; j++)
            {
                int rec = (int)relRecords.GetValue(j);
                var record = relationRecordset.GetRowData(rec);
                Visio.Shape ShapeOne = GetShapeOnPage(record.GetValue(1).ToString());
                Visio.Shape ShapeMany = GetShapeOnPage(record.GetValue(2).ToString());

                // skip if shapes don't exist or if connector already exists
                if (ShapeOne is null | ShapeMany is null) continue;
                var conShape = (GetShapeOnPage(ShapeOne.NameU + "-" + ShapeMany.NameU));
                if (conShape == null)
                    conShape = this.DrawDirectionalDynamicConnector(ShapeOne, ShapeMany, false);

                var commentCell = conShape.get_CellsSRC((short) Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowMisc, (short)Visio.VisCellIndices.visComment);
                commentCell.FormulaU = "\"" + commentCell.FormulaU.Replace("\"", "") + record.GetValue(4).ToString() + "\n\"";
                conShape.get_CellsU("Prop.LinkedField").Formula = commentCell.FormulaU.Replace("\n",", \n");

                // above 500 connections becomes sketchy for Viso
                Utilites.ScreenEvents.DisplayVisioStatus("Added " + i++ + " connectors ");
                if (i >= 500) return;
            }


            if (ribbon.chkDrawRelatedEntities.Checked)
            {   // Only refresh the page layout if we are importing lots of information

                vApplication.ActiveWindow.DeselectAll();
                foreach (Visio.Shape shp in Pg.Shapes)
                {
                    System.Array arr = null;
                    shp.GetLinkedDataRecordsetIDs(out arr);
                    if (arr.Length > 0 ) 
                    {
                        vApplication.ActiveWindow.Select(shp, 3);
                    }
                }
                vApplication.ActiveWindow.Selection.Layout();
                vApplication.ActiveWindow.DeselectAll();
/*                foreach(int containerID in vApplication.ActivePage.GetContainers(Visio.VisContainerNested.visContainerExcludeNested))
                {
                    Visio.Shape vsoContainerShape = vApplication.ActivePage.Shapes.get_ItemFromID(containerID);
                    vApplication.ActiveWindow.Select(vsoContainerShape, 2);
                }
                vApplication.ActiveWindow.Selection.LayoutIncremental((Microsoft.Office.Interop.Visio.VisLayoutIncrementalType)2, Visio.VisLayoutHorzAlignType.visLayoutHorzAlignDefault, Visio.VisLayoutVertAlignType.visLayoutVertAlignDefault                  , .5, .5, Visio.VisUnitCodes.visInches);*/
                Pg.ResizeToFitContents();
                ribbon.chkDrawRelatedEntities.Checked = false;
            }
        }
        internal bool IsEntityInSkipList(string entityName)
        {
            // Some entities (such as SystemUser entity) are not useful for ER diagram. Skip those noisy not useful entities for visualization.
            return (SKIP_ENTITIES_FOR_CONNECTION.Contains($",{entityName},"));
        }

        private Visio.Shape DrawDirectionalDynamicConnector(Visio.Shape shapeFrom, Visio.Shape shapeTo, bool isManyToMany)
        {
            Visio.Application app = Globals.ThisAddIn.Application;

            // Add a dynamic connector to the page.
            Visio.Shape connectorShape = shapeFrom.ContainingPage.Drop(app.ConnectorToolDataObject, 0.0, 0.0);
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowShapeLayout, (short)Visio.VisCellIndices.visSLOLineRouteExt).ResultIU = 2;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowShapeLayout, (short)Visio.VisCellIndices.visSLORouteStyle).ResultIU = 1;
            //.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = "1"
            //.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = "16"
            // Set the connector properties, using different arrows, colors, and patterns for many-to-many relationships.
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visFillShdwPattern).ResultIU = SHDW_PATTERN;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineBeginArrow).ResultIU = isManyToMany ? BEGIN_ARROW_MANY : BEGIN_ARROW;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineEndArrow).ResultIU = END_ARROW;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineColor).ResultIU = isManyToMany ? LINE_COLOR_MANY : 21;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLinePattern).ResultIU = isManyToMany ? LINE_PATTERN : LINE_PATTERN;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visLineRounding).ResultIU = ROUNDING;

            // Connect the starting point.
            Visio.Cell cellBeginX = connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DBeginX);
            cellBeginX.GlueTo(shapeFrom.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));

            // Connect the ending point.
            Visio.Cell cellEndX = connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DEndX);
            cellEndX.GlueTo(shapeTo.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));

            var intPropRow2 = connectorShape.AddNamedRow( (short)Visio.VisSectionIndices.visSectionProp, "LinkedField", (short)Visio.VisRowTags.visTagDefault);
//            var intPropRow2 = connectorShape.AddRow((short)Visio.VisSectionIndices.visSectionProp, (short)Visio.VisRowIndices.visRowLast, (short)Visio.VisRowTags.visTagDefault);
  //          connectorShape.get_Section((short)Visio.VisSectionIndices.visSectionProp).get_Row(intPropRow2).NameU = "LinkedField";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsLabel).FormulaU = "\"Linked Field\"";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsType).FormulaU = "0";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsFormat).FormulaU = "";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsLangID).FormulaU = "2057";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsCalendar).FormulaU = "";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsPrompt).FormulaU = "";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsValue).FormulaU = "\"MyField\"";
            connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionProp, intPropRow2, (short)Visio.VisCellIndices.visCustPropsSortKey).FormulaU = "";

            connectorShape.NameU = shapeFrom.NameU + "-" + shapeTo.NameU;
            return (connectorShape);
        }


        internal void ImportJSONFile()
        {
            if (graphViewer != null) graphViewer.Dispose();

            OpenFileDialog JsonDialog = new OpenFileDialog();
            JsonDialog.Filter = "Json files (*.json)|*.json";
           
            DialogResult result = JsonDialog.ShowDialog();
            if (result != DialogResult.OK) return;

            string jsonFile = JsonDialog.FileName;
            //            labelStatus.Text = $"Loading {Path.GetFileName(jsonFile)}.";

            //GG: No idea why he added the JSON reader into the Graph viewer originally. Pull it out.
            graphViewer = new frmGraphViewer();
            graphViewer.LoadMetadataJsonFile(jsonFile);
            ERI = graphViewer.entityRelations;
            
            graphViewer.Close();
        }

        // Loads data from our Entity Metadata class - usually imported from JSON or Dynamics rather than linked directly from SQL Server or Sharepoint.
        internal void LoadIntoDataTable()
        {
            if (ERI is null) return;
            // Weirdly, DataRecordsets are enumerated from 1
            for (int i = 1; i <= vApplication.ActiveDocument.DataRecordsets.Count; i++)
                if (vApplication.ActiveDocument.DataRecordsets[i].Name == "Tables" | vApplication.ActiveDocument.DataRecordsets[i].Name == "Columns" | vApplication.ActiveDocument.DataRecordsets[i].Name == "Relations")
                    vApplication.ActiveDocument.DataRecordsets[i--].Delete();
            Utilites.ScreenEvents.DoEvents();

            string headString = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'>" +
             " <s:Schema id='RowsetSchema'> <s:ElementType name='row' content='eltOnly' rs:updatable='true'>" +
             " <s:AttributeType name='c1' rs:name='Table_Name' rs:number='1' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c2' rs:name='Database' rs:number='2' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c3' rs:name='Table_Type' rs:number='3' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c4' rs:name='Description' rs:number='4' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c5' rs:name='Application' rs:number='5' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c6' rs:name='Display_Name' rs:number='6' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c7' rs:name='Schema' rs:number='7' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c8' rs:name='Set_Name' rs:number='8' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c9' rs:name='EntityListURL' rs:number='9' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c10' rs:name='EntityDefinitionsPath' rs:number='10' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c11' rs:name='ExternalID' rs:number='11' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c12' rs:name='HideprimaryIdAttribute' rs:number='11' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:extends type='rs:rowbase'/>" +
             " </s:ElementType> </s:Schema>";

            var dts = headString + new System.Xml.Linq.XElement("rs_____data", ERI.EREntities.Select(x => new System.Xml.Linq.XElement("z_____row",
                                                                               new System.Xml.Linq.XAttribute("c1", x.entityLogicalName ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c2", x.databaseName ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c3", x.Table_Type ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c4", x.descriptiontext ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c5", ""),
                                                                               new System.Xml.Linq.XAttribute("c6", x.displayNamesText ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c7", x.entitySchema ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c8", x.entitySetName ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c9", x.EntityListUrl ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c10", x.entityDefinitionsPath ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c11", x.metadataId ?? ""),
                                                                               new System.Xml.Linq.XAttribute("c12", x.primaryIdAttribute ?? "")
                                               ))).ToString();

            // Apparently I don't know how to use Linq namespaces properly, so I'm hacking this for the timebeing.
            dts = dts.Replace("_____", ":");
            dts = dts + " </xml>";

            tableRecordset = vApplication.ActiveDocument.DataRecordsets.AddFromXML(dts,  0, "Tables");
            Utilites.ScreenEvents.DoEvents();

            // make sure the External Data Window is open
            vApplication.ActiveWindow.Windows.ItemFromID[2044].Visible = true;
            Utilites.ScreenEvents.DoEvents();


            headString = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'>" +
             " <s:Schema id='RowsetSchema'> <s:ElementType name='row' content='eltOnly' rs:updatable='true'>" +
             " <s:AttributeType name='c1' rs:name='Column_Name' rs:number='1' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c2' rs:name='Description' rs:number='2' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c3' rs:name='Ordinal_Position' rs:number='3' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c4' rs:name='Table_Name' rs:number='4' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c5' rs:name='Default' rs:number='5' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c6' rs:name='Is_Nullable' rs:number='6' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='bool' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c7' rs:name='Is_Primary' rs:number='7' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='bool' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c8' rs:name='Data_Type' rs:number='8' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c9' rs:name='Character_Maximum_Length' rs:number='9' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='integer' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c10' rs:name='Character_Octet_Length' rs:number='10' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='integer' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c11' rs:name='Numeric_Precision' rs:number='11' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='integer' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c12' rs:name='Numeric_Precision_Radix' rs:number='12' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='integer' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c13' rs:name='Numeric_Scale' rs:number='13' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='integer' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c14' rs:name='Date_Time_Precision' rs:number='14' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='integer' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c15' rs:name='Character_Set_Catalog' rs:number='15' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c16' rs:name='Character_Set_Schema' rs:number='16' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c17' rs:name='Collation_Catalog' rs:number='17' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c18' rs:name='Collation_Schema' rs:number='18' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c19' rs:name='Collation_Name' rs:number='19' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c20' rs:name='ExternalID' rs:number='20' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c21' rs:name='TableUniqueID' rs:number='21' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:extends type='rs:rowbase'/>" +
             " </s:ElementType> </s:Schema>";
            dts = headString + new System.Xml.Linq.XElement("rs_____data", ERI.EREntitieAttributes.Select(x => new System.Xml.Linq.XElement("z_____row",
                                                                                new System.Xml.Linq.XAttribute("c1", x.AttributeName ?? ""),
                                                                                new System.Xml.Linq.XAttribute("c2", x.Description ?? ""),
                                                                                new System.Xml.Linq.XAttribute("c3", x.ColumnNumber ),
                                                                                new System.Xml.Linq.XAttribute("c4", x.EntityName ?? ""),     //GG: change this to be the EntityID
                                                                                new System.Xml.Linq.XAttribute("c5", x.Default ?? "" ),                      //GG: get column default
                                                                                new System.Xml.Linq.XAttribute("c6", x.IsNullable),
                                                                                new System.Xml.Linq.XAttribute("c7", x.IsPrimaryID ),
                                                                                new System.Xml.Linq.XAttribute("c8", x.AttributeType ?? ""),
                                                                                x.CHARACTER_MAXIMUM_LENGTH == null ? null : new System.Xml.Linq.XAttribute("c9", x.CHARACTER_MAXIMUM_LENGTH),
                                                                                x.CHARACTER_OCTET_LENGTH == null ? null :   new System.Xml.Linq.XAttribute("c10", x.CHARACTER_OCTET_LENGTH ?? 0),
                                                                                x.NUMERIC_PRECISION == null ? null :        new System.Xml.Linq.XAttribute("c11", x.NUMERIC_PRECISION ?? 0),
                                                                                x.NUMERIC_PRECISION_RADIX == null ? null :  new System.Xml.Linq.XAttribute("c12", x.NUMERIC_PRECISION_RADIX ?? 0),
                                                                                x.NUMERIC_SCALE == null ? null :            new System.Xml.Linq.XAttribute("c13", x.NUMERIC_SCALE ?? 0),
                                                                                x.DATETIME_PRECISION == null ? null :       new System.Xml.Linq.XAttribute("c14", x.DATETIME_PRECISION ?? 0),
                                                                                x.Character_Set_Catalog == null ? null :    new System.Xml.Linq.XAttribute("c15", x.Character_Set_Catalog),
                                                                                x.Character_Set_Schema == null ? null :     new System.Xml.Linq.XAttribute("c15", x.Character_Set_Schema),
                                                                                x.Collation_Catalog == null ? null :        new System.Xml.Linq.XAttribute("c15", x.Collation_Catalog),
                                                                                x.Collation_Schema == null ? null :         new System.Xml.Linq.XAttribute("c15", x.Collation_Schema),
                                                                                x.Collation_Name == null ? null :           new System.Xml.Linq.XAttribute("c15", x.Collation_Name),
                                                                                new System.Xml.Linq.XAttribute("c20", x.MetadataId),
                                                                                new System.Xml.Linq.XAttribute("c21", "")
                                            ))).ToString();


            // Apparently I don't know how to use Linq namespaces properly, so I'm hacking this for the timebeing.
            dts = dts.Replace("_____", ":");
            dts = dts + " </xml>";

            columnRecordset = vApplication.ActiveDocument.DataRecordsets.AddFromXML(dts, 0, "Columns");
            Utilites.ScreenEvents.DoEvents();


            headString = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'>" +
             " <s:Schema id='RowsetSchema'> <s:ElementType name='row' content='eltOnly' rs:updatable='true'>" +
             " <s:AttributeType name='c1' rs:name='Relation_Name' rs:number='1' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c2' rs:name='TableOne' rs:number='2' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c3' rs:name='TableMany' rs:number='3' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c4' rs:name='ColumnOne' rs:number='4' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c5' rs:name='ColumnMany' rs:number='5' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c6' rs:name='Intersect_Entity' rs:number='6' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c7' rs:name='Connection_Type' rs:number='7' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:extends type='rs:rowbase'/>" +
             " </s:ElementType> </s:Schema>";
            dts = headString + new System.Xml.Linq.XElement("rs_____data", ERI.ERRelations.Select(x => new System.Xml.Linq.XElement("z_____row",
                                                                                new System.Xml.Linq.XAttribute("c1", x.RelationName ?? ""),
                                                                                new System.Xml.Linq.XAttribute("c2", x.EntityOne ?? ""),
                                                                                new System.Xml.Linq.XAttribute("c3", x.EntityMany ?? ""),
                                                                                new System.Xml.Linq.XAttribute("c4", x.AttributeOne ?? ""),
                                                                                new System.Xml.Linq.XAttribute("c5", x.AttributeMany ?? ""),
                                                                                new System.Xml.Linq.XAttribute("c6", ""),
                                                                                new System.Xml.Linq.XAttribute("c7", x.IntersectEntity ?? "")
                                             ))).ToString();


            // Apparently I don't know how to use Linq namespaces properly, so I'm hacking this for the timebeing.
            dts = dts.Replace("_____", ":");
            dts = dts + " </xml>";

            relationRecordset = vApplication.ActiveDocument.DataRecordsets.AddFromXML(dts, 0, "Relations");
            Utilites.ScreenEvents.DoEvents();

            vApplication.ActiveWindow.Windows.ItemFromID[2044].SelectedDataRecordset = tableRecordset;
            Utilites.ScreenEvents.DoEvents();

        }

        private SqlConnection GetConnection(string txtServer, string txtDatabase)
        {
            //            string connectionString = "Provider=SQLOLEDB.1;Data Source=" + txtServer + ";Initial Catalog=" + txtDatabase;

            HttpDownloadClient httpDownloadClient = new HttpDownloadClient("https://database.windows.net");
            httpDownloadClient.Connect("https://database.windows.net");


            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://login.microsoftonline.com/common/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https://database.windows.net/");
            request.Headers["Metadata"] = "true";
            request.Method = "GET";
            string accessToken = null;

            try
            {
                // Call managed identities for Azure resources endpoint.
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                // Pipe response Stream to a StreamReader and extract access token.
                StreamReader streamResponse = new StreamReader(response.GetResponseStream());
                string stringResponse = streamResponse.ReadToEnd();
                JavaScriptSerializer j = new JavaScriptSerializer();
                Dictionary<string, string> list = (Dictionary<string, string>)j.Deserialize(stringResponse, typeof(Dictionary<string, string>));
                accessToken = list["access_token"];
            }
            catch (Exception e)
            {
                string errorText = String.Format("{0} \n\n{1}", e.Message, e.InnerException != null ? e.InnerException.Message : "Acquire token failed");
            }

            //
            // Open a connection to the server using the access token.
            //
            if (accessToken != null)
            {
                string connectionString = "Provider=SQLOLEDB.1;Data Source=" + txtServer + ";Initial Catalog=" + txtDatabase;
                SqlConnection conn = new SqlConnection(connectionString);
                conn.AccessToken = accessToken;
                conn.Open();
                return conn;
            }

            return null;
            //            ds.setMSIClientId("94de34e9-8e8c-470a-96df-08110924b814"); // Replace with Client ID of User-Assigned Managed Identity to be used

        }

        HttpDownloadClient httpDownloadClient = null;
        SqlConnection sqlConnection = null;
        internal void ConnectToSQLServerMSI(string txtServer, string txtDatabase)
        {
            if (httpDownloadClient == null)
            {
                httpDownloadClient = new HttpDownloadClient("https://database.windows.net");
                httpDownloadClient.Connect(txtServer);

                string connectionString = "Data Source=" + txtServer + ";Initial Catalog=" + txtDatabase;
                sqlConnection = new SqlConnection(connectionString);
                sqlConnection.AccessToken = httpDownloadClient.accessToken;
            }

            // a SQL Connection open too long is invalid. 
            if (sqlConnection.State == ConnectionState.Open)
                sqlConnection.Close();
        }


        internal string CreateXMLHeaderLine(int column, string name, string datatype)
        {
            return " <s:AttributeType name='" + name + "' rs:name='" + name + "' rs:nullable='true' rs:maydefer='true' rs:write='true'> " +
                " <s:datatype dt:type='" + datatype + "' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>";
        }

        internal void CreateTableDataRecordsetFromSQL()
        {
            var queryString = "select Table_Name=TABLE_SCHEMA + '.' + TABLE_NAME, [Database]=TABLE_CATALOG, Table_Type, Description= '', Application= '', Display_Name=Table_Name from INFORMATION_SCHEMA.tables [z:row] for xml auto";

            SqlCommand sqlCommand = new SqlCommand(queryString, sqlConnection);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

            string xmlRecordset = "";
            while (sqlDataReader.Read())
                xmlRecordset = xmlRecordset + sqlDataReader[0].ToString();

            string dts = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'>" +
             " <s:Schema id='RowsetSchema'> <s:ElementType name='row' content='eltOnly' rs:updatable='true'>" +
             CreateXMLHeaderLine(1,"Table_Name", "string") +
             CreateXMLHeaderLine(2, "Database", "string") +
             CreateXMLHeaderLine(3, "Table_Type", "string") +
             CreateXMLHeaderLine(4, "Description", "string") +
             CreateXMLHeaderLine(5, "Application", "string") +
             CreateXMLHeaderLine(6, "Display_Name", "string") +
             " <s:extends type='rs:rowbase'/>" +
             " </s:ElementType> </s:Schema>  <rs:data> " +
              
             xmlRecordset + 
             
             "</rs:data>  </xml>";

            //add recordset if it doesn't exist or just refresh it.
            if(tableRecordset is null)
                tableRecordset = vApplication.ActiveDocument.DataRecordsets.AddFromXML(dts, 0, "Tables");
            else
                tableRecordset.RefreshUsingXML(dts);

            vApplication.ActiveWindow.Windows.ItemFromID[2044].Visible = true;
            Utilites.ScreenEvents.DoEvents();
            sqlDataReader.Close();
        }

        internal void CreateColumnDataRecordsetFromSQL()
        {
            var queryString = "select Table_Name = [z:row].Table_schema + '.' + [z:row].table_name, [z:row].Column_name, [z:row].Data_Type, Description = null, [z:row].ORDINAL_POSITION,  " +
                 " Is_Primary = iif(cu.column_name is null, 0, 1), Is_Nullable=iif(is_nullable='yes', 1, 0), [z:row].CHARACTER_MAXIMUM_LENGTH, [z:row].CHARACTER_OCTET_LENGTH, [z:row].NUMERIC_PRECISION, [z:row].NUMERIC_PRECISION_RADIX, [z:row].NUMERIC_SCALE, [z:row].DATETIME_PRECISION, DB_NAME() [Database]  " +
                 "    from INFORMATION_SCHEMA.columns [z:row] left join INFORMATION_SCHEMA.KEY_COLUMN_USAGE cu  " +
                 "         on [z:row].table_schema = cu.table_schema and [z:row].table_name = cu.table_name and [z:row].column_name = cu.column_name and OBJECTPROPERTY(OBJECT_ID(cu.CONSTRAINT_SCHEMA +'.' + QUOTENAME(cu.CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND cu.TABLE_NAME = 'TableName' AND cu.TABLE_SCHEMA = 'Schema' " +
                 " for xml auto";

            SqlCommand sqlCommand = new SqlCommand(queryString, sqlConnection);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

            string xmlRecordset = "";
            while (sqlDataReader.Read())
                xmlRecordset = xmlRecordset + sqlDataReader[0].ToString();

            string dts = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'>" +
             " <s:Schema id='RowsetSchema'> <s:ElementType name='row' content='eltOnly' rs:updatable='true'>" +
             CreateXMLHeaderLine(1, "Table_Name", "string") +
             CreateXMLHeaderLine(2, "Column_name", "string") +
             CreateXMLHeaderLine(3, "Data_Type", "string") +
             CreateXMLHeaderLine(4, "Description", "string") +
             CreateXMLHeaderLine(5, "Ordinal_Position", "int") +
             CreateXMLHeaderLine(6, "Is_Primary", "bool") +
             CreateXMLHeaderLine(6, "Is_Nullable", "bool") +
             CreateXMLHeaderLine(7, "Character_Maximum_Length", "string") +
             CreateXMLHeaderLine(8, "Character_Octet_Length", "string") +
             CreateXMLHeaderLine(9, "Numeric_Precision", "string") +
             CreateXMLHeaderLine(10, "Numeric_Precision_radix", "string") +
             CreateXMLHeaderLine(11, "Numeric_Scale", "string") +
             CreateXMLHeaderLine(12, "DateTime_Precision", "string") +
             CreateXMLHeaderLine(13, "Database", "string") +
             " <s:extends type='rs:rowbase'/>" +
             " </s:ElementType> </s:Schema>  <rs:data> " +

             xmlRecordset +

             "</rs:data>  </xml>";

            //add recordset if it doesn't exist or just refresh it.
            if (columnRecordset is null)
                columnRecordset = vApplication.ActiveDocument.DataRecordsets.AddFromXML(dts, 0, "Columns");
            else
                columnRecordset.RefreshUsingXML(dts);

            vApplication.ActiveWindow.Windows.ItemFromID[2044].Visible = true;
            Utilites.ScreenEvents.DoEvents();
            sqlDataReader.Close();
        }


        internal void CreateRelationsDataRecordsetFromSQL()
        {
            var queryString = "select schema_name(fk_tab.schema_id) + '.' + fk_tab.name as EntityMany, " +
                                    " schema_name(pk_tab.schema_id) + '.' + pk_tab.name as EntityOne,  " +
                                    " substring(column_names, 1, len(column_names) - 1) as [AttributeMany],  " +
                                    " [z:row].name as RelationName, DB_NAME() [Database], 'Foreign Key' as Connection_Type " +
                                " from sys.foreign_keys[z:row] " +
                                        " inner join sys.tables fk_tab on fk_tab.object_id = [z:row].parent_object_id " +
                                        " inner join sys.tables pk_tab  on pk_tab.object_id = [z:row].referenced_object_id " +
                                    " cross apply(select col.[name] + ', '  from sys.foreign_key_columns fk_c " +
                                                    " inner join sys.columns col on fk_c.parent_object_id = col.object_id and fk_c.parent_column_id = col.column_id " +
                                                " where fk_c.parent_object_id = fk_tab.object_id and fk_c.constraint_object_id = [z:row].object_id " +
                                                " order by col.column_id for xml path ('') ) D(column_names) " +
                                " order by 1,2 for xml auto"; 

            SqlCommand sqlCommand = new SqlCommand(queryString, sqlConnection);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

            string xmlRecordset = "";
            while (sqlDataReader.Read())
                xmlRecordset = xmlRecordset + sqlDataReader[0].ToString();

            string dts = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'>" +
             " <s:Schema id='RowsetSchema'> <s:ElementType name='row' content='eltOnly' rs:updatable='true'>" +
             CreateXMLHeaderLine(1, "EntityMany", "string") +
             CreateXMLHeaderLine(2, "EntityOne", "string") +
             CreateXMLHeaderLine(3, "AttributeMany", "string") +
             CreateXMLHeaderLine(4, "RelationName", "string") +
             CreateXMLHeaderLine(5, "Database", "string") +
             " <s:extends type='rs:rowbase'/>" +
             " </s:ElementType> </s:Schema>  <rs:data> " +

             xmlRecordset +

             "</rs:data>  </xml>";

            //add recordset if it doesn't exist or just refresh it.
            if (relationRecordset is null)
                relationRecordset = vApplication.ActiveDocument.DataRecordsets.AddFromXML(dts, 0, "Relations");
            else
                relationRecordset.RefreshUsingXML(dts);

            vApplication.ActiveWindow.Windows.ItemFromID[2044].Visible = true;
            Utilites.ScreenEvents.DoEvents();
            sqlDataReader.Close();
        }


        internal void  ImportFromSQLServer()
        {
            string txtServer = Globals.ThisAddIn.registryKey.GetValue("ImportServerName", "edi-dwuat-tmp").ToString();
            string txtDatabase = Globals.ThisAddIn.registryKey.GetValue("ImportDBName", "DataVault").ToString();
            string getServer = txtServer;
            string getDB = txtDatabase;
            if (Utilites.ScreenEvents.ShowInputDialog(ref getServer, ref getDB, "Server", "Database", "Choose Server and Database to import.") == System.Windows.Forms.DialogResult.Cancel) return;

            // if user has asked for a different connection then drop the current connection and save the new info to the registry
            if (txtServer != getServer | txtDatabase != getDB)
            {
                sqlConnection = null;
                Globals.ThisAddIn.registryKey.SetValue("ImportServerName", getServer);
                Globals.ThisAddIn.registryKey.SetValue("ImportDBName", getDB);
            }


            for (int i = 1; i <= vApplication.ActiveDocument.DataRecordsets.Count; i++)
                if (vApplication.ActiveDocument.DataRecordsets[i].Name == "Tables" | vApplication.ActiveDocument.DataRecordsets[i].Name == "Columns" | vApplication.ActiveDocument.DataRecordsets[i].Name == "Relations")
                    vApplication.ActiveDocument.DataRecordsets[i--].Delete();
            tableRecordset = null;
            columnRecordset = null;
            relationRecordset = null;

            Utilites.ScreenEvents.DoEvents();

            // Try to use Integrated Security. If not, try Microsoft's MSI Security 
            try
            {
                string connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Data Source=" + getServer + ";Initial Catalog=" + getDB + ";";

                var queryString = "select  Table_Name = TABLE_SCHEMA + '.' + TABLE_NAME, Table_Type=TABLE_TYPE, Description=null, Application=null, Display_Name=TABLE_NAME, [Database]=TABLE_CATALOG from Information_Schema.Tables";
                //            tableRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connectionString, queryString, 0, "Tables");
                tableRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connectionString, queryString, 0, "Tables");

                // make sure the External Data Window is open
                vApplication.ActiveWindow.Windows.ItemFromID[2044].Visible = true;
                Utilites.ScreenEvents.DoEvents();

                queryString = "select Table_Name = col.Table_schema + '.' + col.table_name, col.Column_Name, col.Data_Type, Description = null, col.Ordinal_Position, " +
                    " Is_Primary = iif(cu.column_name is null, 0, 1), col.Character_Maximum_Length, col.Character_Octet_Length, col.Numeric_Precision, col.Numeric_Precision_Radix, col.Numeric_Scale, Date_Time_Precision=col.DateTime_Precision, DB_NAME() [Database] " +
                    " from INFORMATION_SCHEMA.columns col left join INFORMATION_SCHEMA.KEY_COLUMN_USAGE cu " +
                    " on col.table_schema = cu.table_schema and col.table_name = cu.table_name and col.column_name = cu.column_name and OBJECTPROPERTY(OBJECT_ID(cu.CONSTRAINT_SCHEMA +'.' + QUOTENAME(cu.CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND cu.TABLE_NAME = 'TableName' AND cu.TABLE_SCHEMA = 'Schema'";

                columnRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connectionString, queryString, 0, "Columns");
                Utilites.ScreenEvents.DoEvents();

                queryString = "select schema_name(fk_tab.schema_id) + '.' + fk_tab.name as  TableMany," +
                    "    schema_name(pk_tab.schema_id) + '.' + pk_tab.name as  TableOne," +
                    "    substring(column_names, 1, len(column_names) - 1) as [ColumnMany]," +
                    "    fk.name as Relation_Name, DB_NAME() as [Database], 'Foreign Key' as Connection_Type" +
                    " from sys.foreign_keys fk" +
                    "    inner join sys.tables fk_tab on fk_tab.object_id = fk.parent_object_id" +
                    "    inner join sys.tables pk_tab on pk_tab.object_id = fk.referenced_object_id" +
                    " cross apply(select col.[name] + ', '" +
                    "          from sys.foreign_key_columns fk_c" +
                    "              inner join sys.columns col on fk_c.parent_object_id = col.object_id" +
                    "                  and fk_c.parent_column_id = col.column_id" +
                    "          where fk_c.parent_object_id = fk_tab.object_id" +
                    "            and fk_c.constraint_object_id = fk.object_id" +
                    "                  order by col.column_id for xml path ('') ) D(column_names)" +
                    " order by schema_name(fk_tab.schema_id) + '.' + fk_tab.name," +
                    "    schema_name(pk_tab.schema_id) + '.' + pk_tab.name";

                relationRecordset = vApplication.ActiveDocument.DataRecordsets.Add(connectionString, queryString, 0, "Relations");
                Utilites.ScreenEvents.DoEvents();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                ConnectToSQLServerMSI(txtServer, txtDatabase);
                sqlConnection.Open();

                CreateTableDataRecordsetFromSQL();
                CreateColumnDataRecordsetFromSQL();
                CreateRelationsDataRecordsetFromSQL();

                sqlConnection.Close();
            }
            vApplication.ActiveWindow.Windows.ItemFromID[2044].SelectedDataRecordset = tableRecordset;
            Utilites.ScreenEvents.DoEvents();
        }

    }
}
