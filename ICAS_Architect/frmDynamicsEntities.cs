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


namespace ICAS_Architect
{
    public partial class frmDynamicsEntities : Form
    {

        const double SHDW_PATTERN = 0;
        const double BEGIN_ARROW_MANY = 29;
        const double BEGIN_ARROW = 0;
        const double END_ARROW = 29;
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

        Visio.Application vApplication = null;
        private frmGraphViewer graphViewer = null;              // ER viewer form
        private readonly Visio.Window _window;
        private ERInformation ERI = null;

        public frmDynamicsEntities()
        {
            InitializeComponent();
            vApplication = Globals.ThisAddIn.Application;
        }

        public frmDynamicsEntities(Visio.Window window)
        {
            _window = window;
            InitializeComponent();
            vApplication = Globals.ThisAddIn.Application;
        }

        private Visio.Shape GetShapeOnPage(string ShapeName)
        {
            foreach(Visio.Shape shp in Globals.ThisAddIn.Application.ActivePage.Shapes )
            {
                if (shp.NameU == ShapeName)
                {
                    return shp;
                }
            }
            return null;
        }

        private void btnDrawChart_Click(object sender, EventArgs e)
        {

            int i = 1;
            Visio.Page Pg = Globals.ThisAddIn.Application.ActivePage;

            // Add relations if the draw relations box is checked;
            if (this.chkDrawRelations.Checked)
            {
                AddRelations();
                this.chkDrawRelations.Checked = false;
            }

            // For each item selected
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked)
                {
                    // First add each box where needed unless it already exists
                    Visio.Shape ShapeOne = GetShapeOnPage(item.Text);
                    if (ShapeOne is null)
                    {
                        ShapeOne = Pg.DrawRectangle(i, i, i + 1.6, i + .6);

                        if (item.Text.CompareTo("systemuser") == 0 | item.Text.CompareTo("Empty") == 0 | item.Text.CompareTo("businessunit") == 0)
                            ShapeOne.NameU = item.Text;// item.SubItems[2].Text;
                        else
                            ShapeOne.NameU = item.Text;// item.SubItems[2].Text;

                        ShapeOne.Text = item.Text;
                    }
                }
            }
            Utilites.ScreenEvents.DoEvents();

            // Now add their relations
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked)
                {
                    Visio.Shape ShapeOne = GetShapeOnPage(item.Text);
                    List<ERRelation> AllRelations = ERI.ERRelations.FindAll(v => v.EntityOne.CompareTo(item.Text) == 0);
                    foreach (ERRelation rel in AllRelations)
                    {
                        try
                        {
                            Visio.Shape ShapeMany = GetShapeOnPage(rel.EntityMany);

                            //Other shape does not yet exist
                            if (ShapeMany is null) continue;

                            // connector already exists
                            if (!(GetShapeOnPage(ShapeOne.NameU + "-" + ShapeMany.NameU) is null)) continue;

                            this.DrawDirectionalDynamicConnector(ShapeOne, ShapeMany, false);
                            Debug.WriteLine("Added {0}", i);
                            i = i + 1;


                            if (i >= 500) return;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("Exception in adding Relation {0} {1} {2}", rel.EntityOne, rel.EntityMany, ex.Message);
                        }
                    }
                }
            }
            Pg.Layout();
            Pg.ResizeToFitContents();
        }


            private void DrawDirectionalDynamicConnector(Visio.Shape shapeFrom, Visio.Shape shapeTo, bool isManyToMany)
        {
            Visio.Application app = Globals.ThisAddIn.Application;

            // Add a dynamic connector to the page.
            Visio.Shape connectorShape = shapeFrom.ContainingPage.Drop(app.ConnectorToolDataObject, 0.0, 0.0 );
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowShapeLayout, (short)Visio.VisCellIndices.visSLOLineRouteExt).ResultIU = 2;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowShapeLayout, (short)Visio.VisCellIndices.visSLORouteStyle).ResultIU = 1;
            //.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = "1"
            //.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = "16"
            // Set the connector properties, using different arrows, colors, and patterns for many-to-many relationships.
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visFillShdwPattern).ResultIU = SHDW_PATTERN;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineBeginArrow).ResultIU = isManyToMany ? BEGIN_ARROW_MANY : BEGIN_ARROW;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineEndArrow).ResultIU = END_ARROW;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineColor).ResultIU = isManyToMany ? LINE_COLOR_MANY : LINE_COLOR;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLinePattern).ResultIU = isManyToMany ? LINE_PATTERN : LINE_PATTERN;
            connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visLineRounding).ResultIU = ROUNDING;

            // Connect the starting point.
            Visio.Cell cellBeginX = connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DBeginX);
            cellBeginX.GlueTo(shapeFrom.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));

            // Connect the ending point.
            Visio.Cell cellEndX = connectorShape.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DEndX);
            cellEndX.GlueTo(shapeTo.get_CellsSRC(VISIO_SECTION_OJBECT_INDEX, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));

            connectorShape.NameU = shapeFrom.NameU + "-" + shapeTo.NameU;
        }

        private string GetSelectedEntityNameInDiagram()
        {
            foreach (ListViewItem item in this.listView1.Items)
                if (item.Checked)
                    return item.Text;
            return null;
        }

        //Adds relation to whatever has been selected in the list view
        private void AddRelations()
        {
            string entityName = GetSelectedEntityNameInDiagram();
            if (entityName == null) return;

            StringBuilder addedEntities = new StringBuilder(" ");
            List<string> relatedEntities = ERI.FindRelatedEntities(entityName);
            int count = 0;
            foreach (ListViewItem item in this.listView1.Items)
            {
                if (!relatedEntities.Contains(item.Text) || item.Checked) continue;
                item.Checked = true;
                addedEntities.Append(item.Text + " ");
                count++;
            }
            this.listView1.Refresh();
        }

        private void btnTmpLoadData_Click(object sender, EventArgs e)
        {
            if (graphViewer != null) return;

            DialogResult result = openJsonDialog.ShowDialog();
            if (result != DialogResult.OK) return;

            string jsonFile = openJsonDialog.FileName;
//            labelStatus.Text = $"Loading {Path.GetFileName(jsonFile)}.";

            //GG: No idea why he added the JSON reader into the Graph viewer originally. Pull it out.
            graphViewer = new frmGraphViewer();
            graphViewer.LoadMetadataJsonFile(jsonFile);
            ERI = graphViewer.entityRelations;
            this.listView1.Items.Clear();
            this.listView1.Refresh();

            // Create columns for the items and subitems.
            // Width of -2 indicates auto-size.
            //listView1.Columns.Add("Entity", -2,  HorizontalAlignment.Left);
            //listView1.Columns.Add("Description", -2, HorizontalAlignment.Left);
            
            this.listView1.Refresh();

            string sPrior = "";
            // First add all of our entities
            foreach (EREntity item in ERI.EREntities)
            {
                if (sPrior.CompareTo(item.entityLogicalName) != 0)
                {
                    ListViewItem lvitem = new ListViewItem(new[] { item.entityLogicalName, item.descriptiontext, item.metadataId });
                    this.listView1.Items.Add(lvitem);

                    sPrior = item.entityLogicalName;
                }
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            //SharepointManager sharepointManager = new SharepointManager();
            //sharepointManager.UploadERInformation(this.ERI);
            Globals.ThisAddIn.sharepointManager.UploadERInformation(this.ERI);
        }

        DataSet dataSet = null;

        internal void LoadIntoDataTable()
        {

            // Weirdly, DataRecordsets are enumerated from 1
            for (int i = 1; i <= vApplication.ActiveDocument.DataRecordsets.Count; i++)
                if (vApplication.ActiveDocument.DataRecordsets[i].Name == "Tables" | vApplication.ActiveDocument.DataRecordsets[i].Name == "Columns" | vApplication.ActiveDocument.DataRecordsets[i].Name == "Relations")
                    vApplication.ActiveDocument.DataRecordsets[i--].Delete();
            Utilites.ScreenEvents.DoEvents();

            string dts = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'>" +
             " <s:Schema id='RowsetSchema'>" +
             " <s:ElementType name='row' content='eltOnly' rs:updatable='true'>" +
             " <s:AttributeType name='c1' rs:name='Table Name' rs:number='2' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c2' rs:name='Database' rs:number='3' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c3' rs:name='Description' rs:number='4' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:AttributeType name='c4' rs:name='Hash' rs:number='5' rs:nullable='true' rs:maydefer='true' rs:write='true'> <s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> </s:AttributeType>" +
             " <s:extends type='rs:rowbase'/>" +
             " </s:ElementType>" +
             " </s:Schema>" +
             " <rs:data>";
            
            for(int i = 0; i < ERI.EREntities.Count; i++)// ERI.EREntities.Count; i++)
                dts = dts + " <z:row c1='" + ERI.EREntities[i].entityLogicalName + "' c2='" + ERI.EREntities[i].databaseName + "' c3='" + SecurityElement.Escape(ERI.EREntities[i].descriptiontext) + "' c4='" + ERI.EREntities[i].metadataId + "'/>";
            dts = dts +             
                " </rs:data>" +
                " </xml>";

            var dataRecordset = vApplication.ActiveDocument.DataRecordsets.AddFromXML(dts, 0, "Tables");
            Utilites.ScreenEvents.DoEvents();

            // make sure the External Data Window is open
            vApplication.ActiveWindow.Windows.ItemFromID[2044].Visible = true;
            Utilites.ScreenEvents.DoEvents();

        }


        private void btnLoadSQL_Click(object sender, EventArgs e)
        {
            dataSet = new DataSet();

            string connectionString = "Data Source=edi-dwuat-tmp;Initial Catalog=DataVault;Integrated Security=true";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {

                string queryString = "Select * from Information_Schema.Tables";
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataReader sqlDataReader = command.ExecuteReader();

                DataTable dataTable = new DataTable("Tables");
                dataTable.Load(sqlDataReader);
                dataTable.Columns.Add("MetadataID", typeof(String));
                dataSet.Tables.Add(dataTable);
                


                queryString = "select col.*, IsPrimary = iif(cu.column_name is null, 0, 1) from INFORMATION_SCHEMA.columns col left join INFORMATION_SCHEMA.KEY_COLUMN_USAGE cu on " +
                    " col.table_schema = cu.table_schema and col.table_name = cu.table_name and col.column_name = cu.column_name and " +
                    "   OBJECTPROPERTY(OBJECT_ID(cu.CONSTRAINT_SCHEMA + '.' + QUOTENAME(cu.CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND cu.TABLE_NAME = 'TableName' AND cu.TABLE_SCHEMA = 'Schema'";
                command = new SqlCommand(queryString, connection);
 //               connection.Open();
                sqlDataReader = command.ExecuteReader();

                dataTable = new DataTable("Columns");
                dataTable.Load(sqlDataReader);
                dataTable.Columns.Add("MetadataID", typeof(String));
                dataSet.Tables.Add(dataTable);



                queryString = "select schema_name(fk_tab.schema_id) + '.' + fk_tab.name as foreign_table," +
                    "    schema_name(pk_tab.schema_id) + '.' + pk_tab.name as primary_table," +
                    "    substring(column_names, 1, len(column_names) - 1) as [fk_columns]," +
                    "    fk.name as fk_constraint_name" +
                    " from sys.foreign_keys fk" +
                    "    inner join sys.tables fk_tab" +
                    "       on fk_tab.object_id = fk.parent_object_id" +
                    "   inner join sys.tables pk_tab " +
                    "       on pk_tab.object_id = fk.referenced_object_id" +
                    " cross apply(select col.[name] + ', '" +
                    "          from sys.foreign_key_columns fk_c" +
                    "              inner join sys.columns col" +
                    "                  on fk_c.parent_object_id = col.object_id" +
                    "                  and fk_c.parent_column_id = col.column_id" +
                    "          where fk_c.parent_object_id = fk_tab.object_id" +
                    "            and fk_c.constraint_object_id = fk.object_id" +
                    "                  order by col.column_id" +
                    "                            for xml path ('') ) D(column_names)" +
                    " order by schema_name(fk_tab.schema_id) + '.' + fk_tab.name," +
                    "    schema_name(pk_tab.schema_id) + '.' + pk_tab.name";

                command = new SqlCommand(queryString, connection);
  //              connection.Open();
                sqlDataReader = command.ExecuteReader();

                dataTable = new DataTable("Relations");
                dataTable.Load(sqlDataReader);
                dataTable.Columns.Add("MetadataID", typeof(String));
                dataSet.Tables.Add(dataTable);
            }
        }

        private void btnXMLLoad_Click(object sender, EventArgs e)
        {
            LoadIntoDataTable();
        }
    }
}

