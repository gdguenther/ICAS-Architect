using System;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
/*using System.Linq;
using System.Web;*/
//using Microsoft.SharePoint.Client;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;
using VA = VisioAutomation;
using Microsoft.Office.Tools.Ribbon;


namespace ICAS_Architect
{
    public partial class ThisAddIn
    {

        internal SharepointManager sharepointManager = null;
        internal OfficeRibbon ribbonref = null;

        public Microsoft.Win32.RegistryKey registryKey;
        //        private Microsoft.Office.Interop.Visio.Event resultEvent;
        //        private Microsoft.Office.Interop.Visio.IVisEventProc applicationEventSink;
        bool DeleteMessageShown = false;

        // To Do: This form is only temporarily holding our data
        public System.Windows.Forms.Form frm = null;

        bool blsICausedCellChanges = false;

        private PanelManager _panelManager;
        //        private Microsoft.Office.Interop.Visio.Document _targetDoc = null;

        internal DrawingManager drawingManager = null;

        public void LinkToRepository()
        {
            frm = new frmMain();
            frm.Show();
        }

        public void ImportDatabaseMetadata()
        {
            frm = new frmDBEntities();
            frm.Show();
        }


        public void AddShape()
        {
            var s0 = Application.ActivePage.DrawRectangle(0, 0, 1, 1);
        }

        public void TogglePanel()
        {
            _panelManager.TogglePanel(Application.ActiveWindow);
        }


        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            sharepointManager = new SharepointManager();
            drawingManager = new DrawingManager();
            _panelManager = new PanelManager(this);
            Application.MarkerEvent += Application_MarkerEvent2;
            Application.BeforeShapeDelete += Application_BeforeShapeDelete;
            Application.DocumentOpened += Application_DocumentOpened;
            Application.ShapeAdded += Application_ShapeAdded;


            //accessing the CurrentUser root element  
            //and adding "OurSettings" subkey to the "SOFTWARE" subkey  
            registryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ICAS Architect");

        }

        private void Application_DocumentOpened(Visio.Document Doc)
        {
 //           throw new NotImplementedException();
        }

        private void Application_ShapeAdded(Visio.Shape Shape)
        {
            if (blsICausedCellChanges)
                return;

            Application.QueueMarkerEvent("ScopeStart");
            //try
            {
                System.Diagnostics.Debug.WriteLine("Shape Added");
                drawingManager.AddShapeFromTableData(Shape);
            }
            //catch (Exception e)
            {
            //    Console.WriteLine(e.Message);
            }

            Application.QueueMarkerEvent("ScopeEnd");
        }

        private void Application_BeforeShapeDelete(Visio.Shape Shape)
        {
            if (!this.DeleteMessageShown & !blsICausedCellChanges & false)
            {   // only show message once (this could be used, but is disabled for the moment)
                System.Windows.Forms.MessageBox.Show("This will only delete the shape from the diagram, not from the database.");
                DeleteMessageShown = true;
            }
        }

        private void Application_MarkerEvent2(Visio.Application app, int SequenceNum, string ContextString)
        {
            //            var shape = visapp.ActiveWindow.Selection.PrimaryItem;
            if (ContextString == "ScopeStart" )
                blsICausedCellChanges = true;
            else
                blsICausedCellChanges = false;
        }


        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _panelManager.Dispose();

        }


        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;

        }


        public void btnMapVisio_Click()
        {
            Application.ActivePage.DrawOval(0, 0, 1, 1);
        }


    }
}
