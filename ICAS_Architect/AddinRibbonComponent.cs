using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using Visio = Microsoft.Office.Interop.Visio;

namespace ICAS_Architect
{
    public partial class ICASArchitect
    {

        private void buttonCommand1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.LinkToRepository();
        }

        private void buttonToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TogglePanel();
        }

        //private void button2_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Globals.ThisAddIn.TogglePanel();
        //}

        private void buttonCommand1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.LinkToRepository();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            
            Globals.ThisAddIn.AddShape();
        }

        private void btnShowEntities_Click(object sender, RibbonControlEventArgs e)
        {
//            SharepointManager sharepointManager = new SharepointManager();
//            sharepointManager.CreateSharepointLists();
            Globals.ThisAddIn.sharepointManager.CreateSharepointLists();
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDownloadDynamics_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.LinkToRepository();
        }

        private void btnLinkToRepository_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.sharepointManager.GetClientContext(true,true);
            Visio.Document docStencle = Globals.ThisAddIn.Application.Documents.OpenEx("ICAS Data Architect.vssx", (short)6);
            //GG Now we need to link the table information
            //Globals.ThisAddIn.sharepointManager.getAllTableDataRecordsets();
        }

        private void btnUploadToSP_Click(object sender, RibbonControlEventArgs e)
        {
//            SharepointManager sharepointManager = new SharepointManager();
//            sharepointManager.UploadTableChanges();
            
            Globals.ThisAddIn.sharepointManager.UploadToSharePoint();
        }

        private void btnOptionSet_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnImportDB_Click(object sender, RibbonControlEventArgs e)
        {
            frmDBEntities frm = new frmDBEntities();
            frm.Show();
        }

        private void cboDB_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnImportJSON_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.drawingManager.ImportJSONFile();
            Globals.ThisAddIn.drawingManager.LoadIntoDataTable();
        }

        private void btnImportDBm_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.drawingManager.ImportFromSQLServer();
//            Globals.ThisAddIn.drawingManager.LoadIntoDataTable();
        }

        private void btnImportFromRepository_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void chkShowAttributes_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDefineDataFlow_Click(object sender, RibbonControlEventArgs e)
        {
            frmDataFlow frm = new frmDataFlow();
            frm.Show();
        }
    }
}
