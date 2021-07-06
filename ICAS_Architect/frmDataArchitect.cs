using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ICAS_Architect
{
    public partial class frmDataArchitect : Form
    {
        SharepointManager sharepointManager = null;

        public frmDataArchitect()
        {
            InitializeComponent();
            sharepointManager = Globals.ThisAddIn.sharepointManager;
        }

        internal void FilterTables()
        {
            sharepointManager.getAllTableDataRecordsets(this.cboApplications.Text, this.cboDatabase.Text, this.chkIncludeViews.Checked, this.chkIncludeAPIs.Checked);
        }






    }
}
