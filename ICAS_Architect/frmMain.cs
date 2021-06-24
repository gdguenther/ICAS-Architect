using System;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using VA = VisioAutomation;


namespace ICAS_Architect
{
    public partial class frmMain : Form
    {

//        public sealed class RetrieveOptionSetRequest : Microsoft.Xrm.Sdk.OrganizationRequest

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

        delegate void SetButtonEnabledCallback();
        delegate void SetButtonTextCallback(string text);
        delegate void SetLabelTextCallback(string text);
        private readonly Visio.Window _window;

        private const string DOWNLOAD = "Download Metadata";
        private const string CANCEL = "Cancel";

        private const string SAMPLE_URL_ONLINE = "https://dev-icas.crm11.dynamics.com/";
        private const string DEFAULT_STATUS = "Please enter the organization url then push Download Metadata.";

        private MetadataDownloader metadataDownloader = null;   // controller to initiate the download metadata and parse the response
        private frmGraphViewer graphViewer = null;              // ER viewer form

        public frmMain()
        {
            InitializeComponent();
            txtBaseUrl.Text = SAMPLE_URL_ONLINE;
            labelStatus.Text = DEFAULT_STATUS;
            // if you only need the trigger information html, you can make this check box visible and untick to avoid scheme download
            cbFullSpec.Visible = false;
        }

        public frmMain(Visio.Window window)
        {
            _window = window;
            InitializeComponent();
            txtBaseUrl.Text = SAMPLE_URL_ONLINE;
            labelStatus.Text = DEFAULT_STATUS;
            // if you only need the trigger information html, you can make this check box visible and untick to avoid scheme download
            cbFullSpec.Visible = false;
        }

        private void DownloadMetadata()
        {
            if (!CheckUrlTextField()) return;
            if (string.IsNullOrWhiteSpace(folderBrowserDialog1.SelectedPath) || !Directory.Exists(folderBrowserDialog1.SelectedPath))
            {
                if (DialogResult.OK != folderBrowserDialog1.ShowDialog()) return;
            }
            if (!Directory.Exists(folderBrowserDialog1.SelectedPath))
            {
                labelStatus.Text = $"{folderBrowserDialog1.SelectedPath} does not exist.";
                return;
            }

            btnGraph.Enabled = false;
            btnDumpMetadata.Text = CANCEL;
            if (graphViewer != null) graphViewer.Close();

            // Now start to run the spec dump procedures.
            metadataDownloader = new MetadataDownloader(txtBaseUrl.Text, folderBrowserDialog1.SelectedPath, cbFullSpec.Checked);

            string endpointVersion = ConfigurationManager.AppSettings["ENDPOINT_VERSION"];
            if (!string.IsNullOrWhiteSpace(endpointVersion)) metadataDownloader.EndpointVersion = endpointVersion;

            // use the adal library
            StartDownloadMetadataUsingAdalAndBackgroundWorker();
        }

        private void StartDownloadMetadataUsingAdalAndBackgroundWorker()
        {
            if (adalDownloadBGworker.IsBusy) return;
            adalDownloadBGworker.WorkerReportsProgress = true;
            adalDownloadBGworker.RunWorkerAsync();
        }

        private void adalDownloadBGworker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                DownloadMetadataUsingAdal();
            }
            catch (Exception ex)
            {
                metadataDownloader.ExceptionCaught = ex;
            }
        }

        private void DownloadMetadataUsingAdal()
        {
            adalDownloadBGworker.ReportProgress(0);
            HttpDownloadClient httpDownloadClient = new HttpDownloadClient(txtBaseUrl.Text);
            httpDownloadClient.Connect(metadataDownloader.WHOAMIURL);
            string nextWebApiUrl = metadataDownloader.GetNextWebApiUrl();
            while (nextWebApiUrl != null && !metadataDownloader.Cancelled)
            {
                string statusText = $"Downloading {new Uri(nextWebApiUrl).AbsolutePath}";
                adalDownloadBGworker.ReportProgress(50, statusText);
                string content = httpDownloadClient.Fetch(nextWebApiUrl);
                metadataDownloader.HandleResponse(content);
                nextWebApiUrl = metadataDownloader.GetNextWebApiUrl();
            }
        }


        private void adalDownloadBGworker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                string txt = "Testing connection and whoami";
                SetText(txt.ToString());
                //GG:                labelStatus.Text = "Testing connection and whoami";
            }
            else
            {
                string entityDownloadProgress = metadataDownloader.FetchAllEntitiesProgressText();
                string txt = (string.IsNullOrWhiteSpace(entityDownloadProgress) ? e.UserState as string : entityDownloadProgress);
                SetText(txt.ToString());
                //GG:               labelStatus.Text = (string.IsNullOrWhiteSpace(entityDownloadProgress) ? e.UserState as string : entityDownloadProgress);
            }
        }

        private void adalDownloadBGworker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            DownloadCompleted();
        }

        private void DownloadCompleted()
        {
            string txt = null;
            if (metadataDownloader.ExceptionCaught != null)
            {
                txt = "Exception caught during download";
                SetText(txt.ToString());
                MessageBox.Show(metadataDownloader.ExceptionCaught.ToString());
                Cleanup();
                return;
            }

            if (metadataDownloader.Cancelled)
            {
                txt = "Download Cancelled.";
                SetText(txt.ToString());
                Cleanup();
                return;
            }


            txt = $"Download completed. Output {metadataDownloader.OutputFolderPath}";
            SetText(txt.ToString());

            // Assign our Entity Relations back to the drawing Manager
            Globals.ThisAddIn.drawingManager.ERI = metadataDownloader.GetERInformation();// graphViewer.entityRelations;
            Globals.ThisAddIn.drawingManager.LoadIntoDataTable();
            ShowReport();
            Cleanup();
        }

        private void ShowReport()
        {
/*          This is unneeded as we'd prefer to show in Visio. It's not working at the moment even when I try to pass it back to the core process - not sure why.
           string jsonFile = $"{metadataDownloader.OutputFileFullPath}.json";
            if (!File.Exists(jsonFile))
            {
                return;
            }

            graphViewer = new frmGraphViewer();
            graphViewer.LoadMetadataJsonFile(jsonFile);
            graphViewer.Show();
            graphViewer.FormClosing += GraphViewer_FormClosing;*/
        }

        private void Cleanup()
        {
            metadataDownloader = null;
            SetButtonStatus(DOWNLOAD);
            //GG: btnDumpMetadata.Text = DOWNLOAD;
            SetButtonsEnabled();
        }

        private void btnDumpMetadata_Click(object sender, EventArgs e)
        {
            if (CANCEL.Equals(btnDumpMetadata.Text) && metadataDownloader != null)
            {
                SetButtonStatus("Cancelling");
                //GG: btnDumpMetadata.Text = "Cancelling";
                metadataDownloader.Cancelled = true;
                return;
            }

            DownloadMetadata();
        }

        private void btnGraph_Click(object sender, EventArgs e)
        {
            // show the empty graph viewer so that user can load graph file directly
            if (graphViewer != null) return;

            DialogResult result = openJsonDialog.ShowDialog();
            if (result != DialogResult.OK) return;

            string jsonFile = openJsonDialog.FileName;
            labelStatus.Text = $"Loading {Path.GetFileName(jsonFile)}.";
            graphViewer = new frmGraphViewer();
            graphViewer.LoadMetadataJsonFile(jsonFile);
            graphViewer.Show();
            graphViewer.FormClosing += GraphViewer_FormClosing;
        }


        private void GraphViewer_FormClosing(object sender, FormClosingEventArgs e)
        {
            graphViewer = null;
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (graphViewer == null) return;
            graphViewer.Close();
        }


        private bool CheckUrlTextField()
        {
            string organizationUrl = txtBaseUrl.Text.Trim();
            bool urlWellFormedAndHttp = Uri.IsWellFormedUriString(organizationUrl, UriKind.Absolute) && Uri.TryCreate(organizationUrl, UriKind.Absolute, out Uri uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
            if (!urlWellFormedAndHttp)
            {
                string txt = DEFAULT_STATUS;
                SetText(txt.ToString());
                //GG:           labelStatus.Text = DEFAULT_STATUS;
                return false;
            }
            if (organizationUrl.EndsWith("/")) organizationUrl = organizationUrl.Substring(0, organizationUrl.Length - 1);
            txtBaseUrl.Text = organizationUrl;
            return true;
        }

        // These three functions check to ensure that we are being called from the same thread.
        // If they are not the same thread, we send a message to the the callback function.
        // InvokeRequired Required returns True if it's the wrong thread.
        private void SetText(string text)
        {
            if (this.labelStatus.InvokeRequired)
            {
                SetLabelTextCallback d = new SetLabelTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.labelStatus.Text = text;
            }
        }

        private void SetButtonStatus(string text)
        {
            if (this.btnDumpMetadata.InvokeRequired)
            {
                SetButtonTextCallback d = new SetButtonTextCallback(SetButtonStatus);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.btnDumpMetadata.Text = text;
            }
        }

        private void SetButtonsEnabled()
        {
            if (this.btnDumpMetadata.InvokeRequired)
            {
                SetButtonEnabledCallback d = new SetButtonEnabledCallback(SetButtonsEnabled);
                this.Invoke(d, new object[] {  });
            }
            else
            {
                btnDumpMetadata.Enabled = true;
                btnGraph.Enabled = true;
                if (!SAMPLE_URL_ONLINE.Equals(txtBaseUrl.Text)) return;
                txtBaseUrl.Text = SAMPLE_URL_ONLINE;


                string jsonFile = $"{metadataDownloader.OutputFileFullPath}.json";
                if (!File.Exists(jsonFile))
                {
                    return;
                }

                graphViewer = new frmGraphViewer();
                graphViewer.LoadMetadataJsonFile(jsonFile);
                graphViewer.Show();
                graphViewer.FormClosing += GraphViewer_FormClosing;

            }

        }
    }
}
