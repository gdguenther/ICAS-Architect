// Utilities.cs
// compile with: /doc:Utilities.xml
// <copyright>Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <summary>This file contains the implementation of Utilities class.</summary>

using System;
using System.Diagnostics;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;
using Microsoft.Win32;
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.Samples.Visio.FlowchartAddIn.CSharp {

    /// <summary>Lists the fields in the ownerColor lookup table.
    /// </summary>
    public enum OwnerColorField
    {

        /// <summary>Index for the Owner field</summary>
        OwnerColorOwner = 0,

        /// <summary>Index for the Color field</summary>
        OwnerColorColor = 1
    }

    /// <summary>This class provides all the constants and methods that are
    /// shared among the classes in the project.</summary>
    [ComVisible(false)]
    public sealed class Utilities
    {

        /// <summary>Comma character is employed as the separator because
        /// Universal functions (i.e. FormulaU) are being used.</summary>
        public const string Separator = ",";

        /// <summary>Minimum version of Visio supported - Microsoft Visio 2013
        /// </summary>
        public const short MinVisioVersion = 15;

        /// <summary>Column in Excel worksheets containing the step ID
        /// </summary>
        public const short ExcelColumnStepId = 1;

        /// <summary>Column in Excel worksheets containing the task name
        /// </summary>
        public const short ExcelColumnTask = 2;
        
        /// <summary>Column in Excel worksheets containing the next steps
        /// </summary>
        public const short ExcelColumnNextSteps = 3;

        /// <summary>Column in Excel worksheets containing the owner name
        /// </summary>
        public const short ExcelColumnOwner = 4;

        /// <summary>Column in Excel worksheets containing the step type
        /// </summary>
        public const short ExcelColumnStepType = 5;

        /// <summary>Column in Excel worksheets containing the duration
        /// </summary>
        public const short ExcelColumnDuration = 6;

        /// <summary>Column in Excel worksheets containing the comments
        /// </summary>
        public const short ExcelColumnComments = 7;

        /// <summary>Column in Excel worksheets containing the hyperlink
        /// address</summary>
        public const short ExcelColumnHyperlink = 8;

        /// <summary>Column in Excel worksheets containing the hyperlink
        /// description</summary>
        public const short ExcelColumnHyperlinkDescription = 9;

        /// <summary>The first row in the Sample Excel Spreadsheet that has 
        /// data</summary>
        public const short ExcelRowStart = 2;

        /// <summary>Name of the flowchart stencil this addin uses</summary>
        public const string FlowchartStencil = "BASFLO_M.VSSX";

        /// <summary>Name of the HTML file that is generated</summary>
        public const string HtmlFileName = "SampleFlowchart.html";

        /// <summary>Name of the Excel file that is used</summary>
        public const string ExcelFileName = "Flowchart.xlsx";

        /// <summary>Command ID representing a document created event. Used
        /// by marker event context strings.</summary>
        public const short DocumentCreateCommandId = 1;

        /// <summary>Command ID representing a document opened event. Used
        /// by marker event context strings.</summary>
        public const short DocumentOpenCommandId = 2;

        /// <summary>Command ID representing a shape RMA event. Used
        /// by marker event context strings.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Rma")]
        public const short ShapeRmaConnectsCommandId = 100;

        /// <summary>SDK Flowchart Application Format used in the Marker 
        /// event context string.</summary>
        public const string ContextSdkFlowchart = "SDKFLW_C#NET";

        /// <summary>Begin marker format used in the marker event context 
        /// string.</summary>
        public const string ContextBeginMarker = "/";

        /// <summary>Document format used in the marker event Context 
        /// string</summary>
        public const string ContextDocument = "DOC=";

        /// <summary>Event format used in the marker event context string
        /// </summary>
        public const string ContextEvent = "EVENT=";

        /// <summary>Page format used in the marker event Context string
        /// </summary>
        public const string ContextPage = "PAGE=";

        /// <summary>Shape format used in the marker event Context string
        /// </summary>
        public const string ContextShape = "SHAPE=";

        /// <summary>Solution format used in the marker event context string
        /// </summary>
        public const string ContextSolution = "SOLUTION=";

        /// <summary>Formula of the right menu action item</summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Rma")]
        public const string RmaMenuFormula =
            "RUNADDONWARGS(\"QueueMarkerEvent\"" + Separator + "\"" + 
            ContextSolution + ContextSdkFlowchart + " " 
            + ContextBeginMarker + ContextEvent + "100\")";

        /// <summary>Extension for binary Visio template files
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Vst")]
        public const string TemplateExtensionVst = "VST";

        /// <summary>Extension for XML-based Visio template files
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Vtx")]
        public const string TemplateExtensionVtx = "VTX";

        /// <summary>Extension for Visio 2013 template
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Vstx")]
        public const string TemplateExtensionVstx = "VSTX";

        /// <summary>Extension for Visio 2013 macro-enabled template
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Vstm")]
        public const string TemplateExtensionVstm = "VSTM";

        /// <summary>Create a global resource manager for the application
        /// to retrieve localized strings.</summary>
        private static ResourceManager theResourceManager = 
            new ResourceManager("Strings",
            System.Reflection.Assembly.GetExecutingAssembly());

        /// <summary>Array of step types found in the sample Excel 
        /// spreadsheet.</summary>
        private static readonly string[] stepShapes_type = { "PROCESS", "DECISION", "END" };
        
        /// <summary>Array of Visio master names for all step types 
        /// found in the sample Excel spreadsheet. MUST match the type
        /// order in stepShapes_type</summary>
        private static readonly string[] stepShapes_master = {"Process", "Decision", "Start/End"};

        /// <summary>Array of owners found in the sample Excel 
        /// spreadsheet.</summary>
        private static readonly string[] ownerColor_name = { 
                LoadString("OwnerAmyLead"),
                LoadString("OwnerJohnMarketer"),
                LoadString("OwnerSarahPlanner")
        };
        
        /// <summary>Array of colors for all owners found in the 
        /// sample Excel spreadsheet. Must match the owner order
        /// in ownerColor_name</summary>
        private static readonly string[] ownerColor_value = {
                "RGB(204, 255, 153)",
                "RGB(149, 203, 221)",
                "RGB(171, 170, 225)"
        };

        /// <summary>This constructor intentionally left blank.</summary>
        private Utilities() {

            // No initialization
        }

        /// <summary>This method displays a message box containing an error
        /// if the alertResponse value is zero.</summary>
        /// <param name="alertResponse">AlertResponse value of the running
        /// Visio instance</param>
        /// <param name="exceptionMessage">Error message to be displayed
        /// </param>
        public static void DisplayException(
            int alertResponse, 
            string exceptionMessage) {

            string title;

            title = LoadString("AddInName");

            // Check AlertResponse to see whether we should display
            // modal UI. 
             if (alertResponse == 0) {
                Utilities.RtlAwareMessageBoxShow(
                    exceptionMessage, title,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else {
                Debug.WriteLine(exceptionMessage);
            }
        }

        /// <summary>This method gets the Flowchart Sample's path from the 
        /// fully qualified add-in name.</summary>
        /// <returns>Flowchart Sample path with an ending backslash (\)
        /// </returns>
        public static string FlowchartSamplePathFromAddIn {

            get {
                string path = "";
                System.Reflection.Module sampleAddIn;

                sampleAddIn = Assembly.GetExecutingAssembly().GetModules()[0];
                path = sampleAddIn.FullyQualifiedName;
                path = System.IO.Path.GetDirectoryName(path);
                path += System.IO.Path.DirectorySeparatorChar;

                return path;
            }
        }

        /// <summary>This method gets the Flowchart Sample's path from the 
        /// registry.</summary>
        /// <returns>Flowchart Sample path with an ending backslash (\)
        /// </returns>
        public static string FlowchartSamplePathFromRegistry {

            get {
                string path = "";
                
                RegistryKey regKey = Registry.LocalMachine.OpenSubKey(
                    "SOFTWARE\\Microsoft\\Office\\16.0\\Visio");
                System.Diagnostics.Debug.Assert(regKey != null);
                path = (string)regKey.GetValue("SDKPath");
                
                path = path + "Samples\\Flowchart\\CSharp\\";
                return path;
            }
        }    

        /// <summary>This method gets the SDK Flowchart Sample Excel workbook.
        /// </summary>
        /// <param name="excelWorkbooks">The Microsoft Excel
        /// application workbooks object.</param>
        /// <param name="excelFileName">Name of the Excel workbook to open
        /// and return.
        /// </param>
        /// <returns>The SDK Flowchart Sample Excel workbook.
        /// </returns>
//        [CLSCompliant(false)]
        public static Microsoft.Office.Interop.Excel.Workbook
            GetSampleExcelWorkbook (
            Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks,
            string excelFileName) {

            if (excelWorkbooks == null)
                return null;

            if (excelFileName == null)
                return null;

            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;

            excelWorkbook = excelWorkbooks.Open(
                excelFileName, 0, true, 1, "", "", true,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, 
                (char)9, false, false, null, false, false, 
                Microsoft.Office.Interop.Excel.XlCorruptLoad.xlNormalLoad);

            return excelWorkbook;
        }

        /// <summary>This method gets the document from the documents 
        /// collection.</summary>
        /// <param name="theDocuments">Document collection from which
        /// the document is to be found</param>
        /// <param name="documentName">Name of the document to be found
        /// </param>
        /// <returns>Document object if found; otherwise null</returns>
//        [CLSCompliant(false)]
        public static Document GetDocumentItem(
            Documents theDocuments,
            string documentName) {

            if (documentName == null)
                return null;

            if (theDocuments == null)
                return null;

            Document currentDocument;
            try {
                currentDocument = theDocuments[documentName];
            }
            catch (COMException) {
                // Document was not found. Ignore the error.
                currentDocument = null;
            }

            return currentDocument;
        }

        /// <summary>This method finds the master specified by masterNameU.
        /// </summary>
        /// <param name="currentDocument">Document object from which the
        /// master is to be found</param>
        /// <param name="masterNameU">Name of the master to be found</param>
        /// <returns>Master object if found; otherwise null</returns>
//        [CLSCompliant(false)]
        // Parameter "masterNameU" is not localizable.
        [SuppressMessage("Microsoft.Globalization", "CA1303:DoNotPassLiteralsAsLocalizedParameters")]
        public static Master GetMasterItem(
            IVDocument currentDocument,
            string masterNameU) {

            if (currentDocument == null)
                return null;

            if (masterNameU == null)
                return null;

            if (currentDocument.Masters == null)
                return null;

            try {
                // Get a master on the stencil by its universal name.
                return currentDocument.Masters.get_ItemU(masterNameU);
            }
            catch (COMException) {
                return null;
            }
       }

        /// <summary>Accessor for the ownerColor field.</summary>
        /// <param name="index">Index to find the corresponding 
        /// value for</param>
        /// <param name="field">What to return: Owner or Color</param>
        /// <returns>Owner or Color, depending on field param</returns>
        public static string GetOwnerColor(
            int index, 
            OwnerColorField field) {

            return (field == OwnerColorField.OwnerColorOwner) ?
                    ownerColor_name[index] : ownerColor_value[index];
        }

        /// <summary>This method finds the requested shape in the shapes 
        /// collection passed in.</summary>
        /// <param name="shapesInPage">Shapes collection</param>
        /// <param name="id">ID of the shape to be found</param>
        /// <returns>Shape object if found; otherwise null</returns>
//        [CLSCompliant(false)]
        public static Shape GetShapeItem(
            Shapes shapesInPage, 
            object id) {

            Shape currentShape;

            if ( shapesInPage == null ) {
                return null;
            }

            try {
                currentShape = shapesInPage[id];
            }
            catch (COMException) {
                // Shape was not found. Ignore the error.
                currentShape = null;
            }

            return currentShape;
        }

        /// <summary>This method finds the stencil in the Documents 
        /// collection. If the stencil is not found, this method loads it 
        /// by calling the OpenEx method of the Documents collection.
        /// </summary>
        /// <param name="applicationDocuments">Documents collection</param>
        /// <param name="stencilName">Name of the stencil</param>
        /// <returns>Document object that corresponds to the stencil name
        /// or nothing if the stencil cannot be found</returns>
//        [CLSCompliant(false)]
        public static Document GetStencil(
            Documents applicationDocuments,
            string stencilName) {

            if (applicationDocuments == null)
                return null;

            if (applicationDocuments.Application == null)
                return null;

            Document stencil = null;
            try {

                stencil = GetDocumentItem(applicationDocuments, stencilName);

                // The stencil is not loaded, so load it.
                if (stencil == null) {
                    stencil = applicationDocuments.OpenEx(stencilName,
                        (short)(short)VisOpenSaveArgs.visOpenRO
                        + (short)VisOpenSaveArgs.visOpenDocked);
                }
            }
            catch (COMException) {
                string error = "Stencil not found: \r\n" + stencilName;
                DisplayException(
                    applicationDocuments.Application.AlertResponse,
                    error);
            }

            return stencil;
        }

        /// <summary>Accessor for the stepShapes field.</summary>
        /// <param name="index">Step type index to find the corresponding 
        /// shape for</param>
        /// <returns>Shape type</returns>
        public static string GetStepShapes(int index) {
            
            return stepShapes_master[index];
        }

        /// <summary>Accessor for the types in the stepShapes field.</summary>
        /// <param name="index">Step type index to find the corresponding 
        /// type for</param>
        /// <returns>Step type</returns>
        public static string GetStepTypes(int index) {
            
            return stepShapes_type[index];
        }

        /// <summary>This method loads a string from the embedded resource.
        /// </summary>
        /// <param name="resourceName">Name of the resource to be loaded
        /// </param>
        /// <returns>Loaded resource string if successful, otherwise empty
        /// string</returns>
        public static string LoadString(string resourceName) {

            return theResourceManager.GetString(resourceName,
                System.Globalization.CultureInfo.CurrentUICulture);
        }

        /// <summary>Returns the size of the ownerColor lookup table.
        /// </summary>
        public static int OwnerColorCount {
            get {
                return ownerColor_name.GetLength(0);
            }
        }

        /// <summary>Returns the size of the stepShapes lookup table.
        /// </summary>
        public static int StepShapeCount {
            get {
                return stepShapes_type.GetLength(0);
            }
        }

        /// <summary>This method converts the input string to a Formula for a
        /// string by replacing each double quotation mark (") with a pair of
        /// double quotation marks ("") and adding double quotation marks 
        /// ("") around the entire string. When this formula is assigned to 
        /// the formula property of a Visio cell it will produce a result 
        /// value equal to the string, input.</summary>
        /// <param name="input">Input string that will be processed
        /// </param>
        /// <returns>Formula for input string</returns>
        public static string StringToFormulaForString(string input) {

            const string quote = "\"";
            string  result = "";

            if (input == null) {
                return null;
            }

            // Replace all (") with ("").
            result = input.Replace(quote,
                (quote + quote));

            // Add ("") around the entire string.
            result = quote + result + quote;

            return result;
        }

        /// <summary>class for right-to-left aware message box</summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Rtl")]
        public static DialogResult RtlAwareMessageBoxShow(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            // Ask the CurrentUICulture if we are running under right-to-left.
            MessageBoxOptions options = 0;
            if (System.Globalization.CultureInfo.CurrentUICulture.TextInfo.IsRightToLeft) {

                options |= MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign;
            }

            return MessageBox.Show(null, text, caption,
                buttons, icon, MessageBoxDefaultButton.Button1, options);
        }

        public static DialogResult ShowInputDialog(ref string input)
        {
            System.Drawing.Size size = new System.Drawing.Size(200, 70);
            Form inputBox = new Form();

            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Name";

            System.Windows.Forms.TextBox textBox = new TextBox();
            textBox.Size = new System.Drawing.Size(size.Width - 10, 23);
            textBox.Location = new System.Drawing.Point(5, 5);
            textBox.Text = input;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, 39);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 80, 39);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            DialogResult result = inputBox.ShowDialog();
            input = textBox.Text;
            return result;
        }
    }
}
