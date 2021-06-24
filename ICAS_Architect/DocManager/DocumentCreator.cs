// DocumentCreator.cs
// compile with: /doc:DocumentCreator.xml
// <copyright>Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <summary>This file contains the DocumentCreator class.</summary>

using System;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;

namespace Microsoft.Samples.Visio.FlowchartAddIn.CSharp {

    /// <summary>This class creates a Visio document containing flowchart
    /// drawings, based on data contained in a Microsoft Excel 
    /// worksheet.  The drawing process is driven by the CreateDrawing method.
    /// </summary>
    ///
    /// <remark>Note: To enable this class to execute its tasks, the 
    /// workbook flowchart.xlsx must be located in the same directory as 
    /// FlowchartSampleCSharp.dll. A sample workbook is provided with this
    /// add-in.</remark>
    [ComVisible(false)]
    public class DocumentCreator {

        private bool demoMode;
        private SortedList shapeIDs;
        private Microsoft.Office.Interop.Visio.Application visioApplication;

        /// <summary>This constructor is intentionally left blank.</summary>
        public DocumentCreator() {

            // No initialization
        }

        /// <summary>This method connects the shapes on the Visio page based 
        /// on data contained in cells in an Excel worksheet.</summary>
        /// <param name="currentPage">Page that has the shapes</param>
        /// <param name="excelSheet">Worksheet that has information on how 
        /// the shapes should be connected</param>
        /// <returns>True if successful; False otherwise</returns>
        private bool connectSteps(
            Page currentPage,
            Microsoft.Office.Interop.Excel.Worksheet excelSheet) {

            bool isSuccessful = true;
            int rowIndex;
            int[] nextStepID;
            int nextStepCount = 0;
            string stepID;
            string nextSteps;
            Shape currentShape;
            Shape nextShape;

            try {

                // Connect the shapes.
                rowIndex = Utilities.ExcelRowStart;
                stepID = getExcelCellValue(excelSheet,
                    rowIndex, 
                    Utilities.ExcelColumnStepId);

                while (stepID.Length != 0) {

                    nextSteps = getExcelCellValue(excelSheet,
                        rowIndex, 
                        Utilities.ExcelColumnNextSteps);
                     nextStepID = new int[nextSteps.Length];

                    if (parseNextSteps(nextSteps, 
                        ref nextStepID, 
                        ref nextStepCount)) {

                        if (nextStepCount > 0) {

                            currentShape = null;

                            // Get a reference to the shape for this row.
                            currentShape = Utilities.GetShapeItem(
                                currentPage.Shapes,
                                shapeIDs[stepID]);

                            // Make sure there is a shape to connect from
                            if (currentShape != null) {

                                // Connect each shape to the next steps.
                                for (int nextStep = 0;
                                    nextStep < nextStepCount;
                                    nextStep++) {

                                    nextShape = null;

                                    // Get a reference to the shape that
                                    // represents this next step.
                                    nextShape = Utilities.GetShapeItem(
                                        currentPage.Shapes,
                                        shapeIDs[Convert.ToString(
                                            nextStepID[nextStep], 
                                            CultureInfo.InvariantCulture)]);

                                    // Make sure there is a shape to connect to
                                    if (nextShape != null) {
                                        
                                        // Connect the two shapes.
                                        connectWithDynamicGlueAndConnector(
                                            currentShape, 
                                            nextShape);
                                    }
                                }
                            }
                        }
                    }
                    else {
                        // An error occurred while parsing the next
                        // steps. Do not continue processing.
                        isSuccessful = false;
                        break;
                    }

                    if (demoMode) {
                        // Allow Visio to repaint.
                        System.Windows.Forms.Application.DoEvents();
                    }

                    rowIndex++;
                    stepID = getExcelCellValue(excelSheet,
                        rowIndex, 
                        Utilities.ExcelColumnStepId);
                }
            }
            catch (COMException err) {
                isSuccessful = false;
                Utilities.DisplayException(visioApplication.AlertResponse, 
                    err.Message);
            }

            return isSuccessful;
        }

        /// <summary>This method drops a Dynamic connector master on
        /// the page, then connects the shapes by gluing the Dynamic 
        /// connector to the PinX values of the 2-D shapes to create 
        /// dynamic glue.</summary>
        /// <param name="shapeFrom">Shape where the Dynamic connector 
        /// begins</param>
        /// <param name="shapeTo">Shape where the Dynamic connector 
        /// ends</param>
        private static void connectWithDynamicGlueAndConnector(
            Shape shapeFrom, 
            Shape shapeTo) {

            Cell beginXCell;
            Cell endXCell;
            Shape connector;

            // Add a Dynamic connector to the page.
            connector = dropMasterOnPage(
                (Page)shapeFrom.ContainingPage, 
                "Dynamic Connector",
                Utilities.FlowchartStencil, 
                0.0, 
                0.0);

            // Connect the begin point.
            beginXCell = connector.get_CellsSRC(
                (short)VisSectionIndices.visSectionObject,
                (short)VisRowIndices.visRowXForm1D,
                (short)VisCellIndices.vis1DBeginX);

            beginXCell.GlueTo(shapeFrom.get_CellsSRC(
                (short)VisSectionIndices.visSectionObject,
                (short)VisRowIndices.visRowXFormOut,
                (short)VisCellIndices.visXFormPinX));

            // Connect the end point.
            endXCell = connector.get_CellsSRC(
                (short)VisSectionIndices.visSectionObject,
                (short)VisRowIndices.visRowXForm1D,
                (short)VisCellIndices.vis1DEndX);

            endXCell.GlueTo(shapeTo.get_CellsSRC(
                (short)VisSectionIndices.visSectionObject,
                (short)VisRowIndices.visRowXFormOut,
                (short)VisCellIndices.visXFormPinX));
        }

        /// <summary>This method reads the Visio SDK Sample Flowchart Excel
        /// workbook (flowchart.xlsx) and creates a Visio drawing based on 
        /// the data in the cells of each worksheet. The user is prompted 
        /// whether or not the method is to run in Demo mode. Demo mode 
        /// does not turn any functionality off; it simply allows the user 
        /// to watch the creation of the drawing as it is happening.
        /// </summary>
        /// <remark>Note: This method expects flowchart.xlsx to be located
        /// in the same directory as FlowchartSampleCSharp.dll.</remark>
        /// <param name="theApplication">Reference to the running Visio
        /// instance</param>
        /// <param name="documentIndex">Index of the newly created
        /// Visio Document</param>
//        [CLSCompliant(false)]
        public void CreateDrawing(
            Microsoft.Office.Interop.Visio.Application theApplication,
            int documentIndex) {

            if (theApplication == null) {
                return;
            }

            int sheetIndex;
            int scopeID = 0;
            bool newDrawing = true;
            bool success = true;
            string excelFile;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks = null;
            Microsoft.Office.Interop.Excel.Sheets excelSheets = null;
            int excelSheetCount;

            string infoText;
            Shape textInfo;
            Document currentDocument = null;
            Page currentPage;
                        
            Document backgroundStencil = null;
            Document titleStencil = null;

            try {
                visioApplication = theApplication;
                if (theApplication.Documents != null)
                    currentDocument = theApplication.Documents[documentIndex];

                // Prompt the user for the mode in which to run the
                // application. Demo mode will allow the user to watch
                // the drawing being created. Running the application
                // in Demo mode is slower, but may help developers to
                // understand how this sample works.
                //
                // Check AlertResponse to see whether we should display
                // modal UI. 
                if (theApplication.AlertResponse == 0) {

                    demoMode = (Utilities.RtlAwareMessageBoxShow(
                        Utilities.LoadString("DemonstrationMode"),
                        Utilities.LoadString("AddInName"), 
                        MessageBoxButtons.YesNo, 
                        MessageBoxIcon.None) == DialogResult.Yes);
                }
                else if (theApplication.AlertResponse ==
                    (short)DialogResult.No) {
                    demoMode = false;
                }
                else {
                    demoMode = true;
                }

                // In non-Demo mode, turn on DeferRecalc and turn off 
                // ShowChanges during drawing creation.
                //
                // Note: Most Visio solutions that create drawings should turn 
                // on DeferRecalc and turn off ShowChanges during the creation
                // process in order to improve performance.
                if (demoMode == false) {
                    theApplication.DeferRecalc = 1;
                    theApplication.ShowChanges = false;
                }

                excelApplication = 
                    new Microsoft.Office.Interop.Excel.Application();

                // Open the SDK Flowchart Sample Excel workbook.  Try first
                // with the path stored in the registry.
                excelFile = Utilities.FlowchartSamplePathFromRegistry + 
                    Utilities.ExcelFileName;

                //  Force macros in the Excel file to be disabled.  The shipping 
                //  Flowchart sample Excel file does not contain macros but other 
                //  data sources may. 
                Office.Core.MsoAutomationSecurity originalSecurity;
                bool resetSecurity = false;
                originalSecurity = excelApplication.AutomationSecurity;
                if (originalSecurity != Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable)
                {
                    excelApplication.AutomationSecurity =
                        Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                    resetSecurity = true;
                }                              

                excelWorkbooks = excelApplication.Workbooks;

                try 
                {
                        excelWorkbook = Utilities.GetSampleExcelWorkbook(
                        excelWorkbooks, excelFile);
                }
                catch (COMException) {
                    // The SDK Flowchart Sample Excel workbook failed
                    // to open.  Try again, but this time obtain the
                    // workbook's path directly from the add-in;
                    excelFile = Utilities.FlowchartSamplePathFromAddIn + 
                        Utilities.ExcelFileName;

                    excelWorkbook = Utilities.GetSampleExcelWorkbook(
                        excelWorkbooks, excelFile);

                }

                // Now that the document is open reset security so that
                // other Excel clients can open files with macros enabled. 
                if (resetSecurity)
                    excelApplication.AutomationSecurity = originalSecurity;

                
                // Begin the undo scope.
                scopeID = theApplication.BeginUndoScope(
                    Utilities.LoadString("UndoScopeName"));

                // Delete the text-only shape on the active page of the template before 
                // creating the drawing.
                textInfo = Utilities.GetShapeItem(
                    visioApplication.ActivePage.Shapes, 1);

                infoText = Utilities.LoadString("UserInfoText");
                if (textInfo.get_CellExistsU(infoText, 0) != 0) {

                    // The text-only shape can be deleted.
                    textInfo.Delete();
                }

                // Set the right menu action for certain masters.
                for (int index = 0; 
                     index < Utilities.StepShapeCount; 
                     index++) {
                    if (!setMasterRightMenuAction(theApplication, 
                        currentDocument,
                        Utilities.GetStepShapes(index))) {
                        success = false;
                        break;
                    }
                }

                if (success)
                {

                    // If the document has one page, then assume it's a new
                    // blank document and begin creating the drawing on the
                    // first page (index == 1).  If the document has multiple
                    // pages, then begin creating the drawing on a new page.
                    newDrawing = currentDocument.Pages.Count > 1 ? false : true;

                    // Create a Visio page for every Excel worksheet.
                    excelSheets = excelWorkbook.Worksheets;
                    excelSheetCount = excelSheets.Count;
                    for (sheetIndex = 1;
                        sheetIndex <= excelSheetCount;
                        sheetIndex++) {

                        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)
                            excelWorkbook.Worksheets[sheetIndex];

                        if (newDrawing) {
                            // A newly created document will already have one page.
                            currentPage = (Page)currentDocument.Pages[1];
                            newDrawing = false;
                        }
                        else {
                            currentPage = (Page)currentDocument.Pages.Add();
                        }

                        // Use the data in the Excel worksheet to create the 
                        // flowchart drawing for this page.
                        //
                        // Note: Failure creating a drawing on one page does
                        // not affect creating a drawing on another page.
                        drawPageFromExcelSheet(currentPage, excelSheet);

                        if (demoMode) {
                            // Allow Visio to repaint.
                            System.Windows.Forms.Application.DoEvents();
                        }
                        NullAndRelease(excelSheet);
                        excelSheet = null;
                    }

                    // Set the footer of the document.
                    currentDocument.FooterCenter = "&d &t";
                }

                // Commit the undo scope.
                theApplication.EndUndoScope(scopeID, true);
                scopeID = 0;
            }
            catch (COMException err) {

                // End the Undo scope if one was started.
                if (scopeID != 0) {
                    theApplication.EndUndoScope(scopeID, false);
                    scopeID = 0;
                }

                Utilities.DisplayException(visioApplication.AlertResponse, 
                    err.Message);
            }
            finally {
                // In non-Demo mode, after completing drawing creation,
                // reset ShowChanges and DeferRecalc. This action
                // will return Visio to its interactive mode and
                // allow Visio to display the newly created drawings.
                if (demoMode == false) {
                    theApplication.ShowChanges = true;
                    theApplication.DeferRecalc = 0;
                }

                if (excelApplication != null) {

                    // Release all references to Excel Objects and quit Excel.
                    NullAndRelease(excelSheet);
                    NullAndRelease(excelSheets);
                    NullAndRelease(excelWorkbooks);


                    if (excelWorkbook != null) {
                        excelWorkbook.Close(false, "", false);
                        NullAndRelease(excelWorkbook);
                        excelWorkbook = null;
                    }

                    excelApplication.Quit();
                    NullAndRelease(excelApplication);
                
                }

                // close background and title stencils
                if (backgroundStencil != null)
                    backgroundStencil.Close();
                if (titleStencil != null)
                    titleStencil.Close();

                // Explicitly set Visio objects to Nothing 
                currentPage = null;
                currentDocument = null;
            }
        }

        /// <summary>This method uses the data in the Excel worksheet to 
        /// create the flowchart drawing on the Visio page.</summary>
        /// <param name="currentPage">Reference to the Visio page</param>
        /// <param name="excelSheet">Worksheet, in the source Excel 
        /// file, which is used to determine the shapes, shape properties and 
        /// connections in the resulting flowchart</param>
        /// <returns>True if successful; False otherwise</returns>        
        private bool drawPageFromExcelSheet(
            Page currentPage, 
            Microsoft.Office.Interop.Excel.Worksheet excelSheet) {

            Cell layoutCell;
            bool returnValue = true;

            try {

                // Try to set the name of the Visio Page to the same name as 
                // the worksheet, excelSheet. Failure to set the name does 
                // not stop the drawing creation and does not yield an error 
                // message.
                setPageName(currentPage, excelSheet.Name);

                // Change the page settings to use rounded connector lines.
                layoutCell = currentPage.PageSheet.get_CellsSRC(
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowPageLayout,
                    (short)VisCellIndices.visPLOLineRouteExt);

                layoutCell.set_Result(VisUnitCodes.visPageUnits,
                    (double)VisCellVals.visLORouteExtNURBS);

                if (dropShapesFromWorksheet(currentPage, excelSheet) == true) {
                    // Only try to connect shapes if dropping shapes
                    // was successful.
                    connectSteps(currentPage, excelSheet);
                }

                // Automatically layout the shapes as a flowchart / tree view.
                layoutCell = currentPage.PageSheet.get_CellsSRC(
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowPageLayout,
                    (short)VisCellIndices.visPLOPlaceStyle);

                layoutCell.set_Result(VisUnitCodes.visPageUnits,
                    (double)VisCellVals.visPLOPlaceTopToBottom);
                currentPage.Layout();
                currentPage.CenterDrawing();
            }
            catch (COMException err) {
                returnValue = false;
                Utilities.DisplayException(visioApplication.AlertResponse, 
                    err.Message);
            }

            return returnValue;
        }

        /// <summary>This method looks for the document with the name
        /// stencilName in the Documents collection, and if the document is 
        /// not open, opens it as a docked stencil. It then gets the master 
        /// within that stencil, using its univeral name and drops it on the 
        /// specified page.</summary>
        /// <param name="dropOnPage">Page where the master will be dropped
        /// </param>
        /// <param name="masterNameU">Universal name of the master in the
        /// stencil</param>
        /// <param name="stencilName">Name of the stencil where the master 
        /// is to be found</param>
        /// <param name="pinX">X-coordinate of the Pin in internal units
        /// </param>
        /// <param name="pinY">Y-coordinate of the Pin in internal units
        /// </param>
        /// <returns>Shape that was dropped on the page</returns>
        private static Shape dropMasterOnPage(
            Page dropOnPage, 
            string masterNameU,
            string stencilName, 
            double pinX, 
            double pinY) {

            Documents applicationDocuments;
            Document stencil;
            Master masterToDrop;
            Shape returnShape = null;

            // Find the stencil in the Documents collection by name.
            applicationDocuments = dropOnPage.Application.Documents;
            stencil = Utilities.GetStencil(applicationDocuments, 
                stencilName);

            // Get a master on the stencil by its universal name.
            masterToDrop = Utilities.GetMasterItem(stencil, masterNameU);

            // Drop the master on the page that is passed in. Set the
            // PinX and PinY using the data passed in parameters pinX
            // and pinY, respectively.
            returnShape = dropOnPage.Drop(masterToDrop, pinX, pinY);

            return returnShape;
        }

        /// <summary>This method drops shapes onto the Visio page based on 
        /// data in the cells of the Excel Worksheet.</summary>
        /// <param name="currentPage">Page on which to drop shapes</param>
        /// <param name="excelSheet">Worksheet that has information about what
        /// shapes to create</param>
        /// <returns>True if successful; False otherwise</returns>
        private bool dropShapesFromWorksheet(
            Page currentPage,
            Microsoft.Office.Interop.Excel.Worksheet excelSheet) {

            bool isSuccessful = true;
            int rowIndex;
            double pinX;
            double pinY;
            Cell multiUseCell;
            string stepID;
            string stepType;
            string owner;
            string colorFormulaU;
            string durationProperty;
            string resourceProperty;
            System.Text.StringBuilder duplicateErrorMessage = new System.Text.StringBuilder();
            double duration = 0.0;
            Shape dropShape;
            Hyperlink shapeHyperlink;
            string hyperlinkString;

            try {

                shapeIDs = new SortedList();

                // Add shapes.
                rowIndex = Utilities.ExcelRowStart;
                stepID = getExcelCellValue(excelSheet,
                    rowIndex, 
                    Utilities.ExcelColumnStepId);

                while (stepID.Length != 0) {

                    // Calculate a random location for the shape. Randomizing
                    // the placement of shapes helps lay out shapes more
                    // efficiently when the Layout method is called for
                    // this page.
                    Random randomLocation = new Random();
                    pinX = randomLocation.NextDouble() * 10;
                    pinY = randomLocation.NextDouble() * 10;

                    stepType = getExcelCellValue(excelSheet,
                        rowIndex, 
                        Utilities.ExcelColumnStepType);

                    dropShape = dropMasterOnPage(currentPage,
                        getMasterNameFromStepType(stepType),
                        Utilities.FlowchartStencil, 
                        pinX, 
                        pinY);

                    // Store the ID of the shape and key it by the Step ID
                    // so that the shapes can be retrieved later.
                    if (!shapeIDs.Contains(stepID)) {
                        shapeIDs.Add(stepID, dropShape.ID);
                    }
                    else {
                        // Create error message
                        duplicateErrorMessage.AppendFormat(
                                        System.Globalization.CultureInfo.CurrentUICulture,
                                        "{0}\n{1}{2}\n{3}{4}\n", 
                                        Utilities.LoadString("ErrorDuplicatedID"),
                                        Utilities.LoadString("ErrorDuplicatedKey"),
                                        stepID,
                                        Utilities.LoadString("ErrorDuplicatedPage"),
                                        currentPage.Name);

                        // Warn user about duplicate ID.
                        if (visioApplication.AlertResponse == 0) {
                            Utilities.RtlAwareMessageBoxShow(
                                duplicateErrorMessage.ToString(),
                                Utilities.LoadString("AddInName"),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else
                            System.Diagnostics.Debug.WriteLine(
                                duplicateErrorMessage);
                    }

                    // Set the text of the shape.
                    dropShape.Text = getExcelCellValue(
                        excelSheet, 
                        rowIndex, 
                        Utilities.ExcelColumnTask);

                    // Set the color of the shape based on the owner.
                    owner = getExcelCellValue(
                        excelSheet, 
                        rowIndex, 
                        Utilities.ExcelColumnOwner);
                    colorFormulaU = getShapeColorFromOwner(owner);

                    if (colorFormulaU.Length > 0) {
                        multiUseCell = dropShape.get_CellsSRC(
                            (short)VisSectionIndices.visSectionObject,
                            (short)VisRowIndices.visRowFill,
                            (short)VisCellIndices.visFillForegnd);
                        multiUseCell.FormulaU = colorFormulaU;
                    }

                    // Add the hyperlink if its address is not empty
                    hyperlinkString = getExcelCellValue(
                        excelSheet, 
                        rowIndex, 
                        Utilities.ExcelColumnHyperlink);

                    if (hyperlinkString.Length != 0) {

                        shapeHyperlink = dropShape.AddHyperlink();
                        shapeHyperlink.Address = hyperlinkString;

                        // Add the Hyperlink description if the Name column is present
                        shapeHyperlink.Description = getExcelCellValue(
                            excelSheet, 
                            rowIndex, 
                            Utilities.ExcelColumnHyperlinkDescription);
                    }

                    // Add the Shape ScreenTip.
                    multiUseCell = dropShape.get_CellsSRC(
                        (short)VisSectionIndices.visSectionObject,
                        (short)VisRowIndices.visRowMisc,
                        (short)VisCellIndices.visComment);
                    multiUseCell.FormulaU = Utilities.StringToFormulaForString(
                        getExcelCellValue(excelSheet, 
                            rowIndex,
                            Utilities.ExcelColumnComments));

                    // Set custom property (shape data) values 

                    // Get the name for the duration property row.
                    durationProperty = 
                        Utilities.LoadString("CustomPropertyDuration");

                    // Add the duration if the shape has a duration property.
                    if (dropShape.get_CellExistsU(durationProperty, 0) != 0) {
                        multiUseCell = dropShape.get_CellsU(durationProperty);

                        try {
                            duration = Convert.ToDouble(getExcelCellValue(
                                excelSheet, 
                                rowIndex, 
                                Utilities.ExcelColumnDuration),
                                CultureInfo.InvariantCulture);
                        }
                        catch (InvalidCastException) {
                            // If an invalid value is in the workbook's duration
                            // cell, then an exception may be thrown.  In this case,
                            // the exception should be caught and ignored so the
                            // rest of the drawing can be created.
                        }

                        multiUseCell.set_Result(
                            VisUnitCodes.visNumber, 
                            duration);
                    }

                    // Get the name for the resource property.
                    resourceProperty =
                        Utilities.LoadString("CustomPropertyResources");

                    // Add the resource if the shape has a resource property.
                    if (dropShape.get_CellExistsU(resourceProperty, 0) != 0) {
                        multiUseCell = dropShape.get_CellsU(resourceProperty);
                        multiUseCell.FormulaU =    
                            Utilities.StringToFormulaForString(owner);
                    }

                    if (demoMode) {
                        // Allow Visio to repaint.
                        System.Windows.Forms.Application.DoEvents();
                    }

                    rowIndex++;
                    stepID = getExcelCellValue(excelSheet,
                        rowIndex, 
                        Utilities.ExcelColumnStepId);
                }
            }
            catch (COMException err) {
                isSuccessful = false;
                Utilities.DisplayException(visioApplication.AlertResponse, 
                    err.Message);
            }

            return isSuccessful;
        }


		/// <summary>This method uses row and column information to return 
		/// the value of a cell in the Excel worksheet that is passed in.
		/// </summary>
		/// <param name="excelSheet">Worksheet that has value to be 
		/// retrieved</param>
		/// <param name="rowIndex">Row index of the cell</param>
		/// <param name="columnIndex">Column index of the cell</param>
		/// <returns>Value of cell as a string, if a value exists;
		/// otherwise an empty string</returns>
		private static string getExcelCellValue(
			Microsoft.Office.Interop.Excel.Worksheet excelSheet, 
			int rowIndex, 
			int columnIndex) {

            string returnString = "";
            Microsoft.Office.Interop.Excel.Range excelCell;
            Microsoft.Office.Interop.Excel.Range excelCells;
            object cellValue = null;

            excelCells = (Microsoft.Office.Interop.Excel.Range)
                excelSheet.Cells;
            excelCell =  (Microsoft.Office.Interop.Excel.Range)
                (excelCells[rowIndex, columnIndex]);

            cellValue = excelCell.get_Value(
                Microsoft.Office.Interop.Excel.XlRangeValueDataType.
                xlRangeValueDefault);

            if (cellValue != null)
                returnString = cellValue.ToString();
            
            NullAndRelease(excelCell);
            NullAndRelease(excelCells);
            
            return returnString;
        }

        /// <summary>This method returns the master name based on the 
        /// step type.</summary>
        /// <param name="stepType">Data from Step Type column in Excel 
        /// worksheet</param>
        /// <returns>Universal name of the master associated with this
        /// step type. The universal name of process master will be 
        /// returned if it cannot recognize the step type</returns>
        private static string getMasterNameFromStepType(string stepType) {

            // Treat unknown step types as process shapes.
            string masterNameU = Utilities.GetStepShapes(0);

            for (int index = 0; 
                 index < Utilities.StepShapeCount; 
                 index++) {
                if (Utilities.GetStepTypes(index).Equals(stepType)) {

                    masterNameU = Utilities.GetStepShapes(index);
                    break;
                }
            }
            return masterNameU;
        }

        /// <summary>This method returns the RGB formula for the shape fill 
        /// color, based on the name of the shape's owner.</summary>
        /// <param name="owner">Owner's name</param>
        /// <returns>Constant that represents the universal formula of the 
        /// appropriate fill color or a blank string if no color for the
        /// specified owner is defined</returns>
        private static string getShapeColorFromOwner(string owner) {

            string shapeColor = "";

            for (int index = 0; 
                 index < Utilities.OwnerColorCount; 
                 index++) {
                
                            if (string.Compare(owner, Utilities.GetOwnerColor(index, OwnerColorField.OwnerColorOwner),
                            true, CultureInfo.CurrentUICulture) == 0) {
                    shapeColor = Utilities.GetOwnerColor(index, 
                        OwnerColorField.OwnerColorColor);
                    break;
                }
            }

            return shapeColor;
        }

        /// <summary>This method parses the comma delimited Step IDs and
        /// returns them in an array.</summary>
        /// <param name="nextSteps">Comma delimited Step IDs</param>
        /// <param name="nextStep">Array to hold the Step IDs</param>
        /// <param name="stepCount">Number of items in the returned array
        /// </param>
        /// <returns>True if successful; False otherwise</returns>
        private bool parseNextSteps(
            string nextSteps,
            ref int[] nextStep, 
            ref int stepCount) {

            int nextComma;
            int    position = 0;
            bool populate = true;
            bool isSuccessful = true;

            try {

                nextSteps = nextSteps.Trim();
                stepCount = 0;

                if (nextSteps.Length > 0) {

                    // Populate the array with Step IDs.
                    while (populate) {

                        nextComma = nextSteps.IndexOf(",", position, StringComparison.Ordinal);

                        if (nextComma > 0) {
                            nextStep[stepCount] = Convert.ToInt16(
                                nextSteps.Substring(position,
                                    nextComma - position), 
                                CultureInfo.InvariantCulture);
                            position = nextComma + 1;
                        }
                        else {
                            nextStep[stepCount] = Convert.ToInt16(
                                nextSteps.Substring(position),
                                CultureInfo.InvariantCulture);

                            // The last Step ID has been reached.
                            populate = false;
                        }

                        stepCount++;
                    }
                }
            }
            catch (InvalidCastException err) {
                isSuccessful = false;
                Utilities.DisplayException(visioApplication.AlertResponse, 
                    err.Message);
            }
            catch (FormatException err) {
                isSuccessful = false;
                Utilities.DisplayException(visioApplication.AlertResponse, 
                    err.Message);
            }

            return isSuccessful;
        }


	/// <summary>This method creates an action row for a master.
        /// </summary>
        /// <remark>Note: The right menu action tells the document to call 
        /// the SheetQueueMarker macro.</remark>
        /// <param name="theApplication">Reference to the running Visio
        /// instance</param>
        /// <param name="currentDocument">Reference to the document which
        /// will contain the drawing</param>
        /// <param name="masterU">Master to which the right menu action will
        /// be added</param>
        private static bool setMasterRightMenuAction(
            Microsoft.Office.Interop.Visio.Application theApplication,
            Document currentDocument, 
            string masterU) {

            int actionRow;
            int actionRows;
            Cell actionCell;
            Documents applicationDocuments;
            Document stencil;
            Master stencilMaster;
            Master documentMaster;
            Master masterEditCopy;
            string menuCaption;
            bool mustDropMaster;
            bool mustModifyMaster;
            Shape masterShape;

            try {

                //Assume master has not been added to the document yet
                // and must both be dropped and modified.
                mustModifyMaster = true;
                mustDropMaster = true;

                // Set our caption string.
                menuCaption = Utilities.StringToFormulaForString(
                    Utilities.LoadString("RmaMenuCaption"));

                // Determine if the master really needs to be dropped.
                documentMaster = Utilities.GetMasterItem(currentDocument, 
                    masterU);

                // If it exists, then determine if it needs to be changed.
                // Get a reference to the title master.
                if (documentMaster != null) {

                    mustDropMaster = false;
                    masterShape = documentMaster.Shapes[1];

                    // Check if the right menu action item already exists.
                    actionRows = masterShape.get_RowCount(
                        (short)VisSectionIndices.visSectionAction);

                    for (actionRow = 1; 
                        actionRow <= actionRows && mustModifyMaster; 
                        actionRow++) {

                        actionCell = masterShape.get_CellsSRC(
                            (short)VisSectionIndices.visSectionAction,
                            (short)((int)VisRowIndices.visRowAction+actionRow),
                            (short)VisCellIndices.visActionMenu);

                        if (actionCell.FormulaU == menuCaption) {

                            // The right menu action item already exists;
                            // do not add it again.
                            mustModifyMaster = false;
                        }
                    }
                }

                if (mustDropMaster) {

                    // Find the Basic Flowchart Shapes stencil.
                    applicationDocuments = theApplication.Documents;
                    stencil = Utilities.GetStencil(applicationDocuments,
                        Utilities.FlowchartStencil);

                    if (stencil != null)  {
                        // Get the master in the stencil.
                        stencilMaster = Utilities.GetMasterItem(stencil, masterU);

                        // Drop the master from the stencil onto the document.    
                        documentMaster = currentDocument.Drop(stencilMaster, 0, 0);
                    }
                    else
                        return false;
                }

                if ((documentMaster != null) && mustModifyMaster) {

                    // Open the master for editing.  Open returns a copy of
                    // this master for editing.  Close will merge the changes
                    // into the master.
                    masterEditCopy = documentMaster.Open();

                    // Add a menu item to the right menu action for
                    // shapes created from this master.
                    masterShape = masterEditCopy.Shapes[1];

                    // The right menu action item does not yet exist,
                    // so add it.
                    actionRow = masterShape.AddRow(
                        (short)VisSectionIndices.visSectionAction,
                        (short)VisRowIndices.visRowLast,
                        (short)VisRowIndices.visRowAction);

                    // Set the menu caption.
                    actionCell = masterShape.get_CellsSRC(
                        (short)VisSectionIndices.visSectionAction,
                        (short)actionRow,
                        (short)VisCellIndices.visActionMenu);
                    actionCell.FormulaU = menuCaption;

                    // Set the action for the menu item.
                    actionCell = masterShape.get_CellsSRC(
                        (short)VisSectionIndices.visSectionAction,
                        (short)actionRow,
                        (short)VisCellIndices.visActionAction);
                    actionCell.FormulaU = Utilities.RmaMenuFormula;

                    // Release all objects related to this copy of the master
                    // before calling Close which will delete this copy.
                    // Failure to do this can cause an exception to be thrown.
                    actionCell = null;
                    masterShape = null;
                    masterEditCopy.Close();

                    // Tell Visio to match by name on drop. 
                    // Editing the master changes its unique ID. So
                    // by telling Visio to match by name on drop, the 
                    // master can be dropped from the stencil onto any page
                    // in this document, and Visio will create a shape that 
                    // is an instance of the document's modified master 
                    // rather than an instance of the stencil's master.
                    documentMaster.MatchByName = 1;                    
                }
            }
            catch (COMException err) {
                Utilities.DisplayException(theApplication.AlertResponse, 
                    err.Message);
            }
            return true;
        }

        /// <summary>This method sets the name of the Visio page to the new 
        /// page name passed in, if possible. If it does not succeed, the 
        /// Visio page will keep its default name. </summary>
        /// <remark>Note: Visio has its own definition for "sheet" and  
        /// often will not accept the default sheet name from an Excel 
        /// spreadsheet.</remark>
        /// <param name="currentPage">Page to be renamed</param>
        /// <param name="newPageName">New name for the page</param>
        private static void setPageName(
            Page currentPage, 
            string newPageName) {

            try {
                currentPage.Name = newPageName;
            }
            catch (COMException) {
                // Consume error
            }
        }

        /// <summary>This method releases all references to a COM object. When
        /// Visual Studio .NET calls a COM object from managed code, it
        /// automatically creates a Runtime Callable Wrapper (RCW). The RCW
        /// marshals calls between the .NET application and the COM object. The
        /// RCW keeps a reference count on the COM object. Calling
        /// ReleaseComObject when you are finished using an object will cause
        /// the reference count of the RCW to be decremented.</summary>
        /// <param name="runtimeObject">The runtime callable wrapper whose
        /// underlying COM object will be released</param>
        [SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")]
        private static void NullAndRelease(object runtimeObject)        {
            
            try {

                if (runtimeObject != null && 
                    System.Runtime.InteropServices.Marshal.IsComObject(runtimeObject))     {

                    // The RCW's reference count gets incremented each time the
                    // COM pointer is passed from unmanaged to managed code.
                    // Call ReleaseComObject in a loop until it returns 0 to be
                    // sure that the underlying COM object gets released.
                    int referenceCount = System.Runtime.InteropServices.
                        Marshal.ReleaseComObject(runtimeObject);

                    while (0 < referenceCount) {
                        referenceCount =
                         System.Runtime.InteropServices.Marshal.ReleaseComObject(runtimeObject);
                    }
                }
            }
            finally {
                runtimeObject = null;
            }
        }
    }
}
