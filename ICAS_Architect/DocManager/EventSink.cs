// EventSink.cs
// compile with: /doc:EventSink.xml
// <copyright>Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <summary>This file contains the EventSink class.</summary>

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;

namespace Microsoft.Samples.Visio.FlowchartAddIn.CSharp {

    /// <summary>This class is an event sink for Visio events. In other  
    /// words, it handles events from Visio which were specified using 
    /// AddAdvise. It handles event notification by implementing the 
    /// IVisEventProc interface, which is defined in the Visio type library.
    /// </summary>
    /// <remark>
    /// In order to be notified of events, an instance of this class must be
    /// passed as the EventSink argument in calls to AddAdvise. The Flowchart
    /// Sample calls AddAdvise once in Visio on the Connect module where it
    /// registers for Marker events sourced by the Application object.  The
    /// code in this class processes marker events created by this Flowchart
    /// sample.</remark>
    [ProgId("FlowchartSampleCSharp.EventSink")]
    [ComVisible(true)]
    public class EventSink
        : IVisEventProc {

        /// <summary>This constructor creates an instance of the 
        /// CustomCommandBar class, since it will be used to create the  
        /// command bar in this class.</summary>
        public EventSink() {}

        /// <summary>This method is called when the DocumentCreate event is
        /// triggered by a user creating or copying a document based on the
        /// Visio SDK Flowchart Sample template, Flowchart (CSharp).vstx. When
        /// a new document is created, this method creates a drawing and a 
        /// custom command bar. If the first foreground page of a document 
        /// contains shapes, this method does not create a drawing.</summary>
        /// <remark>
        /// Note: The DocumentCreate event is also triggered when Visio 
        /// creates a read only copy of a drawing that was originally created 
        /// based on our template. This can happen when a user tries to open 
        /// an restricted file. For example, a user can direct Visio to open 
        /// a read only copy of a file if the user tries to open a file that is
        /// being edited by another user.  In the special case where the 
        /// drawing already exists and is just being copied, do not create a 
        /// drawing. However, a CommandBar needs to be created.</remark>
        /// <param name="visioApplication">Reference to the running Visio
        /// instance</param>
        /// <param name="documentIndex">Index of the document that was 
        /// created</param>
        private static void handleDocumentCreate(
            Microsoft.Office.Interop.Visio.Application visioApplication,
            int documentIndex) {
            string infoText;
            Shape textInfo;
            DocumentCreator visioDocumentCreator = null;
            
            // When a user creates a new document based on the
            // Visio SDK Flowchart Sample template,
            // create a drawing if the only shape on the page
            // is the text-only shape instructing the user about
            // how the template is used.
            if (visioApplication.ActivePage.Shapes.Count == 1) 
            {

                textInfo = Utilities.GetShapeItem(
                    visioApplication.ActivePage.Shapes, 1);
                
                infoText = Utilities.LoadString("UserInfoText");
                if (textInfo.get_CellExistsU(infoText, 0) != 0) {
                    // Create the drawing.
                    visioDocumentCreator = new DocumentCreator();
                    visioDocumentCreator.CreateDrawing(
                        visioApplication, 
                        documentIndex);
                }
            }
        }

        /// <summary>This method is called when the user opens a document 
        /// that was created with the Visio SDK Flowchart Sample template, 
        /// Flowchart (CSharp).vstx, or when the user opens the template for  
        /// editing. In both cases it creates a custom CommandBar. If the 
        /// extension of the document opened indicates that the user opened 
        /// the template for editing, a messagebox will appear that directs 
        /// the user to create a document based on the template in order to 
        /// see the drawing creation functionality.</summary>
        /// <param name="visioApplication">Reference to the running Visio
        /// instance</param>
        /// <param name="documentIndex">Index of the document that was 
        /// opened</param>
        private static void handleDocumentOpen(
            Microsoft.Office.Interop.Visio.Application visioApplication,
            int documentIndex) {

            bool binaryTemplate;
            bool XMLTemplate;
            Document currentDocument;
            string documentName;

            // See if the user is editing the template by getting
            // the document and checking its extension. The document
            // is a template if its extension is .vst, .vstx, .vstm or.vtx.

            // First get the document that was opened
            // using the documentIndex parameter.
            currentDocument = visioApplication.Documents[documentIndex];

            // Next get the document name and check its extension.  
            // If the document has been saved, the document
            // name is its file name. If the document is a template,
            // its extension is .vst, .vstx, .vstm or .vtx.
            documentName = currentDocument.Name;

            // This sample ships with the VSTX template but
            // it is possible to create other types of templates so
            // check for all types of templates.
            binaryTemplate = documentName.EndsWith("." + Utilities.TemplateExtensionVst, StringComparison.OrdinalIgnoreCase);

            XMLTemplate = documentName.EndsWith("." + Utilities.TemplateExtensionVtx, StringComparison.OrdinalIgnoreCase);

            bool vstxTemplate = documentName.EndsWith("." + Utilities.TemplateExtensionVstx, StringComparison.OrdinalIgnoreCase);

            bool vstmTemplate = documentName.EndsWith("." + Utilities.TemplateExtensionVstm, StringComparison.OrdinalIgnoreCase);

            // If the user opens the template for editing, show a
            // message box which explains how to access the
            // drawing creation feature of this sample.
            if (vstxTemplate || vstmTemplate || binaryTemplate || XMLTemplate) {
                Utilities.DisplayException(
                    visioApplication.AlertResponse,
                    Utilities.LoadString("TemplateOpen"));
            }
        }

        /// <summary>This method is called when the user clicks on the
        /// right menu action of a 2-D shape added to the flowchart shapes by
        /// this add-in. It displays the number of connections and the names
        /// of the 1-D shapes that are connected to it.</summary>
        /// <param name="visioApplication">Reference to the running Visio
        /// instance</param>
        /// <param name="documentIndex">Index of the document that contains
        /// the shape</param>
        /// <param name="pageIndex">Index of the page that contains the shape
        /// </param>
        /// <param name="shapeNameID">Unique identifier which identifies the
        /// 2-D shape on the page</param>
        private static void handleRightMouseClick(
            Microsoft.Office.Interop.Visio.Application visioApplication,
            int documentIndex, 
            int pageIndex, 
            string shapeNameID) {

            Connects shapeConnects;
            Document currentDocument;
            Page currentPage;
            Shape currentShape;
            System.Text.StringBuilder message = new System.Text.StringBuilder();

            // Get the connections to the 2-D shape.
            currentDocument = visioApplication.Documents[documentIndex];
            currentPage = (Page)currentDocument.Pages[pageIndex];
            currentShape = currentPage.Shapes[shapeNameID];
            shapeConnects = currentShape.FromConnects;

            // Show how many connections are made to the 2-D shape and
            // the names of the 1-D shapes that are connected to it.
            message.AppendFormat(CultureInfo.CurrentCulture,
                Utilities.LoadString(
                "DisplayNumberOfConnections"),
                shapeConnects.Count);
            message.Append("\n");

            foreach (Microsoft.Office.Interop.Visio.Connect connect 
                            in shapeConnects) {
                message.Append(connect.FromSheet.Name);
                message.Append("\n");
            }

            // Check AlertResponse to see whether we should display
            // modal UI. 
            if (visioApplication.AlertResponse == 0) {
                Utilities.RtlAwareMessageBoxShow(
                    message.ToString(),
                    Utilities.LoadString("AddInName"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        /// <summary>This method is called by Visio when an event, which has
        /// been added to the EventList collection, has been triggered.
        /// This COM add-in calls the AddAdvise method on the EventList
        /// collection for the Application object once in the Connect class 
        /// to register for marker events. As a result, this method will be
        /// called whenever a MarkerEvent event occurs in Visio, but will  
        /// only respond to marker events from documents based on our 
        /// template.</summary>
        /// <param name="nEventCode">Event code of the event that fired
        /// </param>
        /// <param name="pSourceObj">Reference to object having an EventList,
        /// which contains the Event object that fired</param>
        /// <param name="nEventID">Unique identifier of the Event object
        /// that fired</param>
        /// <param name="nEventSeqNum">The sequence of the event 
        /// relative to events that have fired so far in the instance of 
        /// Visio</param>
        /// <param name="pSubjectObj">Reference to the subject of the Event,
        /// which is the object to which the event occurred</param>
        /// <param name="vMoreInfo">Additional information about the event
        /// </param>
        /// <returns>Empty object: The Visio Application object ignores the 
        /// return value unless the event is a query event. Since this 
        /// add-in does not process query events, this method returns an 
        /// empty variant.</returns>
        public object VisEventProc(
            short nEventCode,
            object pSourceObj,
            int nEventID,
            int nEventSeqNum,
            object pSubjectObj,
            object vMoreInfo) {

            Microsoft.Office.Interop.Visio.Application visioApplication = null;
            int commandID = 0;
            int documentIndex = 0;
            int pageIndex = 0;
            string context = null;
            string shapeNameID = "";

            try {

                // Only respond to Marker events.
                if (nEventCode ==
                    ((short)VisEventCodes.visEvtApp +
                    (short)VisEventCodes.visEvtMarker)) {

                    visioApplication =
                        (Microsoft.Office.Interop.Visio.Application) pSourceObj;

                    // Check if this Marker event was fired by a document
                    // created with the Visio SDK Flowchart Sample template.
                    if (visioApplication != null)
                        context = visioApplication.get_EventInfo(
                            (int)VisEventCodes.visEvtIdMostRecent);

                    if ((context != null) &&
                        parseEventContext(visioApplication, 
                            ref context,
                            ref commandID, 
                            ref documentIndex,
                            ref pageIndex, 
                            ref shapeNameID)) {

                        // Determine which command was called and handle it.
                        switch (commandID) {
                            case Utilities.DocumentCreateCommandId:
                                handleDocumentCreate(
                                    visioApplication, 
                                    documentIndex);
                                break;

                            case Utilities.DocumentOpenCommandId:
                                handleDocumentOpen(
                                    visioApplication, 
                                    documentIndex);
                                break;

                            case Utilities.ShapeRmaConnectsCommandId:
                                handleRightMouseClick(visioApplication,
                                    documentIndex, 
                                    pageIndex, 
                                    shapeNameID);
                                break;

                            default:
                                break;
                        }
                    }
                }
            }
            catch (COMException err) {
                Utilities.DisplayException(visioApplication.AlertResponse, 
                    err.Message);
            }

            return null;
        }

        /// <summary>This method extracts specific information requested by 
        /// the ParseEventContext method, as indicated via the parse 
        /// parameter.</summary>
        /// <param name="visioApplication">Reference to the Visio 
        /// Application object</param>
        /// <param name="context">Context string to be parsed</param>
        /// <param name="parse">Contains one of the following values:
        ///        ContextEvent       Parses the Event ID
        ///        ContextDocument    Parses the Document Index
        ///        ContextPage        Parses the Page
        ///        ContextShape       Parses the Shape Name ID</param>
        /// <param name="returnValue">Holds the requested value</param>
        /// <param name="showError">Determines if an error message should 
        /// be shown if the format of the string is incorrect</param>
        /// <returns>True if successful; False otherwise</returns>
        private static bool parseContextString(
            Microsoft.Office.Interop.Visio.Application visioApplication,
            string context, 
            string parse, 
            ref string returnValue,
            bool showError) {

            bool isSuccessful = true;
            int startPosition;
            int endPosition;
            string contextUpper = context.ToUpper(CultureInfo.InvariantCulture);

            // Find the position in the context string where
            // the parse string begins. For example, if the 
            // context string is "/event=2 /Doc=1 /solution=SDKFLW_VBNET",
            // and the parse string is "event=" then start position
            // is the location of 'e' is the context string.
            startPosition = contextUpper.IndexOf(parse, StringComparison.Ordinal);

            if (startPosition > 0) {

                // Skip over the parse string. For example, if
                // the parse string is "event=" then the startPosition
                // is the location of the character after '=' in the
                // context string.
                startPosition += parse.Length;

                // The endPosition is the location of the next 
                // ContextBeginMarker following the parse string.
                // If the parse string is the last command in the
                // context string then endPostion will be 0.
                endPosition = contextUpper.IndexOf(
                        Utilities.ContextBeginMarker,
                        startPosition, StringComparison.Ordinal);

                // Return the value of the element in the string.
                // So in the above example where the parse string
                // is "event=" then the string "2" will be returned.
                if (endPosition > 0) {
                    returnValue = context.Substring(startPosition,
                        endPosition - startPosition);
                }
                else {
                    returnValue = context.Substring(startPosition);
                }

                returnValue = returnValue.TrimEnd();
            }
            else if (showError) {
                Utilities.DisplayException(visioApplication.AlertResponse,
                    Utilities.LoadString("InvalidSyntax"));
                isSuccessful = false;
            }
        
            return isSuccessful;
        }

        /// <summary>This method parses the context string and extracts
        /// the command ID, document index and, in the case of the right 
        /// menu action command, the page number and the shape name ID.
        /// </summary>
        /// <remark>The following is the format of a context string: 
        /// "/event=[Command ID] /doc=[Document Index] /page=[Page]
        ///      /shape=[Shape Name ID] /solution=SDKFLW_C#NET"</remark>
        /// <param name="visioApplication">Reference to the running Visio
        /// instance</param>
        /// <param name="context">Context string that holds the data to be 
        /// parsed</param>
        /// <param name="commandID">Output variable that will hold the 
        /// command ID</param>
        /// <param name="documentIndex">Output variable that will hold the 
        /// document index</param>
        /// <param name="pageIndex">Output variable that will hold the
        /// page number</param>
        /// <param name="shapeNameID">Output variable that will hold the 
        /// shape name ID</param>
        /// <returns>True if the context string follows the correct syntax;
        /// False otherwise</returns>returns>
        private static bool parseEventContext(
            Microsoft.Office.Interop.Visio.Application visioApplication,
            ref string context, 
            ref int commandID, 
            ref int documentIndex,
            ref int pageIndex, 
            ref string  shapeNameID) {

            string parsedValue = "";

            // Check if this context string was created
            // by the Visio SDK Flowchart Sample Template.
            if (parseContextString(visioApplication, context,
                Utilities.ContextSolution, ref parsedValue, false)) {

                if (parsedValue.ToUpper(CultureInfo.InvariantCulture).
                        IndexOf(Utilities.ContextSdkFlowchart, StringComparison.Ordinal) == 0)
                {

                    // Get the Command ID.
                    if (!parseContextString(visioApplication, 
                        context,
                        Utilities.ContextEvent, 
                        ref parsedValue, 
                        true)) {

                        return false;
                    }
                    else {
                        commandID = Convert.ToInt32(parsedValue,
                            CultureInfo.InvariantCulture);
                    }

                    // Get the document Index.
                    if (parseContextString(visioApplication, 
                        context,
                        Utilities.ContextDocument, 
                        ref parsedValue, 
                        true)) {

                        documentIndex = Convert.ToInt32(parsedValue,
                            CultureInfo.InvariantCulture);
                    }

                    // If the command ID corresponds with the right menu
                    // action command, the page and the shape name ID are
                    // needed.
                    if (commandID == Utilities.ShapeRmaConnectsCommandId) {

                        // Get the page index.
                        if (parseContextString(visioApplication, 
                            context,
                            Utilities.ContextPage, 
                            ref parsedValue, 
                            true)) {

                            pageIndex = Convert.ToInt32(parsedValue,
                                CultureInfo.InvariantCulture);
                        }

                        // Get the shape name ID.
                        if (parseContextString(visioApplication, 
                            context,
                            Utilities.ContextShape, 
                            ref parsedValue, 
                            true)) {

                            shapeNameID = parsedValue;
                        }
                    }
                }
            }

            return true;
        }
    }
}
