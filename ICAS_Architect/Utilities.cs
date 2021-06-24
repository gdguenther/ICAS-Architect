using System;
using System.Windows.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Permissions;
using System.Windows.Forms;

namespace Utilites
{
    /// <summary>
    /// Emulates the VB6 DoEvents to refresh a window during long running events
    /// </summary>
    public class ScreenEvents
    {
        [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        public static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
                new DispatcherOperationCallback(ExitFrame), frame);
            Dispatcher.PushFrame(frame);
        }

        public static object ExitFrame(object f)
        {
            ((DispatcherFrame)f).Continue = false;

            return null;
        }

        public static void DisplayVisioStatus(string message)
        {
            ICAS_Architect.Globals.ThisAddIn.Application.QueueMarkerEvent("ScopeStart");
            //            Microsoft.Office.Interop.Visio.Shape shpStatus = Globals.ThisAddIn.Application.ActivePage.DrawRectangle(0, 0, 10, .25);
            Microsoft.Office.Interop.Visio.Shape shpStatus = ICAS_Architect.Globals.ThisAddIn.Application.ActivePage.DrawRectangle(0, 0, 4, .25);
            shpStatus.Text = message;
            Utilites.ScreenEvents.DoEvents();
            shpStatus.Delete();
            ICAS_Architect.Globals.ThisAddIn.Application.QueueMarkerEvent("ScopeEnd");
        }

        internal static DialogResult ShowInputDialog(ref string input, ref string input2, string label1="", string label2="", string message="")
        {
            System.Drawing.Size size = new System.Drawing.Size(360, input2 == "-1" ? 110 : 140);
            Form inputBox = new Form();
            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Name";

            Label label = new Label();
            label.Size = new System.Drawing.Size(size.Width - 20, 23);
            label.Location = new System.Drawing.Point(5, 10);
            label.Text = message;
            inputBox.Controls.Add(label);

            System.Windows.Forms.TextBox textBox = new TextBox();
            textBox.Size = new System.Drawing.Size(size.Width - 70, 23);
            textBox.Location = new System.Drawing.Point(65, 35);
            textBox.Text = input;
            inputBox.Controls.Add(textBox);
            Label labelBox = new Label();
            labelBox.Size = new System.Drawing.Size(55, 23);
            labelBox.Location = new System.Drawing.Point(5, 35);
            labelBox.Text = label1;
            inputBox.Controls.Add(labelBox);

            System.Windows.Forms.TextBox textBox2 = new TextBox();
            textBox2.Size = new System.Drawing.Size(size.Width - 70, 23);
            textBox2.Location = new System.Drawing.Point(65, 63);
            textBox2.Text = input2;
            Label labelBox2 = new Label();
            labelBox2.Size = new System.Drawing.Size(55, 23);
            labelBox2.Location = new System.Drawing.Point(5, 63);
            labelBox2.Text = label2;
            if (input2 != "-1") {
                inputBox.Controls.Add(textBox2);
                inputBox.Controls.Add(labelBox2);
            }

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, input2=="-1" ? 69 : 99);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 80, input2 == "-1" ? 69 : 99);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            inputBox.StartPosition = FormStartPosition.Manual;
            inputBox.Location = new System.Drawing.Point(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y);
            DialogResult result = inputBox.ShowDialog();
            input = textBox.Text;
            input2 = textBox2.Text;
            return result;
        }
    }
}