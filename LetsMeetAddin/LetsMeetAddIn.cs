using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Drawing;


namespace LetsMeetAddin
{
    public partial class LetsMeetAddIn
    {
        Outlook.Inspectors inspectors;
        private static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "abcdefghijklmnopqrstuvwxyz0123456789";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Settings_ribbon();
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {

            Outlook.AppointmentItem agendaMeeting = Inspector.CurrentItem as Outlook.AppointmentItem;

            if (agendaMeeting != null)
            {

                if (agendaMeeting.Location is null && agendaMeeting.Subject is null && agendaMeeting.Body is null)
                {
                    string meetid = RandomString(20);
                    agendaMeeting.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
                    agendaMeeting.Location = Properties.Settings.Default.MeetName + " --- " + Properties.Settings.Default.MeetLink + meetid;
                    RichTextBox dynamicRichTextBox = new RichTextBox();
                    dynamicRichTextBox.BackColor = Color.White;
                    dynamicRichTextBox.Clear();
                    dynamicRichTextBox.SelectionFont = new Font("Calibri", 11);
                    dynamicRichTextBox.SelectedText = "\n\n";
                    dynamicRichTextBox.BulletIndent = 0;
                    dynamicRichTextBox.SelectionBullet = false;
                    dynamicRichTextBox.SelectionFont = new Font("Arial", 16);
                    dynamicRichTextBox.AppendText("___________________________\n\n");
                    dynamicRichTextBox.SelectionFont = new Font("Arial", 16);
                    dynamicRichTextBox.SelectedText = Properties.Settings.Default.MeetName + "\n\n";
                    dynamicRichTextBox.SelectionFont = new Font("Arial", 12);
                    dynamicRichTextBox.SelectedText = Properties.Settings.Default.MeetDesc + "\n\n";
                    dynamicRichTextBox.SelectionFont = new Font("Arial", 12);
                    dynamicRichTextBox.SelectedText = Properties.Settings.Default.MeetLink + meetid + "\n";
                    agendaMeeting.RTFBody = System.Text.Encoding.UTF8.GetBytes(dynamicRichTextBox.Rtf);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
