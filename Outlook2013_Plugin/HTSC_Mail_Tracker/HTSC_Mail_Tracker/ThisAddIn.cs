using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;

namespace HTSC_Mail_Tracker
{
    public partial class ThisAddIn
    {
        //private FileStream fileLog;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //fileLog = new FileStream("C:\\Outlook_AddInLog.txt", FileMode.OpenOrCreate | FileMode.Append, FileAccess.ReadWrite);
            //String str = @"\nStart Up Outlook";
            //byte[] toBytes = Encoding.ASCII.GetBytes(str);

            //fileLog.Write(toBytes, 0, toBytes.Length);            
            //Console.WriteLine(str);

            this.Application.NewMail += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailEventHandler(ThisApplication_NewMail);

            System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();

            myTimer.Tick += new EventHandler(TimerEventProcessor);

            // Sets the timer interval to 60 seconds.
            myTimer.Interval = 60000;
            myTimer.Start();

            //MessageBox.Show(@"Loaded");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
            //fileLog.Close();

        }

        private void ThisApplication_NewMail()
        {
            Outlook.MAPIFolder inBox = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items inBoxItems = inBox.Items;
            Outlook.MailItem newEmail = null;
            inBoxItems = inBoxItems.Restrict("[Unread] = true");

            try
            {
                foreach (object collectionItem in inBoxItems)
                {
                    newEmail = collectionItem as Outlook.MailItem;
                    if (newEmail != null)
                    {
                        if ((newEmail.Attachments.Count > 0) && (newEmail.Subject.Contains("AEWA FINAL APPROVAL")))
                        {
                            for (int i = 1; i <= newEmail.Attachments.Count; i++)
                            {
                                newEmail.Attachments[i].SaveAsFile(@"C:\Temp\" + newEmail.Attachments[i].FileName);
                            }
                        }
                        else
                        {
                            String str = @"\nAttachment not expected" + newEmail.Subject;
                            //byte[] toBytes = Encoding.ASCII.GetBytes(str);
                            //fileLog.Write(toBytes, 0, toBytes.Length);
                            //Console.WriteLine(str);

                            //MessageBox.Show(str);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errorInfo = (string)ex.Message.Substring(0, 11);
                if (errorInfo == "Cannot save")
                {
                    //String str = @"\nCannot Save";
                    //byte[] toBytes = Encoding.ASCII.GetBytes(str);
                    //fileLog.Write(toBytes, 0, toBytes.Length);
                    //Console.WriteLine(str);

                    MessageBox.Show(@"Create Folder C:\Temp");
                }
            }
        }
        // This is the method to run when the timer is raised.
        private void TimerEventProcessor(Object myObject,
                                                EventArgs myEventArgs)
        {
            this.ThisApplication_NewMail();
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
            //this.Application.NewMail += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailEventHandler(ThisApplication_NewMail);

        }

        

        #endregion
    }
}
