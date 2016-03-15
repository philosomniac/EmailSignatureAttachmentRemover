using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.IO;


namespace EmailSignatureAttachmentRemover
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            /*
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem!= null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }
            }
            */
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            Cancel = true;
            // int numberOfAttachments = Item.Attachments.Count;
            Outlook.MailItem m = (Outlook.MailItem)Item;
            string temppath = Path.GetTempPath();

            m.SaveAs(temppath + "tempmailitem.msg");
            

            // Application_ItemSend(m, ref false);

            foreach (Outlook.Attachment a in m.Attachments)
            {
                MessageBox.Show(a.PathName + "|" + a.Size);

                //a.SaveAsFile(Path.GetTempPath() + a.FileName);
                //FileStream matchingfile = File.Open(Path.GetTempPath() + a.FileName, FileMode.Open);

                
                //// int attachmentHash = a.GetHashCode();

                
                //using (var md5 = MD5.Create())
                //{
                //        byte[] hash = md5.ComputeHash(matchingfile);
                //        string realhash = BitConverter.ToString(hash);

                //}
                
                // MessageBox.Show("There is an attachment called " + a.FileName + " in this message." + " It has a filesize of: " + a.Size);   



            }
            

        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
