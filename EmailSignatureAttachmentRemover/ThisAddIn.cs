using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.IO;
/*
using ImagesToRemove = EmailSignatureAttachmentRemover.Properties.Resources;
using System.Globalization;
using System.Resources;
*/

namespace EmailSignatureAttachmentRemover
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        List<byte[]> SignatureImageHashes = new List<byte[]>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

            /*
            ResourceSet EmailSignatureImages = ImagesToRemove.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true);
            foreach (DictionaryEntry entry in EmailSignatureImages)
            {
                FileStream s = new FileStream(entry.Value)
            }
            */
            
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("3E08DF1B9B209E867D2C2A24199D9E4C".ToCharArray()));
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("208865EF92C1D09942F1B2D349105F0B".ToCharArray()));
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("0A9220ADF16797B639C671FB1898C06F".ToCharArray()));
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("2DE1CF13DB70B611564C42DA2214AC2A".ToCharArray()));
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
            string fullpath = temppath + "tempmailitem.msg";

            m.SaveAs(fullpath);
            /*
            Outlook.MailItem savedMailItem = Application.Session.OpenSharedItem(fullpath);

            foreach (Outlook.Attachment a in savedMailItem.Attachments)
            {
                MessageBox.Show(a.PathName + "|" + a.Size);
            }

            */

            // Application_ItemSend(m, ref false);

            foreach (Outlook.Attachment a in m.Attachments)
            {
                MessageBox.Show(a.PathName + "|" + a.Size);

                if (a.Size == 0)
                {
                    string attachmentPath = temppath + a.FileName;
                    a.SaveAsFile(attachmentPath);
                    FileStream savedAttachment = new FileStream(attachmentPath, FileMode.Open);
                }

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

        private byte[] GetHash(FileStream stream)
        {
            using (var md5 = MD5.Create())
            {
                return md5.ComputeHash(stream);
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
