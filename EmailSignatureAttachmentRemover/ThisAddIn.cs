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
using ImagesToRemove = EmailSignatureAttachmentRemover.Properties.Resources;
using System.Globalization;
using System.Resources;
using System.Runtime.Serialization.Formatters.Binary;

namespace EmailSignatureAttachmentRemover
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        // List<byte[]> SignatureImageHashes = new List<byte[]>();
        List<string> SignatureImageHashes = new List<string>();

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

            // var image = ImagesToRemove.ResourceManager.GetStream("image001.png");

            /*
            SignatureImageHashes.Add(GetHash(ImagesToRemove.ResourceManager.GetStream("image001")));
            SignatureImageHashes.Add(GetHash(ImagesToRemove.ResourceManager.GetStream("image002")));
            SignatureImageHashes.Add(GetHash(ImagesToRemove.ResourceManager.GetStream("image003")));
            SignatureImageHashes.Add(GetHash(ImagesToRemove.ResourceManager.GetStream("image004")));
            */


            //System.Drawing.Bitmap myStream = (System.Drawing.Bitmap)ImagesToRemove.ResourceManager.GetObject("image001");
            //myStream.Save
            //var mything = ImagesToRemove.ResourceManager.get

            /*
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("3E08DF1B9B209E867D2C2A24199D9E4C".ToCharArray()));
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("208865EF92C1D09942F1B2D349105F0B".ToCharArray()));
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("0A9220ADF16797B639C671FB1898C06F".ToCharArray()));
            SignatureImageHashes.Add(Encoding.ASCII.GetBytes("2DE1CF13DB70B611564C42DA2214AC2A".ToCharArray()));
            */

            //SignatureImageHashes.Add(GetHash(ObjectToByteArray(ImagesToRemove.ResourceManager.GetObject("image001"))));
            //SignatureImageHashes.Add(GetHash(ObjectToByteArray(ImagesToRemove.ResourceManager.GetObject("image002"))));
            //SignatureImageHashes.Add(GetHash(ObjectToByteArray(ImagesToRemove.ResourceManager.GetObject("image003"))));
            //SignatureImageHashes.Add(GetHash(ObjectToByteArray(ImagesToRemove.ResourceManager.GetObject("image004"))));

            //SignatureImageHashes.Add("3E08DF1B9B209E867D2C2A24199D9E4C");
            //SignatureImageHashes.Add("208865EF92C1D09942F1B2D349105F0B");
            //SignatureImageHashes.Add("0A9220ADF16797B639C671FB1898C06F");
            //SignatureImageHashes.Add("2DE1CF13DB70B611564C42DA2214AC2A");

            //Hashes from the outlook files:
            SignatureImageHashes.Add("58151A0BDA3DD5E858E479B1F1D775AB");
            SignatureImageHashes.Add("3ED7310957F00CAB095EEB12C2E14FD9");
            SignatureImageHashes.Add("60DE89A1F6FBCD3DAD0DE7AA7A6F6E90");
            SignatureImageHashes.Add("537CB06043EBC21F37BE29D724B75415");

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
                // MessageBox.Show(a.PathName + "|" + a.Size);

                if (a.Size == 0)
                {
                    string attachmentPath = temppath + a.FileName;
                    a.SaveAsFile(attachmentPath);
                    // FileStream savedAttachment = new FileStream(attachmentPath, FileMode.Open);
                    Stream savedAttachment = File.Open(attachmentPath, FileMode.Open);
                    // byte[] currentAttachmentHash = GetHash(ObjectToByteArray(savedAttachment));
                    string currentAttachmentStringHash = GetHashString(savedAttachment);

                    foreach (string b in SignatureImageHashes)
                    {
                        if (b == currentAttachmentStringHash)
                        {
                            MessageBox.Show("The hashes match!");
                        }

                    }
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
        /*
        private byte[] GetHash(UnmanagedMemoryStream stream)
        {
            using (var md5 = MD5.Create())
            {
                return md5.ComputeHash(stream);
            }
        }
        */

        private byte[] GetHash(Stream stream)
        {
            using (var md5 = MD5.Create())
            {
                return md5.ComputeHash(stream);
            }
        }

        private byte[] GetHash(byte[] bytes)
        {
            using (var md5 = MD5.Create())
            {
                return md5.ComputeHash(bytes);
            }
        }

        private string GetHashString(Stream stream)
        {
            using (var md5 = MD5.Create())
            {
                byte[] bytehash = md5.ComputeHash(stream);
                //return md5.ComputeHash(stream);
                string stringhash = BitConverter.ToString(bytehash).Replace("-","");
                return stringhash;
            }
        }

        //private byte[] ObjectToByteArray(object obj)
        //{
        //    BinaryFormatter bf = new BinaryFormatter();
        //    using (var ms = new MemoryStream())
        //    {
        //        bf.Serialize(ms, obj);
        //        return ms.ToArray();
        //    }
        //}

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
