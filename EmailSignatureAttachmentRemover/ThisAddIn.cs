using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace EmailSignatureAttachmentRemover
{
    public partial class ThisAddIn
    {

        const string TARGET_EMAIL_ADDRESS = "techsupport@apexrevtech.com";
        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }


        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            //Cancel = true;

            Outlook.MailItem m = (Outlook.MailItem)Item;
            string temppath = Path.GetTempPath();
            int currentAttachmentIndex = 1; // office interop indices start at 1.
            int attachmentCount = m.Attachments.Count;
            string notificationMessage = "";
            bool beingSentToSpiceworks = false;


            //foreach (Outlook.Recipient r in m.Recipients)
            //{
            //    if (r.Address.ToLower().Trim() == TARGET_EMAIL_ADDRESS)
            //    {
            //        beingSentToSpiceworks = true;
            //        break;
            //    }
            //}

            foreach (Outlook.Recipient r in m.Recipients)
            {
                Outlook.PropertyAccessor pa = r.PropertyAccessor;
                string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS);
                if (smtpAddress.ToLower().Trim() == TARGET_EMAIL_ADDRESS)
                {
                    beingSentToSpiceworks = true;
                    break;
                }
            }


            try {
                if (beingSentToSpiceworks)
                {
                    Regex unnamedImageAttachmentPattern = new Regex(@"image0\d\d\.png|image0\d\d.jpg"); // this is the format outlook chooses for unnamed image attachments
                    int minAttachmentSize = 9000; // The apex logo clocks in around 8600 bytes, so 9000 should allow for some overhead
                    m.SaveAs(temppath + "tempoutlookmessage.msg"); // need to save before modifying the message or outlook gets unhappy.

                    while (currentAttachmentIndex <= m.Attachments.Count)
                    {
                        Outlook.Attachment a = m.Attachments[currentAttachmentIndex];
                        if (unnamedImageAttachmentPattern.IsMatch(a.FileName))
                        {

                            if (a.Size > 0)
                            {
                                // MessageBox.Show("This is an attachment that has not been saved manually with a name of: " + a.FileName + " and size of" + a.Size);
                                if (a.Size < minAttachmentSize)
                                {
                                    notificationMessage += a.FileName;
                                    a.Delete();
                                    currentAttachmentIndex--;
                                    notificationMessage += Environment.NewLine;
                                }
                            }

                            else
                            {

                                string attachmentPath = temppath + a.FileName;

                                a.SaveAsFile(attachmentPath);
                                Stream savedAttachment = File.Open(attachmentPath, FileMode.Open);

                                // MessageBox.Show("This is an attachment that was saved manually with name: " + a.FileName + " and filesize of: " + savedAttachment.Length);

                                if (savedAttachment.Length < minAttachmentSize)
                                {
                                    notificationMessage += a.FileName;
                                    a.Delete();
                                    currentAttachmentIndex--;
                                    notificationMessage += Environment.NewLine;
                                }

                                savedAttachment.Close();

                            }

                        }

                        Marshal.ReleaseComObject(a); // not sure what this does but the internet says it's good practice
                        currentAttachmentIndex++;
                    }

                    if (notificationMessage.Length > 0) // means we removed attachments and should let the user know that. Is this a sloppy check?
                    {

                        NotifyIcon ni = new NotifyIcon();
                        ni.Visible = true;

                        ni.BalloonTipTitle = "The following attachments were removed:";

                        ni.BalloonTipText = notificationMessage;
                        ni.Icon = System.Drawing.SystemIcons.Application;
                        ni.ShowBalloonTip(20000);
                        ni.Dispose();
                    }
                }
            }

            catch
            {
                NotifyIcon ni = new NotifyIcon();
                ni.Visible = true;

                ni.BalloonTipTitle = "Outlook attachment remover error";

                ni.BalloonTipText = "Something went wrong with the outlook attachment remover addon.";
                ni.Icon = System.Drawing.SystemIcons.Application;
                ni.ShowBalloonTip(20000);
                ni.Dispose();

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
