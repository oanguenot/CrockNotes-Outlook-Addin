using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections.Specialized;

namespace PiccataAddin
{
    public partial class ThisAddIn
    {

        private Office.CommandBar menuBar;
        private Office.CommandBarPopup newMenuBar;
        private Office.CommandBarButton buttonOne;
        private string menuTag = "WorldAddIn";
        private Outlook.Explorer explorer = null;
        private System.Collections.ArrayList selectedItems = new System.Collections.ArrayList();

        private void AddMenuBar()
        {
            menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
            newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, false);
            if (newMenuBar != null)
            {
                newMenuBar.Caption = "Piccata";
                newMenuBar.Tag = menuTag;
                buttonOne = (Office.CommandBarButton)newMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                buttonOne.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                buttonOne.Caption = "Export E-Mail(s)";
                //This is the Icon near the Text
                buttonOne.FaceId = 610;
                buttonOne.Tag = "c123";
                buttonOne.Click += new Office._CommandBarButtonEvents_ClickEventHandler(buttonOne_Click);
                //Insert Here the Button1.Click event    
                newMenuBar.Visible = true;
               
            }
        }

        private void RemoveMenubar()
        {
            // If the menu already exists, remove it.
            try
            {
                Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)
                    this.Application.ActiveExplorer().CommandBars.ActiveMenuBar.
                    FindControl(Office.MsoControlType.msoControlPopup,
                    missing, menuTag, true, true);
                if (foundMenu != null)
                {
                    foundMenu.Delete(true);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void buttonOne_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            extract();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            try
            {
                //Application.NewMail += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailEventHandler(Application_NewMail);

                //Application.NewMailEx += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailExEventHandler(Application_NewMailEx);

                //Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

                explorer = this.Application.Explorers.Application.ActiveExplorer();
                explorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(explorer_SelectionChange);

                RemoveMenubar();
                AddMenuBar();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message.ToString(), "Piccata Addin Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        void extract()
        {
            try
            {
                System.Windows.Forms.Clipboard.Clear();

                string message = "<messageList>";

                foreach (object selectedItem in selectedItems)
                {
                    Outlook.MailItem mailItem = selectedItem as Outlook.MailItem;

                    if (mailItem != null)
                    {
                        string senderEmail = mailItem.SenderEmailAddress;
                        string received = mailItem.ReceivedTime.Year.ToString() + "-" + mailItem.ReceivedTime.Month.ToString() + "-" + mailItem.ReceivedTime.Day.ToString() + " " + mailItem.ReceivedTime.Hour.ToString() + ":" + mailItem.ReceivedTime.Minute.ToString() + ":" + mailItem.ReceivedTime.Second.ToString();
                        string from = mailItem.SenderName;
                        string body = mailItem.Body;
                        string subject = mailItem.Subject;

                        message += "<message>";
                        message += "<fromDisplay>" + from + "</fromDisplay>";
                        message += "<fromAddress>" + senderEmail + "</fromAddress>";

                        message += "<subject>" + subject + "</subject>";
                        message += "<deliveryTime>" + received + "</deliveryTime>";
                        
                        if (mailItem.Attachments.Count > 0)
                        {
                            message += "<attachments>";
                            for (int i = 1; i <= mailItem.Attachments.Count; i++)
                            {
                                string name = mailItem.Attachments[i].FileName;
                                string path = GetTempDir();
                                
                                mailItem.Attachments[i].SaveAsFile(path + name);
                                
                                message += "<attachment>";
                                message += "<fileName>" +  path + name + "</fileName>";
                                message += "<fileType>" + mailItem.Attachments[i].Type + "</fileType>";
                                message += "</attachment>";
                               
                            }
                            message += "</attachments>";
                        }

                        message += "<contentText><![CDATA[" + body +"]]></contentText>";
                        message += "</message>";
                    }
                }
                message += "</messageList>";

                System.Windows.Forms.Clipboard.SetText(message, System.Windows.Forms.TextDataFormat.Text);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Extract Error: " + ex.Message.ToString(), "Piccata Addin Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

        }

        public static String GetMyDocumentsDir()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        }

        public static String GetTempDir()
        {
            return System.IO.Path.GetTempPath();
        }

        void explorer_SelectionChange()
        {
            selectedItems.Clear();

            try
            {
                foreach (object selectedItem in explorer.Selection)
                {
                    selectedItems.Add(selectedItem);
                }
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message.ToString(), "Piccata Addin Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

        }

        /*
        void Application_NewMailEx(string EntryIDCollection)
        {
           
        }

        void Application_NewMail()
        {
            
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
          
        }
         */

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
           
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
