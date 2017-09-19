using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GetNotify_Outlook_Add_in
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.Application application = this.Application;
            application.ItemSend +=
                new Outlook.ApplicationEvents_11_ItemSendEventHandler(ItemSend_BeforeSend);
        }

        void ItemSend_BeforeSend(object item, ref bool cancel)
        {
            ThisRibbonCollection ribbonCollection =
                Globals.Ribbons[Globals.ThisAddIn.Application.ActiveInspector()];
            Outlook.MailItem mailItem = (Outlook.MailItem)item;

            // Check to see if track button is enabled.
            if (ribbonCollection.GetNotifyRibbon.trackToggleBtn.Checked)
                if (mailItem != null)
                {
                    Outlook.Recipients recips = mailItem.Recipients;
                    string[] emailAdrses = new String[recips.Count];
                    int[] emailTypes = new int[recips.Count];
                    int arrayItr = 0;

                    // Lets add email addresses and their types in new arrays,
                    // modify the address and remove current recipients. 
                    while (recips.Count > 0)
                    {
                        Outlook.Recipient recip = recips[1];

                        if (recip.Address.Split('@')[1].ToLower().Contains("getnotify.com"))
                            emailAdrses[arrayItr] = recip.Address;
                        else
                            emailAdrses[arrayItr] = recip.Address + ".getnotify.com";

                        emailTypes[arrayItr] = recip.Type;
                        arrayItr++;
                        recips.Remove(1);
                    }

                    // Add new recipients using the arrays we populated before.
                    for (int i = 0; i < emailAdrses.Length; i++)
                    {
                        // We use recipient's address as its name.
                        Outlook.Recipient newRecip = recips.Add(emailAdrses[i]);
                        newRecip.AddressEntry.Address = emailAdrses[i];
                        newRecip.Type = emailTypes[i];
                    }
                }

            cancel = false;
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
