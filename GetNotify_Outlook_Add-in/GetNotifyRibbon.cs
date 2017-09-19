using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace GetNotify_Outlook_Add_in
{
    public partial class GetNotifyRibbon
    {
        private void GetNotifyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            trackToggleBtn.ScreenTip = "Track Email";
            trackToggleBtn.SuperTip = "By enabling this button, all email addresses in" + 
                " \"TO\", \"CC\" and \"BCC\" fields will be tracked with GetNotify.";
        }

        private void aboutBtn_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("This Add-in is created to integrate GetNotify email tracker" +
                " service into Microsoft Outlook. You will receive notification emails" + 
                " each time a recipient opens a tracked email." +
                Environment.NewLine +
                Environment.NewLine +
                "To create a GetNotify account and learn more about the tracker visit" +
                " http://getnotify.com" +
                Environment.NewLine +
                Environment.NewLine +
                "For more information or suggestions about the add-in visit the GitHub" +
                " page located at https://github.com/omid-jf/GetNotify_Outlook_Add-in",
                "GetNotify Add-in " + 
                System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
