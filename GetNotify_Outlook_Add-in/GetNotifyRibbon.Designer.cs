namespace GetNotify_Outlook_Add_in
{
    partial class GetNotifyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public GetNotifyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.trackToggleBtn = this.Factory.CreateRibbonToggleButton();
            this.aboutBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "GetNotify";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.trackToggleBtn);
            this.group1.Label = "Tracking";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.aboutBtn);
            this.group2.Label = "Help";
            this.group2.Name = "group2";
            // 
            // trackToggleBtn
            // 
            this.trackToggleBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.trackToggleBtn.Image = global::GetNotify_Outlook_Add_in.Properties.Resources.email_track_icon;
            this.trackToggleBtn.Label = "Track Email";
            this.trackToggleBtn.Name = "trackToggleBtn";
            this.trackToggleBtn.ShowImage = true;
            // 
            // aboutBtn
            // 
            this.aboutBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.aboutBtn.Image = global::GetNotify_Outlook_Add_in.Properties.Resources.info_icon;
            this.aboutBtn.Label = "About";
            this.aboutBtn.Name = "aboutBtn";
            this.aboutBtn.ShowImage = true;
            this.aboutBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.aboutBtn_Click);
            // 
            // GetNotifyRibbon
            // 
            this.Name = "GetNotifyRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.GetNotifyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton trackToggleBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton aboutBtn;
    }

    partial class ThisRibbonCollection
    {
        internal GetNotifyRibbon GetNotifyRibbon
        {
            get { return this.GetRibbon<GetNotifyRibbon>(); }
        }
    }
}
