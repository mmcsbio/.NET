namespace BayantechAddIn
{
    partial class BayantechRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public BayantechRibbon()
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
            this.tab_bayantech = this.Factory.CreateRibbonTab();
            this.grp_bidi = this.Factory.CreateRibbonGroup();
            this.btn_fix_bidi = this.Factory.CreateRibbonButton();
            this.tab_bayantech.SuspendLayout();
            this.grp_bidi.SuspendLayout();
            // 
            // tab_bayantech
            // 
            this.tab_bayantech.Groups.Add(this.grp_bidi);
            this.tab_bayantech.Label = "Bayantech";
            this.tab_bayantech.Name = "tab_bayantech";
            // 
            // grp_bidi
            // 
            this.grp_bidi.Items.Add(this.btn_fix_bidi);
            this.grp_bidi.Label = "Bidirectional Languages";
            this.grp_bidi.Name = "grp_bidi";
            // 
            // btn_fix_bidi
            // 
            this.btn_fix_bidi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_fix_bidi.Label = "Fix Bidi Issues";
            this.btn_fix_bidi.Name = "btn_fix_bidi";
            this.btn_fix_bidi.ShowImage = true;
            this.btn_fix_bidi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_fix_bidi_Click);
            // 
            // BayantechRibbon
            // 
            this.Name = "BayantechRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab_bayantech);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.BayantechRibbon_Load);
            this.tab_bayantech.ResumeLayout(false);
            this.tab_bayantech.PerformLayout();
            this.grp_bidi.ResumeLayout(false);
            this.grp_bidi.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_bayantech;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_bidi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_fix_bidi;
    }

    partial class ThisRibbonCollection
    {
        internal BayantechRibbon BayantechRibbon
        {
            get { return this.GetRibbon<BayantechRibbon>(); }
        }
    }
}
