namespace DeckLanguageAddIn
{
    partial class LanguageSelector : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public LanguageSelector()
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
            this.dropDownLanguage = this.Factory.CreateRibbonDropDown();
            this.buttonSetLanguage = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabReview";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabReview";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.dropDownLanguage);
            this.group1.Items.Add(this.buttonSetLanguage);
            this.group1.Label = "Deck Language";
            this.group1.Name = "group1";
            // 
            // dropDownLanguage
            // 
            this.dropDownLanguage.Label = "Language";
            this.dropDownLanguage.Name = "dropDownLanguage";
            this.dropDownLanguage.ShowLabel = false;
            // 
            // buttonSetLanguage
            // 
            this.buttonSetLanguage.Label = "Set";
            this.buttonSetLanguage.Name = "buttonSetLanguage";
            this.buttonSetLanguage.ScreenTip = "Set language";
            this.buttonSetLanguage.SuperTip = "Sets the language on all slides to the selected language.\n\nThe language will be r" +
                "eset on both slides and notes.";
            this.buttonSetLanguage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSetLanguage_Click);
            // 
            // LanguageSelector
            // 
            this.Name = "LanguageSelector";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.LanguageSelector_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSetLanguage;
    }

    partial class ThisRibbonCollection
    {
        internal LanguageSelector LanguageSelector
        {
            get { return this.GetRibbon<LanguageSelector>(); }
        }
    }
}
