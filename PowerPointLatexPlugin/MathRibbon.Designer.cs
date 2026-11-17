namespace PowerPointLatexPlugin
{
    partial class MathRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MathRibbon()
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
            this.tglInsertLatex = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabInsert";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabInsert";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.tglInsertLatex);
            this.group1.Label = "LaTeX";
            this.group1.Name = "group1";
            // 
            // tglInsertLatex
            // 
            this.tglInsertLatex.Label = "Insert LaTeX";
            this.tglInsertLatex.Name = "tglInsertLatex";
            this.tglInsertLatex.ScreenTip = "Insert LaTeX Field";
            this.tglInsertLatex.SuperTip = "Insert a new editable LaTeX field.";
            this.tglInsertLatex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tglInsertLatex_Click);
            // 
            // MathRibbon
            // 
            this.Name = "MathRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MathRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton tglInsertLatex;
    }

    partial class ThisRibbonCollection
    {
        internal MathRibbon MathRibbon
        {
            get { return this.GetRibbon<MathRibbon>(); }
        }
    }
}
