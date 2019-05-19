namespace AprioriAlgorithm
{
    partial class Design : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Design()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Design));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.SupportValue = this.Factory.CreateRibbonEditBox();
            this.btn_Uygula = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btn_Info = this.Factory.CreateRibbonButton();
            this.btn_Help = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "APRIORI ALGORITMASI";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.SupportValue);
            this.group1.Items.Add(this.btn_Uygula);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btn_Info);
            this.group1.Items.Add(this.btn_Help);
            this.group1.Label = "CREATED BY KAYACI";
            this.group1.Name = "group1";
            // 
            // SupportValue
            // 
            this.SupportValue.Label = "Support Value :";
            this.SupportValue.MaxLength = 4;
            this.SupportValue.Name = "SupportValue";
            this.SupportValue.Text = "0,00";
            // 
            // btn_Uygula
            // 
            this.btn_Uygula.Image = ((System.Drawing.Image)(resources.GetObject("btn_Uygula.Image")));
            this.btn_Uygula.Label = "Uygula";
            this.btn_Uygula.Name = "btn_Uygula";
            this.btn_Uygula.ShowImage = true;
            this.btn_Uygula.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Uygula_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btn_Info
            // 
            this.btn_Info.Image = ((System.Drawing.Image)(resources.GetObject("btn_Info.Image")));
            this.btn_Info.Label = "Hakkında";
            this.btn_Info.Name = "btn_Info";
            this.btn_Info.ShowImage = true;
            this.btn_Info.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Info_Click);
            // 
            // btn_Help
            // 
            this.btn_Help.Image = ((System.Drawing.Image)(resources.GetObject("btn_Help.Image")));
            this.btn_Help.Label = "Yardım";
            this.btn_Help.Name = "btn_Help";
            this.btn_Help.ShowImage = true;
            this.btn_Help.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Help_Click);
            // 
            // Design
            // 
            this.Name = "Design";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Design_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Uygula;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SupportValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Info;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Help;
    }

    partial class ThisRibbonCollection
    {
        internal Design Ribbon1
        {
            get { return this.GetRibbon<Design>(); }
        }
    }
}
