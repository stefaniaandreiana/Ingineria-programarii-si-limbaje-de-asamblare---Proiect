namespace ProjectIP_2
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.buttonVizualizeazaOferte = this.Factory.CreateRibbonButton();
            this.buttonDeschideExcel = this.Factory.CreateRibbonButton();
            this.buttonGenereazaContract = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonVizualizeazaOferte);
            this.group1.Items.Add(this.buttonDeschideExcel);
            this.group1.Items.Add(this.buttonGenereazaContract);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // buttonVizualizeazaOferte
            // 
            this.buttonVizualizeazaOferte.Label = "Vizualizeaza Oferte PPT";
            this.buttonVizualizeazaOferte.Name = "buttonVizualizeazaOferte";
            this.buttonVizualizeazaOferte.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonVizualizeazaOferte_Click);
            // 
            // buttonDeschideExcel
            // 
            this.buttonDeschideExcel.Label = "Deschide Excel";
            this.buttonDeschideExcel.Name = "buttonDeschideExcel";
            this.buttonDeschideExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDeschideExcel_Click);
            // 
            // buttonGenereazaContract
            // 
            this.buttonGenereazaContract.Label = "Genereaza Contract";
            this.buttonGenereazaContract.Name = "buttonGenereazaContract";
            this.buttonGenereazaContract.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenereazaContract_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonPPT_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonVizualizeazaOferte;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeschideExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGenereazaContract;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon RibbonPPT
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
