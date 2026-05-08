
using System.Windows.Forms;
namespace TestRunViewerSqlite
{
    partial class MainForm
    {
        private TextBox txtPlanRunName;
        private Button btnOpenDb; private TextBox txtDbPath; private Button btnLoad;
        private DataGridView dgvOverview; private DataGridView dgvDetails;
        private DataGridView dgvStatsAll; private DataGridView dgvStatsSerial;
        private ComboBox cboVerdict; private ComboBox cboStatus; private DateTimePicker dtFrom; private DateTimePicker dtTo;
        private ComboBox cmbMapPlanRunNumber; private TextBox txtSerial;
        private Button btnExportOverview; private Button btnExportDetails; private Button btnExportStatsAll; private Button btnExportStatsSerial;
        private Button btnApplyFilters; private Button btnLoadStatsAll; private Button btnLoadStatsSerial;
        private ToolStripStatusLabel lblStatus; private Label lblStatusFilter;
        private GroupBox grpFilters; private GroupBox grpMapping; private GroupBox grpStats; private StatusStrip statusStrip1;
        private void InitializeComponent()
        {
            this.btnOpenDb = new Button(); this.txtDbPath = new TextBox(); this.btnLoad = new Button();
            this.dgvOverview = new DataGridView(); this.dgvDetails = new DataGridView();
            this.dgvStatsAll = new DataGridView(); this.dgvStatsSerial = new DataGridView();
            this.cboVerdict = new ComboBox(); this.cboStatus = new ComboBox(); this.dtFrom = new DateTimePicker(); this.dtTo = new DateTimePicker();
            this.cmbMapPlanRunNumber = new ComboBox(); this.txtSerial = new TextBox();
            this.btnExportOverview = new Button(); this.btnExportDetails = new Button(); this.btnExportStatsAll = new Button(); this.btnExportStatsSerial = new Button();
            this.btnApplyFilters = new Button(); this.btnLoadStatsAll = new Button(); this.btnLoadStatsSerial = new Button();
            this.lblStatusFilter = new Label(); this.grpFilters = new GroupBox(); this.grpMapping = new GroupBox(); this.grpStats = new GroupBox(); this.statusStrip1 = new StatusStrip(); this.lblStatus = new ToolStripStatusLabel();
            var split = new SplitContainer(); var tabs = new TabControl(); var tabDetails = new TabPage("Details"); var tabAll = new TabPage("Stats (All)"); var tabSer = new TabPage("Stats (By Serial)"); var pnlSer = new Panel(); var lblSerial = new Label();
            this.SuspendLayout();
            // Top bar
            this.btnOpenDb.Text = "Open DB..."; this.btnOpenDb.SetBounds(12,12,120,32); this.btnOpenDb.Click += btnOpenDb_Click;
            this.txtDbPath.ReadOnly = true; this.txtDbPath.SetBounds(138,14,500,28);
            this.btnLoad.Text = "Load Overview"; this.btnLoad.SetBounds(650,12,150,32); this.btnLoad.Click += btnLoad_Click;
            // Filters group
            this.grpFilters.Text = "Filters"; this.grpFilters.SetBounds(12,52,788,80);
            this.cboVerdict.DropDownStyle = ComboBoxStyle.DropDownList; this.cboVerdict.SetBounds(10,34,100,28);
            var lblVerdict = new Label(){Text="Verdict:", AutoSize=true}; lblVerdict.SetBounds(10,10,60,20);
            this.grpFilters.Controls.Add(lblVerdict); this.grpFilters.Controls.Add(this.cboVerdict);
            this.lblStatusFilter.Text = "Status:"; this.lblStatusFilter.AutoSize = true; this.lblStatusFilter.SetBounds(120,10,60,20);
            this.cboStatus.DropDownStyle = ComboBoxStyle.DropDownList; this.cboStatus.SetBounds(120,34,100,28);
            this.grpFilters.Controls.Add(this.lblStatusFilter); this.grpFilters.Controls.Add(this.cboStatus);
            var lblFrom = new Label(){Text="From:", AutoSize=true}; lblFrom.SetBounds(230,10,60,20);
            this.dtFrom.Format = DateTimePickerFormat.Custom; this.dtFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss"; this.dtFrom.SetBounds(230,34,180,28);
            var lblTo = new Label(){Text="To:", AutoSize=true}; lblTo.SetBounds(420,10,60,20);
            this.dtTo.Format = DateTimePickerFormat.Custom; this.dtTo.CustomFormat = "yyyy-MM-dd HH:mm:ss"; this.dtTo.SetBounds(420,34,180,28);
            this.btnApplyFilters.Text = "Apply"; this.btnApplyFilters.SetBounds(610,32,80,30); this.btnApplyFilters.Click += btnApplyFilters_Click;
            this.grpFilters.Controls.AddRange(new Control[]{lblFrom, this.dtFrom, lblTo, this.dtTo, this.btnApplyFilters});
            // Mapping group
            this.grpMapping.Text = "Column Mapping"; this.grpMapping.SetBounds(12,138,788,56);
            var lblMap = new Label(){Text="PlanRunNumber column:", AutoSize=true}; lblMap.SetBounds(10,25,180,20);
            this.cmbMapPlanRunNumber.DropDownStyle = ComboBoxStyle.DropDownList; this.cmbMapPlanRunNumber.SetBounds(200,20,570,28);
            this.grpMapping.Controls.AddRange(new Control[]{lblMap, this.cmbMapPlanRunNumber});
            // Stats group
            this.grpStats.Text = "Statistics"; this.grpStats.SetBounds(12,196,788,56);
            this.btnLoadStatsAll.Text = "Load Stats (All)"; this.btnLoadStatsAll.SetBounds(12,20,150,30); this.btnLoadStatsAll.Click += btnLoadStatsAll_Click;
            this.btnExportOverview.Text = "Export Overview..."; this.btnExportOverview.SetBounds(620,20,150,30); this.btnExportOverview.Click += btnExportOverview_Click;
            this.grpStats.Controls.AddRange(new Control[]{this.btnLoadStatsAll, this.btnExportOverview});
            // Split bottom
            split.SetBounds(12,258,788,440); split.Orientation = Orientation.Horizontal; split.SplitterDistance=210; split.Anchor = AnchorStyles.Left|AnchorStyles.Top|AnchorStyles.Right|AnchorStyles.Bottom;
            this.dgvOverview.ReadOnly=true; this.dgvOverview.SelectionMode = DataGridViewSelectionMode.FullRowSelect; this.dgvOverview.Dock = DockStyle.Fill; this.dgvOverview.SelectionChanged += dgvOverview_SelectionChanged;
            var containerTop = new Panel(){Dock=DockStyle.Fill}; containerTop.Controls.Add(this.dgvOverview);
            split.Panel1.Controls.Add(containerTop);
            // Tabs
            tabs.Dock = DockStyle.Fill;
            // Details tab
            this.dgvDetails.ReadOnly=true; this.dgvDetails.Dock = DockStyle.Fill; this.btnExportDetails.Text = "Export..."; this.btnExportDetails.Dock = DockStyle.Top; this.btnExportDetails.Click += btnExportDetails_Click;
            tabDetails.Controls.Add(this.dgvDetails); tabDetails.Controls.Add(this.btnExportDetails);
            // Stats (All)
            this.dgvStatsAll.ReadOnly=true; this.dgvStatsAll.Dock = DockStyle.Fill; this.btnExportStatsAll.Text = "Export..."; this.btnExportStatsAll.Dock = DockStyle.Top; this.btnExportStatsAll.Click += btnExportStatsAll_Click;
            tabAll.Controls.Add(this.dgvStatsAll); tabAll.Controls.Add(this.btnExportStatsAll);
            // Stats by Serial
            this.dgvStatsSerial.ReadOnly=true; this.dgvStatsSerial.Dock = DockStyle.Fill; this.btnLoadStatsSerial.Text = "Load Stats"; this.btnLoadStatsSerial.Click += btnLoadStatsSerial_Click; this.btnExportStatsSerial.Text = "Export..."; this.btnExportStatsSerial.Click += btnExportStatsSerial_Click; lblSerial.Text = "Serial:"; lblSerial.AutoSize=true; lblSerial.SetBounds(3,10,50,20); this.txtSerial.SetBounds(60,6,150,28); this.btnLoadStatsSerial.SetBounds(220,5,100,30); this.btnExportStatsSerial.SetBounds(330,5,100,30);
            var pnlSer1= new Panel(){Dock=DockStyle.Top, Height=40}; pnlSer1.Controls.AddRange(new Control[]{lblSerial,this.txtSerial,this.btnLoadStatsSerial,this.btnExportStatsSerial});
            tabSer.Controls.Add(this.dgvStatsSerial); tabSer.Controls.Add(pnlSer1);
            tabs.TabPages.AddRange(new[]{tabDetails,tabAll,tabSer});
            split.Panel2.Controls.Add(tabs);
            // Status
            this.statusStrip1.Items.Add(this.lblStatus); this.lblStatus.Text = "Ready."; this.statusStrip1.Dock = DockStyle.Bottom;
            // Form
            this.ResumeLayout(false);
            this.txtPlanRunName = new TextBox();
            this.txtPlanRunName.PlaceholderText = "PlanRun Name contains...";
            this.txtPlanRunName.SetBounds(10, 60, 300, 28);
            //this.grpFilters.Controls.Add(this.txtPlanRunName);
            
            this.Controls.AddRange(new Control[]{this.btnOpenDb,this.txtDbPath,this.btnLoad,this.grpFilters,this.grpMapping,this.grpStats,split,this.statusStrip1, this.txtPlanRunName });
            this.Text = "Test Run Viewer (SQLite)"; this.ClientSize = new System.Drawing.Size(812,720);
            

        }
    }
}
