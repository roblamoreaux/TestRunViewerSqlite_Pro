
using System.Windows.Forms;
namespace TestRunViewerSqlite
{
    partial class MainForm
    {
        private TextBox txtPlanRunName;
        private Button btnOpenDb; 
        private TextBox txtDbPath; 
        private Button btnLoad;
        private DataGridView dgvOverview; 
        private DataGridView dgvDetails;
        private DataGridView dgvSummary;
        private DataGridView dgvStatsAll;
        private DataGridView dgvStatsSerial;
        private ComboBox cboVerdict; 
        private ComboBox cboStatus; 
        private DateTimePicker dtFrom; 
        private DateTimePicker dtTo;
        private ComboBox cmbMapPlanRunNumber; 
        private Button btnExportOverview; 
        private Button btnExportDetails; 
        private Button btnExportStatsAll; 
        private Button btnExportStatsSerial;
        private Button btnApplyFilters;
        private CheckBox chkWithSerialNumber;
        private Button btnGetDetails;
        private Button btnLoadStatsAll; 
        private Button btnLoadStatsSerial;
        private ToolStripStatusLabel lblStatus;
        private Label lblStatusFilter;
        private GroupBox grpFilters; 
        private GroupBox grpMapping; 
        private GroupBox grpStats; 
        private StatusStrip statusStrip1;
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            btnOpenDb = new Button();
            txtDbPath = new TextBox();
            btnLoad = new Button();
            dgvOverview = new DataGridView();
            dgvDetails = new DataGridView();
            dgvSummary = new DataGridView();
            dgvStatsAll = new DataGridView();
            dgvStatsSerial = new DataGridView();
            cboVerdict = new ComboBox();
            cboStatus = new ComboBox();
            dtFrom = new DateTimePicker();
            dtTo = new DateTimePicker();
            cmbMapPlanRunNumber = new ComboBox();
            btnExportOverview = new Button();
            btnExportDetails = new Button();
            btnExportStatsAll = new Button();
            btnExportStatsSerial = new Button();
            btnApplyFilters = new Button();
            chkWithSerialNumber = new CheckBox();
            btnGetDetails = new Button();
            btnLoadStatsAll = new Button();
            btnLoadStatsSerial = new Button();
            lblStatusFilter = new Label();
            grpFilters = new GroupBox();
            lblVerdict = new Label();
            lblFrom = new Label();
            lblTo = new Label();
            grpMapping = new GroupBox();
            lblMap = new Label();
            grpStats = new GroupBox();
            lblSerial = new Label();
            txtSerial = new TextBox();
            statusStrip1 = new StatusStrip();
            lblStatus = new ToolStripStatusLabel();
            split = new SplitContainer();
            containerTop = new Panel();
            tabs = new TabControl();
            tabSummary = new TabPage();
            tabDetails = new TabPage();
            tabAll = new TabPage();
            tabSer = new TabPage();
            pnlSer1 = new Panel();
            pnlSer = new Panel();
            txtPlanRunName = new TextBox();
            ((System.ComponentModel.ISupportInitialize)dgvOverview).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvDetails).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvSummary).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvStatsAll).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvStatsSerial).BeginInit();
            grpFilters.SuspendLayout();
            grpMapping.SuspendLayout();
            grpStats.SuspendLayout();
            statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)split).BeginInit();
            split.Panel1.SuspendLayout();
            split.Panel2.SuspendLayout();
            split.SuspendLayout();
            containerTop.SuspendLayout();
            tabs.SuspendLayout();
            tabSummary.SuspendLayout();
            tabDetails.SuspendLayout();
            tabAll.SuspendLayout();
            tabSer.SuspendLayout();
            SuspendLayout();
            // 
            // btnOpenDb
            // 
            btnOpenDb.Location = new Point(12, 12);
            btnOpenDb.Name = "btnOpenDb";
            btnOpenDb.Size = new Size(120, 32);
            btnOpenDb.TabIndex = 0;
            btnOpenDb.Text = "Open DB...";
            btnOpenDb.Click += btnOpenDb_Click;
            // 
            // txtDbPath
            // 
            txtDbPath.Location = new Point(138, 14);
            txtDbPath.Name = "txtDbPath";
            txtDbPath.ReadOnly = true;
            txtDbPath.Size = new Size(500, 23);
            txtDbPath.TabIndex = 1;
            // 
            // btnLoad
            // 
            btnLoad.Location = new Point(650, 12);
            btnLoad.Name = "btnLoad";
            btnLoad.Size = new Size(150, 32);
            btnLoad.TabIndex = 2;
            btnLoad.Text = "Load Overview";
            btnLoad.Click += btnLoad_Click;
            // 
            // dgvOverview
            // 
            dgvOverview.ColumnHeadersHeight = 29;
            dgvOverview.Dock = DockStyle.Left;
            dgvOverview.Location = new Point(0, 0);
            dgvOverview.Name = "dgvOverview";
            dgvOverview.ReadOnly = true;
            dgvOverview.RowHeadersWidth = 51;
            dgvOverview.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvOverview.Size = new Size(888, 207);
            dgvOverview.TabIndex = 0;
            dgvOverview.SelectionChanged += dgvOverview_SelectionChanged;
            // 
            // dgvDetails
            // 
            dgvDetails.Anchor = AnchorStyles.Left;
            dgvDetails.ColumnHeadersHeight = 29;
            dgvDetails.Location = new Point(0, 39);
            dgvDetails.MaximumSize = new Size(2048, 475);
            dgvDetails.Name = "dgvDetails";
            dgvDetails.ReadOnly = true;
            dgvDetails.RowHeadersWidth = 51;
            dgvDetails.Size = new Size(1327, 385);
            dgvDetails.TabIndex = 0;
            // 
            // dgvSummary
            // 
            dgvSummary.Anchor = AnchorStyles.Left;
            dgvSummary.ColumnHeadersHeight = 29;
            dgvSummary.Location = new Point(0, 0);
            dgvSummary.Name = "dgvSummary";
            dgvSummary.ReadOnly = true;
            dgvSummary.RowHeadersWidth = 51;
            dgvSummary.Size = new Size(888, 424);
            dgvSummary.TabIndex = 0;
            // 
            // dgvStatsAll
            // 
            dgvStatsAll.Anchor = AnchorStyles.Left;
            dgvStatsAll.ColumnHeadersHeight = 29;
            dgvStatsAll.Location = new Point(0, 39);
            dgvStatsAll.MaximumSize = new Size(2048, 475);
            dgvStatsAll.Name = "dgvStatsAll";
            dgvStatsAll.ReadOnly = true;
            dgvStatsAll.RowHeadersWidth = 51;
            dgvStatsAll.Size = new Size(1327, 385);
            dgvStatsAll.TabIndex = 0;
            // 
            // dgvStatsSerial
            // 
            dgvStatsSerial.Anchor = AnchorStyles.Left;
            dgvStatsSerial.ColumnHeadersHeight = 29;
            dgvStatsSerial.Location = new Point(0, 39);
            dgvStatsSerial.Name = "dgvStatsSerial";
            dgvStatsSerial.ReadOnly = true;
            dgvStatsSerial.RowHeadersWidth = 51;
            dgvStatsSerial.Size = new Size(1327, 405);
            dgvStatsSerial.TabIndex = 0;
            // 
            // cboVerdict
            // 
            cboVerdict.DropDownStyle = ComboBoxStyle.DropDownList;
            cboVerdict.Location = new Point(10, 47);
            cboVerdict.Name = "cboVerdict";
            cboVerdict.Size = new Size(100, 23);
            cboVerdict.TabIndex = 1;
            // 
            // cboStatus
            // 
            cboStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            cboStatus.Location = new Point(731, 45);
            cboStatus.Name = "cboStatus";
            cboStatus.Size = new Size(100, 23);
            cboStatus.TabIndex = 3;
            // 
            // dtFrom
            // 
            dtFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dtFrom.Format = DateTimePickerFormat.Custom;
            dtFrom.Location = new Point(230, 47);
            dtFrom.Name = "dtFrom";
            dtFrom.Size = new Size(180, 23);
            dtFrom.TabIndex = 5;
            // 
            // dtTo
            // 
            dtTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dtTo.Format = DateTimePickerFormat.Custom;
            dtTo.Location = new Point(420, 47);
            dtTo.Name = "dtTo";
            dtTo.Size = new Size(180, 23);
            dtTo.TabIndex = 8;
            // 
            // cmbMapPlanRunNumber
            // 
            cmbMapPlanRunNumber.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbMapPlanRunNumber.Location = new Point(620, -2);
            cmbMapPlanRunNumber.Name = "cmbMapPlanRunNumber";
            cmbMapPlanRunNumber.Size = new Size(90, 23);
            cmbMapPlanRunNumber.TabIndex = 1;
            cmbMapPlanRunNumber.Visible = false;
            // 
            // btnExportOverview
            // 
            btnExportOverview.Location = new Point(815, 25);
            btnExportOverview.Name = "btnExportOverview";
            btnExportOverview.Size = new Size(150, 30);
            btnExportOverview.TabIndex = 2;
            btnExportOverview.Text = "Export Overview...";
            btnExportOverview.Click += btnExportOverview_Click;
            // 
            // btnExportDetails
            // 
            btnExportDetails.Location = new Point(267, 3);
            btnExportDetails.Name = "btnExportDetails";
            btnExportDetails.Size = new Size(600, 30);
            btnExportDetails.TabIndex = 1;
            btnExportDetails.Text = "Export...";
            btnExportDetails.Click += btnExportDetails_Click;
            // 
            // btnExportStatsAll
            // 
            btnExportStatsAll.Location = new Point(267, 3);
            btnExportStatsAll.Name = "btnExportStatsAll";
            btnExportStatsAll.Size = new Size(600, 30);
            btnExportStatsAll.TabIndex = 1;
            btnExportStatsAll.Text = "Export...";
            btnExportStatsAll.Click += btnExportStatsAll_Click;
            // 
            // btnExportStatsSerial
            // 
            btnExportStatsSerial.Location = new Point(267, 3);
            btnExportStatsSerial.Name = "btnExportStatsSerial";
            btnExportStatsSerial.Size = new Size(600, 30);
            btnExportStatsSerial.TabIndex = 0;
            btnExportStatsSerial.Text = "Export...";
            btnExportStatsSerial.Click += btnExportStatsSerial_Click;
            // 
            // btnApplyFilters
            // 
            btnApplyFilters.Location = new Point(610, 45);
            btnApplyFilters.Name = "btnApplyFilters";
            btnApplyFilters.Size = new Size(100, 30);
            btnApplyFilters.TabIndex = 9;
            btnApplyFilters.Text = "Apply";
            btnApplyFilters.Click += btnApplyFilters_Click;
            // 
            // chkWithSerialNumber
            // 
            chkWithSerialNumber.Checked = true;
            chkWithSerialNumber.CheckState = CheckState.Checked;
            chkWithSerialNumber.Location = new Point(116, 44);
            chkWithSerialNumber.Name = "chkWithSerialNumber";
            chkWithSerialNumber.Size = new Size(100, 30);
            chkWithSerialNumber.TabIndex = 7;
            chkWithSerialNumber.Text = "With Serial#";
            // 
            // btnGetDetails
            // 
            btnGetDetails.Location = new Point(12, 20);
            btnGetDetails.Name = "btnGetDetails";
            btnGetDetails.Size = new Size(150, 30);
            btnGetDetails.TabIndex = 1;
            btnGetDetails.Text = "Load Details";
            btnGetDetails.Click += btnGetDetails_Click;
            // 
            // btnLoadStatsAll
            // 
            btnLoadStatsAll.Location = new Point(172, 20);
            btnLoadStatsAll.Name = "btnLoadStatsAll";
            btnLoadStatsAll.Size = new Size(150, 30);
            btnLoadStatsAll.TabIndex = 0;
            btnLoadStatsAll.Text = "Load Stats (All)";
            btnLoadStatsAll.Click += btnLoadStatsAll_Click;
            // 
            // btnLoadStatsSerial
            // 
            btnLoadStatsSerial.Location = new Point(354, 20);
            btnLoadStatsSerial.Name = "btnLoadStatsSerial";
            btnLoadStatsSerial.Size = new Size(150, 30);
            btnLoadStatsSerial.TabIndex = 2;
            btnLoadStatsSerial.Text = "Load Stats (Serial#)";
            btnLoadStatsSerial.Click += btnLoadStatsSerial_Click;
            // 
            // lblStatusFilter
            // 
            lblStatusFilter.AutoSize = true;
            lblStatusFilter.Location = new Point(731, 21);
            lblStatusFilter.Name = "lblStatusFilter";
            lblStatusFilter.Size = new Size(42, 15);
            lblStatusFilter.TabIndex = 2;
            lblStatusFilter.Text = "Status:";
            // 
            // grpFilters
            // 
            grpFilters.Controls.Add(lblVerdict);
            grpFilters.Controls.Add(cboVerdict);
            grpFilters.Controls.Add(lblStatusFilter);
            grpFilters.Controls.Add(cboStatus);
            grpFilters.Controls.Add(lblFrom);
            grpFilters.Controls.Add(dtFrom);
            grpFilters.Controls.Add(lblTo);
            grpFilters.Controls.Add(chkWithSerialNumber);
            grpFilters.Controls.Add(dtTo);
            grpFilters.Controls.Add(btnApplyFilters);
            grpFilters.Location = new Point(12, 52);
            grpFilters.Name = "grpFilters";
            grpFilters.Size = new Size(988, 80);
            grpFilters.TabIndex = 3;
            grpFilters.TabStop = false;
            grpFilters.Text = "Filters";
            // 
            // lblVerdict
            // 
            lblVerdict.Location = new Point(10, 23);
            lblVerdict.Name = "lblVerdict";
            lblVerdict.Size = new Size(60, 18);
            lblVerdict.TabIndex = 0;
            lblVerdict.Text = "Verdict";
            // 
            // lblFrom
            // 
            lblFrom.Location = new Point(230, 10);
            lblFrom.Name = "lblFrom";
            lblFrom.Size = new Size(60, 20);
            lblFrom.TabIndex = 4;
            // 
            // lblTo
            // 
            lblTo.Location = new Point(420, 10);
            lblTo.Name = "lblTo";
            lblTo.Size = new Size(60, 20);
            lblTo.TabIndex = 6;
            // 
            // grpMapping
            // 
            grpMapping.Controls.Add(lblMap);
            grpMapping.Controls.Add(cmbMapPlanRunNumber);
            grpMapping.Location = new Point(12, 138);
            grpMapping.Name = "grpMapping";
            grpMapping.Size = new Size(988, 27);
            grpMapping.TabIndex = 4;
            grpMapping.TabStop = false;
            grpMapping.Text = "Column Mapping";
            grpMapping.Visible = false;
            // 
            // lblMap
            // 
            lblMap.Location = new Point(10, 25);
            lblMap.Name = "lblMap";
            lblMap.Size = new Size(180, 20);
            lblMap.TabIndex = 0;
            lblMap.Visible = false;
            // 
            // grpStats
            // 
            grpStats.Controls.Add(lblSerial);
            grpStats.Controls.Add(btnLoadStatsAll);
            grpStats.Controls.Add(txtSerial);
            grpStats.Controls.Add(btnLoadStatsSerial);
            grpStats.Controls.Add(btnGetDetails);
            grpStats.Controls.Add(btnExportOverview);
            grpStats.Location = new Point(12, 138);
            grpStats.Name = "grpStats";
            grpStats.Size = new Size(988, 67);
            grpStats.TabIndex = 5;
            grpStats.TabStop = false;
            grpStats.Text = "Statistics";
            // 
            // lblSerial
            // 
            lblSerial.AutoSize = true;
            lblSerial.Location = new Point(576, 28);
            lblSerial.Name = "lblSerial";
            lblSerial.Size = new Size(38, 15);
            lblSerial.TabIndex = 0;
            lblSerial.Text = "Serial:";
            // 
            // txtSerial
            // 
            txtSerial.Location = new Point(623, 25);
            txtSerial.Name = "txtSerial";
            txtSerial.Size = new Size(150, 23);
            txtSerial.TabIndex = 1;
            // 
            // statusStrip1
            // 
            statusStrip1.ImageScalingSize = new Size(20, 20);
            statusStrip1.Items.AddRange(new ToolStripItem[] { lblStatus });
            statusStrip1.Location = new Point(0, 931);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(1362, 22);
            statusStrip1.TabIndex = 7;
            // 
            // lblStatus
            // 
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(42, 17);
            lblStatus.Text = "Ready.";
            // 
            // split
            // 
            split.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            split.Location = new Point(12, 211);
            split.Name = "split";
            split.Orientation = Orientation.Horizontal;
            // 
            // split.Panel1
            // 
            split.Panel1.Controls.Add(containerTop);
            // 
            // split.Panel2
            // 
            split.Panel2.BackColor = SystemColors.ControlLight;
            split.Panel2.Controls.Add(tabs);
            split.Size = new Size(1372, 687);
            split.SplitterDistance = 208;
            split.TabIndex = 6;
            // 
            // containerTop
            // 
            containerTop.BackColor = SystemColors.ControlLight;
            containerTop.Controls.Add(dgvOverview);
            containerTop.Location = new Point(0, 0);
            containerTop.Name = "containerTop";
            containerTop.Size = new Size(907, 207);
            containerTop.TabIndex = 0;
            // 
            // tabs
            // 
            tabs.Anchor = AnchorStyles.Left;
            tabs.Controls.Add(tabSummary);
            tabs.Controls.Add(tabDetails);
            tabs.Controls.Add(tabAll);
            tabs.Controls.Add(tabSer);
            tabs.Location = new Point(0, 27);
            tabs.MaximumSize = new Size(2048, 475);
            tabs.Name = "tabs";
            tabs.SelectedIndex = 0;
            tabs.Size = new Size(1338, 475);
            tabs.TabIndex = 0;
            // 
            // tabSummary
            // 
            tabSummary.BackColor = SystemColors.ControlLight;
            tabSummary.Controls.Add(dgvSummary);
            tabSummary.Location = new Point(4, 24);
            tabSummary.Name = "tabSummary";
            tabSummary.Size = new Size(1330, 447);
            tabSummary.TabIndex = 0;
            tabSummary.Text = "Summary";
            // 
            // tabDetails
            // 
            tabDetails.Controls.Add(dgvDetails);
            tabDetails.Controls.Add(btnExportDetails);
            tabDetails.Location = new Point(4, 24);
            tabDetails.Name = "tabDetails";
            tabDetails.Size = new Size(1330, 447);
            tabDetails.TabIndex = 1;
            tabDetails.Text = "Details";
            // 
            // tabAll
            // 
            tabAll.Controls.Add(dgvStatsAll);
            tabAll.Controls.Add(btnExportStatsAll);
            tabAll.Location = new Point(4, 24);
            tabAll.Name = "tabAll";
            tabAll.Size = new Size(1330, 447);
            tabAll.TabIndex = 2;
            tabAll.Text = "Stats (All)";
            // 
            // tabSer
            // 
            tabSer.Controls.Add(dgvStatsSerial);
            tabSer.Controls.Add(btnExportStatsSerial);
            tabSer.Controls.Add(pnlSer1);
            tabSer.Location = new Point(4, 24);
            tabSer.Name = "tabSer";
            tabSer.Size = new Size(1330, 447);
            tabSer.TabIndex = 3;
            tabSer.Text = "Stats (By Serial)";
            // 
            // pnlSer1
            // 
            pnlSer1.AutoScroll = true;
            pnlSer1.Location = new Point(0, 0);
            pnlSer1.Name = "pnlSer1";
            pnlSer1.Size = new Size(200, 100);
            pnlSer1.TabIndex = 1;
            // 
            // pnlSer
            // 
            pnlSer.AutoScroll = true;
            pnlSer.Location = new Point(0, 0);
            pnlSer.Name = "pnlSer";
            pnlSer.Size = new Size(200, 100);
            pnlSer.TabIndex = 0;
            // 
            // txtPlanRunName
            // 
            txtPlanRunName.Location = new Point(10, 60);
            txtPlanRunName.Name = "txtPlanRunName";
            txtPlanRunName.PlaceholderText = "PlanRun Name contains...";
            txtPlanRunName.Size = new Size(300, 23);
            txtPlanRunName.TabIndex = 8;
            // 
            // MainForm
            // 
            CancelButton = btnApplyFilters;
            ClientSize = new Size(1362, 953);
            Controls.Add(btnOpenDb);
            Controls.Add(txtDbPath);
            Controls.Add(btnLoad);
            Controls.Add(grpFilters);
            Controls.Add(grpMapping);
            Controls.Add(grpStats);
            Controls.Add(split);
            Controls.Add(statusStrip1);
            Controls.Add(txtPlanRunName);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Location = new Point(100, 50);
            Name = "MainForm";
            Text = "Test Run Viewer (SQLite)";
            ((System.ComponentModel.ISupportInitialize)dgvOverview).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvDetails).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvSummary).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvStatsAll).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvStatsSerial).EndInit();
            grpFilters.ResumeLayout(false);
            grpFilters.PerformLayout();
            grpMapping.ResumeLayout(false);
            grpStats.ResumeLayout(false);
            grpStats.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            split.Panel1.ResumeLayout(false);
            split.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)split).EndInit();
            split.ResumeLayout(false);
            containerTop.ResumeLayout(false);
            tabs.ResumeLayout(false);
            tabSummary.ResumeLayout(false);
            tabDetails.ResumeLayout(false);
            tabAll.ResumeLayout(false);
            tabSer.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();



        }
        private Label lblVerdict;
        private Label lblFrom;
        private Label lblTo;
        private Label lblMap;
        private SplitContainer split;
        private Panel containerTop;
        private TabControl tabs;
        private TabPage tabSummary;
        private TabPage tabDetails;
        private TabPage tabAll;
        private TabPage tabSer;
        private Panel pnlSer1;
        private Label lblSerial;
        private Panel pnlSer;
        private TextBox txtSerial;
    }
}
