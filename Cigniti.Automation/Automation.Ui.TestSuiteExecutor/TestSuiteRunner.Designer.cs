namespace Maxim.Automation.Ui.TestSuiteExecutor

{
    partial class form_TestSuiteRunner
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel_Form = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.tableLayout_Filter = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.lstbox_Criteria = new System.Windows.Forms.ListBox();
            this.lstbox_SubModule = new System.Windows.Forms.ListBox();
            this.lstbox_Module = new System.Windows.Forms.ListBox();
            this.lbl_FilterByCategory = new System.Windows.Forms.Label();
            this.lbl_FilterBySubModule = new System.Windows.Forms.Label();
            this.lbl_FilterByModule = new System.Windows.Forms.Label();
            this.chkbox_SelectAll = new System.Windows.Forms.CheckBox();
            this.btn_ExecuteTests = new System.Windows.Forms.Button();
            this.btn_ClearAllFilters = new System.Windows.Forms.Button();
            this.lstbox_UserStory = new System.Windows.Forms.ListBox();
            this.btnSettings = new System.Windows.Forms.Button();
            this.tableLayout_ExecSummary = new System.Windows.Forms.TableLayoutPanel();
            this.lbl_TotalTCsPassPerResult = new System.Windows.Forms.Label();
            this.lbl_TotalTCsFailedResult = new System.Windows.Forms.Label();
            this.lbl_TotalTCsPassedResult = new System.Windows.Forms.Label();
            this.lbl_TotalTCsPassPercentage = new System.Windows.Forms.Label();
            this.lbl_TotalTCsFailed = new System.Windows.Forms.Label();
            this.lbl_TotalTCsPassed = new System.Windows.Forms.Label();
            this.lbl_TotalTCsExecuted = new System.Windows.Forms.Label();
            this.lbl_ExecutionSummary = new System.Windows.Forms.Label();
            this.lbl_TotalTCs = new System.Windows.Forms.Label();
            this.lbl_TotalTCsResult = new System.Windows.Forms.Label();
            this.lbl_TotalTCsExecutedResult = new System.Windows.Forms.Label();
            this.tableLayout_Panel = new System.Windows.Forms.TableLayoutPanel();
            this.pic_ClientLogo = new System.Windows.Forms.PictureBox();
            this.pic_CompanyLogo = new System.Windows.Forms.PictureBox();
            this.lbl_Header = new System.Windows.Forms.Label();
            this.panel_Form.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tableLayout_Filter.SuspendLayout();
            this.tableLayout_ExecSummary.SuspendLayout();
            this.tableLayout_Panel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ClientLogo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_CompanyLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // panel_Form
            // 
            this.panel_Form.AutoSize = true;
            this.panel_Form.BackColor = System.Drawing.Color.White;
            this.panel_Form.Controls.Add(this.dataGridView1);
            this.panel_Form.Controls.Add(this.tableLayout_Filter);
            this.panel_Form.Controls.Add(this.tableLayout_ExecSummary);
            this.panel_Form.Controls.Add(this.tableLayout_Panel);
            this.panel_Form.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel_Form.Location = new System.Drawing.Point(0, 0);
            this.panel_Form.Name = "panel_Form";
            this.panel_Form.Size = new System.Drawing.Size(1250, 494);
            this.panel_Form.TabIndex = 0;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 308);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1250, 186);
            this.dataGridView1.TabIndex = 3;
            // 
            // tableLayout_Filter
            // 
            this.tableLayout_Filter.ColumnCount = 4;
            this.tableLayout_Filter.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.16917F));
            this.tableLayout_Filter.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.65385F));
            this.tableLayout_Filter.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 36.21795F));
            this.tableLayout_Filter.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 337F));
            this.tableLayout_Filter.Controls.Add(this.label1, 3, 0);
            this.tableLayout_Filter.Controls.Add(this.lstbox_Criteria, 2, 1);
            this.tableLayout_Filter.Controls.Add(this.lstbox_SubModule, 1, 1);
            this.tableLayout_Filter.Controls.Add(this.lstbox_Module, 0, 1);
            this.tableLayout_Filter.Controls.Add(this.lbl_FilterByCategory, 0, 0);
            this.tableLayout_Filter.Controls.Add(this.lbl_FilterBySubModule, 0, 0);
            this.tableLayout_Filter.Controls.Add(this.lbl_FilterByModule, 0, 0);
            this.tableLayout_Filter.Controls.Add(this.chkbox_SelectAll, 0, 2);
            this.tableLayout_Filter.Controls.Add(this.btn_ExecuteTests, 1, 2);
            this.tableLayout_Filter.Controls.Add(this.btn_ClearAllFilters, 2, 2);
            this.tableLayout_Filter.Controls.Add(this.lstbox_UserStory, 3, 1);
            this.tableLayout_Filter.Controls.Add(this.btnSettings, 3, 2);
            this.tableLayout_Filter.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayout_Filter.Location = new System.Drawing.Point(0, 132);
            this.tableLayout_Filter.Name = "tableLayout_Filter";
            this.tableLayout_Filter.RowCount = 3;
            this.tableLayout_Filter.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayout_Filter.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 75F));
            this.tableLayout_Filter.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayout_Filter.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayout_Filter.Size = new System.Drawing.Size(1250, 176);
            this.tableLayout_Filter.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.LightGreen;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(915, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(332, 17);
            this.label1.TabIndex = 19;
            this.label1.Text = "Filter by - User Story";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // lstbox_Criteria
            // 
            this.lstbox_Criteria.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstbox_Criteria.FormattingEnabled = true;
            this.lstbox_Criteria.Location = new System.Drawing.Point(585, 20);
            this.lstbox_Criteria.Name = "lstbox_Criteria";
            this.lstbox_Criteria.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lstbox_Criteria.Size = new System.Drawing.Size(324, 126);
            this.lstbox_Criteria.TabIndex = 14;
            this.lstbox_Criteria.SelectedIndexChanged += new System.EventHandler(this.ListboxCriteriaSelectedIndexChanged);
            // 
            // lstbox_SubModule
            // 
            this.lstbox_SubModule.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstbox_SubModule.FormattingEnabled = true;
            this.lstbox_SubModule.Location = new System.Drawing.Point(278, 20);
            this.lstbox_SubModule.Name = "lstbox_SubModule";
            this.lstbox_SubModule.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lstbox_SubModule.Size = new System.Drawing.Size(301, 126);
            this.lstbox_SubModule.TabIndex = 13;
            this.lstbox_SubModule.SelectedIndexChanged += new System.EventHandler(this.ListboxSubModuleSelectedIndexChanged);
            // 
            // lstbox_Module
            // 
            this.lstbox_Module.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstbox_Module.FormattingEnabled = true;
            this.lstbox_Module.Location = new System.Drawing.Point(3, 20);
            this.lstbox_Module.Name = "lstbox_Module";
            this.lstbox_Module.Size = new System.Drawing.Size(269, 126);
            this.lstbox_Module.TabIndex = 8;
            this.lstbox_Module.SelectedIndexChanged += new System.EventHandler(this.ListBoxModuleSelectedIndexChanged);
            // 
            // lbl_FilterByCategory
            // 
            this.lbl_FilterByCategory.AutoSize = true;
            this.lbl_FilterByCategory.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_FilterByCategory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_FilterByCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_FilterByCategory.Location = new System.Drawing.Point(3, 0);
            this.lbl_FilterByCategory.Name = "lbl_FilterByCategory";
            this.lbl_FilterByCategory.Size = new System.Drawing.Size(269, 17);
            this.lbl_FilterByCategory.TabIndex = 6;
            this.lbl_FilterByCategory.Text = "Filter by - Module";
            this.lbl_FilterByCategory.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lbl_FilterByCategory.Click += new System.EventHandler(this.lbl_FilterByCategory_Click);
            // 
            // lbl_FilterBySubModule
            // 
            this.lbl_FilterBySubModule.AutoSize = true;
            this.lbl_FilterBySubModule.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_FilterBySubModule.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_FilterBySubModule.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_FilterBySubModule.Location = new System.Drawing.Point(585, 0);
            this.lbl_FilterBySubModule.Name = "lbl_FilterBySubModule";
            this.lbl_FilterBySubModule.Size = new System.Drawing.Size(324, 17);
            this.lbl_FilterBySubModule.TabIndex = 5;
            this.lbl_FilterBySubModule.Text = "Filter by - Category";
            this.lbl_FilterBySubModule.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_FilterByModule
            // 
            this.lbl_FilterByModule.AutoSize = true;
            this.lbl_FilterByModule.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_FilterByModule.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_FilterByModule.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_FilterByModule.Location = new System.Drawing.Point(278, 0);
            this.lbl_FilterByModule.Name = "lbl_FilterByModule";
            this.lbl_FilterByModule.Size = new System.Drawing.Size(301, 17);
            this.lbl_FilterByModule.TabIndex = 4;
            this.lbl_FilterByModule.Text = "Filter by - Sub Module";
            this.lbl_FilterByModule.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // chkbox_SelectAll
            // 
            this.chkbox_SelectAll.AutoSize = true;
            this.chkbox_SelectAll.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chkbox_SelectAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkbox_SelectAll.Location = new System.Drawing.Point(3, 152);
            this.chkbox_SelectAll.Name = "chkbox_SelectAll";
            this.chkbox_SelectAll.Size = new System.Drawing.Size(269, 21);
            this.chkbox_SelectAll.TabIndex = 15;
            this.chkbox_SelectAll.Text = "Select All";
            this.chkbox_SelectAll.UseVisualStyleBackColor = true;
            this.chkbox_SelectAll.CheckedChanged += new System.EventHandler(this.CheckBox_SelectAll_CheckedChanged);
            // 
            // btn_ExecuteTests
            // 
            this.btn_ExecuteTests.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_ExecuteTests.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_ExecuteTests.Location = new System.Drawing.Point(278, 152);
            this.btn_ExecuteTests.Name = "btn_ExecuteTests";
            this.btn_ExecuteTests.Size = new System.Drawing.Size(301, 21);
            this.btn_ExecuteTests.TabIndex = 16;
            this.btn_ExecuteTests.Text = "Execute Tests";
            this.btn_ExecuteTests.UseVisualStyleBackColor = true;
            this.btn_ExecuteTests.Click += new System.EventHandler(this.btn_ExecuteTests_Click);
            // 
            // btn_ClearAllFilters
            // 
            this.btn_ClearAllFilters.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_ClearAllFilters.Location = new System.Drawing.Point(585, 152);
            this.btn_ClearAllFilters.Name = "btn_ClearAllFilters";
            this.btn_ClearAllFilters.Size = new System.Drawing.Size(324, 21);
            this.btn_ClearAllFilters.TabIndex = 17;
            this.btn_ClearAllFilters.Text = "Clear All";
            this.btn_ClearAllFilters.UseVisualStyleBackColor = true;
            this.btn_ClearAllFilters.Click += new System.EventHandler(this.btn_ClearAllFilters_Click);
            // 
            // lstbox_UserStory
            // 
            this.lstbox_UserStory.FormattingEnabled = true;
            this.lstbox_UserStory.Location = new System.Drawing.Point(915, 20);
            this.lstbox_UserStory.Name = "lstbox_UserStory";
            this.lstbox_UserStory.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lstbox_UserStory.Size = new System.Drawing.Size(322, 121);
            this.lstbox_UserStory.TabIndex = 18;
            this.lstbox_UserStory.SelectedIndexChanged += new System.EventHandler(this.ListboxUserStorySelectedIndexChanged);
            // 
            // btnSettings
            // 
            this.btnSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSettings.Location = new System.Drawing.Point(915, 152);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(328, 21);
            this.btnSettings.TabIndex = 20;
            this.btnSettings.Text = "Settings";
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.AppsetterClick);
            // 
            // tableLayout_ExecSummary
            // 
            this.tableLayout_ExecSummary.ColumnCount = 5;
            this.tableLayout_ExecSummary.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayout_ExecSummary.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayout_ExecSummary.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayout_ExecSummary.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayout_ExecSummary.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsPassPerResult, 4, 2);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsFailedResult, 3, 2);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsPassedResult, 2, 2);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsPassPercentage, 4, 1);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsFailed, 3, 1);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsPassed, 2, 1);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsExecuted, 1, 1);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_ExecutionSummary, 0, 0);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCs, 0, 1);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsResult, 0, 2);
            this.tableLayout_ExecSummary.Controls.Add(this.lbl_TotalTCsExecutedResult, 1, 2);
            this.tableLayout_ExecSummary.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayout_ExecSummary.Location = new System.Drawing.Point(0, 73);
            this.tableLayout_ExecSummary.Name = "tableLayout_ExecSummary";
            this.tableLayout_ExecSummary.RowCount = 3;
            this.tableLayout_ExecSummary.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.8983F));
            this.tableLayout_ExecSummary.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.8983F));
            this.tableLayout_ExecSummary.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 31.37255F));
            this.tableLayout_ExecSummary.Size = new System.Drawing.Size(1250, 59);
            this.tableLayout_ExecSummary.TabIndex = 1;
            // 
            // lbl_TotalTCsPassPerResult
            // 
            this.lbl_TotalTCsPassPerResult.AutoSize = true;
            this.lbl_TotalTCsPassPerResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsPassPerResult.Location = new System.Drawing.Point(1003, 40);
            this.lbl_TotalTCsPassPerResult.Name = "lbl_TotalTCsPassPerResult";
            this.lbl_TotalTCsPassPerResult.Size = new System.Drawing.Size(244, 19);
            this.lbl_TotalTCsPassPerResult.TabIndex = 10;
            this.lbl_TotalTCsPassPerResult.Text = "0";
            this.lbl_TotalTCsPassPerResult.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsFailedResult
            // 
            this.lbl_TotalTCsFailedResult.AutoSize = true;
            this.lbl_TotalTCsFailedResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsFailedResult.Location = new System.Drawing.Point(753, 40);
            this.lbl_TotalTCsFailedResult.Name = "lbl_TotalTCsFailedResult";
            this.lbl_TotalTCsFailedResult.Size = new System.Drawing.Size(244, 19);
            this.lbl_TotalTCsFailedResult.TabIndex = 9;
            this.lbl_TotalTCsFailedResult.Text = "0";
            this.lbl_TotalTCsFailedResult.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsPassedResult
            // 
            this.lbl_TotalTCsPassedResult.AutoSize = true;
            this.lbl_TotalTCsPassedResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsPassedResult.Location = new System.Drawing.Point(503, 40);
            this.lbl_TotalTCsPassedResult.Name = "lbl_TotalTCsPassedResult";
            this.lbl_TotalTCsPassedResult.Size = new System.Drawing.Size(244, 19);
            this.lbl_TotalTCsPassedResult.TabIndex = 8;
            this.lbl_TotalTCsPassedResult.Text = "0";
            this.lbl_TotalTCsPassedResult.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsPassPercentage
            // 
            this.lbl_TotalTCsPassPercentage.AutoSize = true;
            this.lbl_TotalTCsPassPercentage.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_TotalTCsPassPercentage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsPassPercentage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_TotalTCsPassPercentage.Location = new System.Drawing.Point(1003, 20);
            this.lbl_TotalTCsPassPercentage.Name = "lbl_TotalTCsPassPercentage";
            this.lbl_TotalTCsPassPercentage.Size = new System.Drawing.Size(244, 20);
            this.lbl_TotalTCsPassPercentage.TabIndex = 5;
            this.lbl_TotalTCsPassPercentage.Text = "Pass %";
            this.lbl_TotalTCsPassPercentage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsFailed
            // 
            this.lbl_TotalTCsFailed.AutoSize = true;
            this.lbl_TotalTCsFailed.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_TotalTCsFailed.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsFailed.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_TotalTCsFailed.Location = new System.Drawing.Point(753, 20);
            this.lbl_TotalTCsFailed.Name = "lbl_TotalTCsFailed";
            this.lbl_TotalTCsFailed.Size = new System.Drawing.Size(244, 20);
            this.lbl_TotalTCsFailed.TabIndex = 4;
            this.lbl_TotalTCsFailed.Text = "Failed";
            this.lbl_TotalTCsFailed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsPassed
            // 
            this.lbl_TotalTCsPassed.AutoSize = true;
            this.lbl_TotalTCsPassed.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_TotalTCsPassed.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsPassed.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_TotalTCsPassed.Location = new System.Drawing.Point(503, 20);
            this.lbl_TotalTCsPassed.Name = "lbl_TotalTCsPassed";
            this.lbl_TotalTCsPassed.Size = new System.Drawing.Size(244, 20);
            this.lbl_TotalTCsPassed.TabIndex = 3;
            this.lbl_TotalTCsPassed.Text = "Passed";
            this.lbl_TotalTCsPassed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsExecuted
            // 
            this.lbl_TotalTCsExecuted.AutoSize = true;
            this.lbl_TotalTCsExecuted.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_TotalTCsExecuted.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsExecuted.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_TotalTCsExecuted.Location = new System.Drawing.Point(253, 20);
            this.lbl_TotalTCsExecuted.Name = "lbl_TotalTCsExecuted";
            this.lbl_TotalTCsExecuted.Size = new System.Drawing.Size(244, 20);
            this.lbl_TotalTCsExecuted.TabIndex = 2;
            this.lbl_TotalTCsExecuted.Text = "Executed";
            this.lbl_TotalTCsExecuted.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_ExecutionSummary
            // 
            this.lbl_ExecutionSummary.AutoSize = true;
            this.lbl_ExecutionSummary.BackColor = System.Drawing.Color.LightBlue;
            this.tableLayout_ExecSummary.SetColumnSpan(this.lbl_ExecutionSummary, 5);
            this.lbl_ExecutionSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_ExecutionSummary.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_ExecutionSummary.Location = new System.Drawing.Point(3, 0);
            this.lbl_ExecutionSummary.Name = "lbl_ExecutionSummary";
            this.lbl_ExecutionSummary.Size = new System.Drawing.Size(1244, 20);
            this.lbl_ExecutionSummary.TabIndex = 0;
            this.lbl_ExecutionSummary.Text = "Execution Summary";
            this.lbl_ExecutionSummary.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCs
            // 
            this.lbl_TotalTCs.AutoSize = true;
            this.lbl_TotalTCs.BackColor = System.Drawing.Color.LightGreen;
            this.lbl_TotalTCs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCs.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_TotalTCs.Location = new System.Drawing.Point(3, 20);
            this.lbl_TotalTCs.Name = "lbl_TotalTCs";
            this.lbl_TotalTCs.Size = new System.Drawing.Size(244, 20);
            this.lbl_TotalTCs.TabIndex = 1;
            this.lbl_TotalTCs.Text = "Total Test Cases";
            this.lbl_TotalTCs.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsResult
            // 
            this.lbl_TotalTCsResult.AutoSize = true;
            this.lbl_TotalTCsResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsResult.Location = new System.Drawing.Point(3, 40);
            this.lbl_TotalTCsResult.Name = "lbl_TotalTCsResult";
            this.lbl_TotalTCsResult.Size = new System.Drawing.Size(244, 19);
            this.lbl_TotalTCsResult.TabIndex = 6;
            this.lbl_TotalTCsResult.Text = "0";
            this.lbl_TotalTCsResult.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_TotalTCsExecutedResult
            // 
            this.lbl_TotalTCsExecutedResult.AutoSize = true;
            this.lbl_TotalTCsExecutedResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_TotalTCsExecutedResult.Location = new System.Drawing.Point(253, 40);
            this.lbl_TotalTCsExecutedResult.Name = "lbl_TotalTCsExecutedResult";
            this.lbl_TotalTCsExecutedResult.Size = new System.Drawing.Size(244, 19);
            this.lbl_TotalTCsExecutedResult.TabIndex = 11;
            this.lbl_TotalTCsExecutedResult.Text = "0";
            this.lbl_TotalTCsExecutedResult.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayout_Panel
            // 
            this.tableLayout_Panel.ColumnCount = 3;
            this.tableLayout_Panel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayout_Panel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40F));
            this.tableLayout_Panel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayout_Panel.Controls.Add(this.pic_ClientLogo, 0, 0);
            this.tableLayout_Panel.Controls.Add(this.pic_CompanyLogo, 2, 0);
            this.tableLayout_Panel.Controls.Add(this.lbl_Header, 1, 0);
            this.tableLayout_Panel.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayout_Panel.Location = new System.Drawing.Point(0, 0);
            this.tableLayout_Panel.Name = "tableLayout_Panel";
            this.tableLayout_Panel.RowCount = 1;
            this.tableLayout_Panel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayout_Panel.Size = new System.Drawing.Size(1250, 73);
            this.tableLayout_Panel.TabIndex = 0;
            // 
            // pic_ClientLogo
            // 
            this.pic_ClientLogo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pic_ClientLogo.Image = global::Automation.Ui.TestSuiteExecutor.Properties.Resources.Client_Logo;
            this.pic_ClientLogo.Location = new System.Drawing.Point(3, 3);
            this.pic_ClientLogo.Name = "pic_ClientLogo";
            this.pic_ClientLogo.Size = new System.Drawing.Size(369, 67);
            this.pic_ClientLogo.TabIndex = 0;
            this.pic_ClientLogo.TabStop = false;
            // 
            // pic_CompanyLogo
            // 
            this.pic_CompanyLogo.Dock = System.Windows.Forms.DockStyle.Right;
            this.pic_CompanyLogo.Image = global::Automation.Ui.TestSuiteExecutor.Properties.Resources.Company_logo;
            this.pic_CompanyLogo.Location = new System.Drawing.Point(1102, 3);
            this.pic_CompanyLogo.Name = "pic_CompanyLogo";
            this.pic_CompanyLogo.Size = new System.Drawing.Size(145, 67);
            this.pic_CompanyLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pic_CompanyLogo.TabIndex = 1;
            this.pic_CompanyLogo.TabStop = false;
            // 
            // lbl_Header
            // 
            this.lbl_Header.AutoSize = true;
            this.lbl_Header.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_Header.Font = new System.Drawing.Font("Microsoft Sans Serif", 35F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Header.Location = new System.Drawing.Point(378, 0);
            this.lbl_Header.Name = "lbl_Header";
            this.lbl_Header.Size = new System.Drawing.Size(494, 73);
            this.lbl_Header.TabIndex = 2;
            this.lbl_Header.Text = "Test Suite Executor";
            this.lbl_Header.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // form_TestSuiteRunner
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1250, 494);
            this.Controls.Add(this.panel_Form);
            this.Name = "form_TestSuiteRunner";
            this.Text = "TestSuiteRunner";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.form_TestSuiteRunner_Load);
            this.panel_Form.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tableLayout_Filter.ResumeLayout(false);
            this.tableLayout_Filter.PerformLayout();
            this.tableLayout_ExecSummary.ResumeLayout(false);
            this.tableLayout_ExecSummary.PerformLayout();
            this.tableLayout_Panel.ResumeLayout(false);
            this.tableLayout_Panel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ClientLogo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_CompanyLogo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Panel panel_Form;
        private System.Windows.Forms.TableLayoutPanel tableLayout_Panel;
        private System.Windows.Forms.PictureBox pic_ClientLogo;
        private System.Windows.Forms.PictureBox pic_CompanyLogo;
        private System.Windows.Forms.Label lbl_Header;
        private System.Windows.Forms.TableLayoutPanel tableLayout_ExecSummary;
        private System.Windows.Forms.Label lbl_ExecutionSummary;
        private System.Windows.Forms.Label lbl_TotalTCsPassPerResult;
        private System.Windows.Forms.Label lbl_TotalTCsFailedResult;
        private System.Windows.Forms.Label lbl_TotalTCsPassedResult;
        private System.Windows.Forms.Label lbl_TotalTCsResult;
        private System.Windows.Forms.Label lbl_TotalTCsPassPercentage;
        private System.Windows.Forms.Label lbl_TotalTCsFailed;
        private System.Windows.Forms.Label lbl_TotalTCsPassed;
        private System.Windows.Forms.Label lbl_TotalTCsExecuted;
        private System.Windows.Forms.Label lbl_TotalTCs;
        private System.Windows.Forms.TableLayoutPanel tableLayout_Filter;
        private System.Windows.Forms.ListBox lstbox_SubModule;
        private System.Windows.Forms.ListBox lstbox_Criteria;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ListBox lstbox_Module;
        private System.Windows.Forms.Label lbl_FilterByCategory;
        private System.Windows.Forms.Label lbl_FilterBySubModule;
        private System.Windows.Forms.Label lbl_FilterByModule;
        private System.Windows.Forms.CheckBox chkbox_SelectAll;
        private System.Windows.Forms.Button btn_ExecuteTests;
        private System.Windows.Forms.Label lbl_TotalTCsExecutedResult;
        private System.Windows.Forms.Button btn_ClearAllFilters;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstbox_UserStory;
        private System.Windows.Forms.Button btnSettings;
    }
}