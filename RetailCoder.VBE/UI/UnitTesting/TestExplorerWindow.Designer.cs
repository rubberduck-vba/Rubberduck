namespace Rubberduck.UI.UnitTesting
{
    partial class TestExplorerWindow
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.testProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.passedTestsLabel = new System.Windows.Forms.ToolStripLabel();
            this.failedTestsLabel = new System.Windows.Forms.ToolStripLabel();
            this.inconclusiveTestsLabel = new System.Windows.Forms.ToolStripLabel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.testOutputGridView = new System.Windows.Forms.DataGridView();
            this.refreshTestsButton = new System.Windows.Forms.ToolStripButton();
            this.RunButton = new System.Windows.Forms.ToolStripDropDownButton();
            this.runSelectedTestMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runAllTestsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.runNotRunTestsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runFailedTestsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runPassedTestsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.runLastRunMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addButton = new System.Windows.Forms.ToolStripDropDownButton();
            this.addTestModuleButton = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.addTestMethodButton = new System.Windows.Forms.ToolStripMenuItem();
            this.addExpectedErrorTestMethodButton = new System.Windows.Forms.ToolStripMenuItem();
            this.gotoSelectionButton = new System.Windows.Forms.ToolStripButton();
            this.toolStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.testOutputGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshTestsButton,
            this.toolStripSeparator1,
            this.RunButton,
            this.addButton,
            this.gotoSelectionButton,
            this.toolStripSeparator5,
            this.testProgressBar,
            this.passedTestsLabel,
            this.failedTestsLabel,
            this.inconclusiveTestsLabel});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(600, 25);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(6, 25);
            // 
            // testProgressBar
            // 
            this.testProgressBar.ForeColor = System.Drawing.Color.LimeGreen;
            this.testProgressBar.Name = "testProgressBar";
            this.testProgressBar.Size = new System.Drawing.Size(100, 22);
            this.testProgressBar.Step = 1;
            // 
            // passedTestsLabel
            // 
            this.passedTestsLabel.Name = "passedTestsLabel";
            this.passedTestsLabel.Size = new System.Drawing.Size(52, 22);
            this.passedTestsLabel.Text = "0 Passed";
            // 
            // failedTestsLabel
            // 
            this.failedTestsLabel.Name = "failedTestsLabel";
            this.failedTestsLabel.Size = new System.Drawing.Size(47, 22);
            this.failedTestsLabel.Text = "0 Failed";
            // 
            // inconclusiveTestsLabel
            // 
            this.inconclusiveTestsLabel.Name = "inconclusiveTestsLabel";
            this.inconclusiveTestsLabel.Size = new System.Drawing.Size(82, 22);
            this.inconclusiveTestsLabel.Text = "0 Inconclusive";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.testOutputGridView);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(600, 175);
            this.panel1.TabIndex = 2;
            // 
            // testOutputGridView
            // 
            this.testOutputGridView.AllowUserToAddRows = false;
            this.testOutputGridView.AllowUserToDeleteRows = false;
            this.testOutputGridView.AllowUserToOrderColumns = true;
            this.testOutputGridView.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.WhiteSmoke;
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            this.testOutputGridView.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.testOutputGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.testOutputGridView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.testOutputGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.testOutputGridView.DefaultCellStyle = dataGridViewCellStyle2;
            this.testOutputGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.testOutputGridView.Location = new System.Drawing.Point(0, 0);
            this.testOutputGridView.Name = "testOutputGridView";
            this.testOutputGridView.ReadOnly = true;
            this.testOutputGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.testOutputGridView.Size = new System.Drawing.Size(600, 175);
            this.testOutputGridView.TabIndex = 1;
            this.testOutputGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.GridCellDoubleClicked);
            // 
            // refreshTestsButton
            // 
            this.refreshTestsButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.refreshTestsButton.Image = global::Rubberduck.Properties.Resources.Refresh;
            this.refreshTestsButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.refreshTestsButton.Name = "refreshTestsButton";
            this.refreshTestsButton.Size = new System.Drawing.Size(23, 22);
            this.refreshTestsButton.Click += new System.EventHandler(this.RefreshTestsButtonClick);
            // 
            // RunButton
            // 
            this.RunButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.runSelectedTestMenuItem,
            this.runAllTestsMenuItem,
            this.toolStripSeparator3,
            this.runNotRunTestsMenuItem,
            this.runFailedTestsMenuItem,
            this.runPassedTestsMenuItem,
            this.toolStripSeparator4,
            this.runLastRunMenuItem});
            this.RunButton.Image = global::Rubberduck.Properties.Resources.Play;
            this.RunButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RunButton.Name = "RunButton";
            this.RunButton.Size = new System.Drawing.Size(57, 22);
            this.RunButton.Text = "&Run";
            // 
            // runSelectedTestMenuItem
            // 
            this.runSelectedTestMenuItem.Enabled = false;
            this.runSelectedTestMenuItem.Name = "runSelectedTestMenuItem";
            this.runSelectedTestMenuItem.Size = new System.Drawing.Size(230, 22);
            this.runSelectedTestMenuItem.Text = "&Selected Tests";
            // 
            // runAllTestsMenuItem
            // 
            this.runAllTestsMenuItem.Image = global::Rubberduck.Properties.Resources.AllLoadedTests_8644_24;
            this.runAllTestsMenuItem.Name = "runAllTestsMenuItem";
            this.runAllTestsMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.A)));
            this.runAllTestsMenuItem.Size = new System.Drawing.Size(230, 22);
            this.runAllTestsMenuItem.Text = "&All Tests";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(227, 6);
            // 
            // runNotRunTestsMenuItem
            // 
            this.runNotRunTestsMenuItem.Name = "runNotRunTestsMenuItem";
            this.runNotRunTestsMenuItem.Size = new System.Drawing.Size(230, 22);
            this.runNotRunTestsMenuItem.Text = "&Not Run Tests";
            // 
            // runFailedTestsMenuItem
            // 
            this.runFailedTestsMenuItem.Enabled = false;
            this.runFailedTestsMenuItem.Name = "runFailedTestsMenuItem";
            this.runFailedTestsMenuItem.Size = new System.Drawing.Size(230, 22);
            this.runFailedTestsMenuItem.Text = "&Failed Tests";
            // 
            // runPassedTestsMenuItem
            // 
            this.runPassedTestsMenuItem.Enabled = false;
            this.runPassedTestsMenuItem.Name = "runPassedTestsMenuItem";
            this.runPassedTestsMenuItem.Size = new System.Drawing.Size(230, 22);
            this.runPassedTestsMenuItem.Text = "&Passed Tests";
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(227, 6);
            // 
            // runLastRunMenuItem
            // 
            this.runLastRunMenuItem.Enabled = false;
            this.runLastRunMenuItem.Name = "runLastRunMenuItem";
            this.runLastRunMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.L)));
            this.runLastRunMenuItem.Size = new System.Drawing.Size(230, 22);
            this.runLastRunMenuItem.Text = "Repeat &Last Run";
            // 
            // addButton
            // 
            this.addButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addTestModuleButton,
            this.toolStripSeparator2,
            this.addTestMethodButton,
            this.addExpectedErrorTestMethodButton});
            this.addButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(42, 22);
            this.addButton.Text = "&Add";
            // 
            // addTestModuleButton
            // 
            this.addTestModuleButton.Image = global::Rubberduck.Properties.Resources.AddModule_368_321;
            this.addTestModuleButton.Name = "addTestModuleButton";
            this.addTestModuleButton.Size = new System.Drawing.Size(227, 22);
            this.addTestModuleButton.Text = "Test &Module";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(224, 6);
            // 
            // addTestMethodButton
            // 
            this.addTestMethodButton.ImageTransparentColor = System.Drawing.Color.Fuchsia;
            this.addTestMethodButton.Name = "addTestMethodButton";
            this.addTestMethodButton.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.T)));
            this.addTestMethodButton.Size = new System.Drawing.Size(227, 22);
            this.addTestMethodButton.Text = "&Test Method";
            // 
            // addExpectedErrorTestMethodButton
            // 
            this.addExpectedErrorTestMethodButton.Name = "addExpectedErrorTestMethodButton";
            this.addExpectedErrorTestMethodButton.Size = new System.Drawing.Size(227, 22);
            this.addExpectedErrorTestMethodButton.Text = "Test Method (Expected &Error)";
            // 
            // gotoSelectionButton
            // 
            this.gotoSelectionButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gotoSelectionButton.Image = global::Rubberduck.Properties.Resources.GoLtrHS;
            this.gotoSelectionButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.gotoSelectionButton.Name = "gotoSelectionButton";
            this.gotoSelectionButton.Size = new System.Drawing.Size(23, 22);
            // 
            // TestExplorerWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolStrip1);
            this.MinimumSize = new System.Drawing.Size(600, 200);
            this.Name = "TestExplorerWindow";
            this.Size = new System.Drawing.Size(600, 200);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.testOutputGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton refreshTestsButton;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView testOutputGridView;
        private System.Windows.Forms.ToolStripDropDownButton RunButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem runSelectedTestMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runAllTestsMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runFailedTestsMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runNotRunTestsMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runPassedTestsMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runLastRunMenuItem;
        private System.Windows.Forms.ToolStripDropDownButton addButton;
        private System.Windows.Forms.ToolStripMenuItem addTestModuleButton;
        private System.Windows.Forms.ToolStripMenuItem addTestMethodButton;
        private System.Windows.Forms.ToolStripMenuItem addExpectedErrorTestMethodButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripButton gotoSelectionButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripProgressBar testProgressBar;
        private System.Windows.Forms.ToolStripLabel passedTestsLabel;
        private System.Windows.Forms.ToolStripLabel failedTestsLabel;
        private System.Windows.Forms.ToolStripLabel inconclusiveTestsLabel;

    }
}