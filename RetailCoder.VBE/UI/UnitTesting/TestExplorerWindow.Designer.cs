using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.UnitTesting
{
    partial class TestExplorerWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TestExplorerWindow));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.refreshTestsButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.runButton = new System.Windows.Forms.ToolStripDropDownButton();
            this.runAllTestsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runSelectedTestMenuItem = new System.Windows.Forms.ToolStripMenuItem();
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
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.testProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.TotalElapsedMilisecondsLabel = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.passedTestsLabel = new System.Windows.Forms.ToolStripLabel();
            this.failedTestsLabel = new System.Windows.Forms.ToolStripLabel();
            this.inconclusiveTestsLabel = new System.Windows.Forms.ToolStripLabel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.testOutputGridView = new System.Windows.Forms.DataGridView();
            this.toolStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.testOutputGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshTestsButton,
            this.toolStripSeparator1,
            this.runButton,
            this.addButton,
            this.gotoSelectionButton,
            this.toolStripSeparator5,
            this.testProgressBar,
            this.TotalElapsedMilisecondsLabel,
            this.toolStripSeparator6,
            this.passedTestsLabel,
            this.failedTestsLabel,
            this.inconclusiveTestsLabel});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(800, 30);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // refreshTestsButton
            // 
            this.refreshTestsButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.refreshTestsButton.Image = ((System.Drawing.Image)(resources.GetObject("refreshTestsButton.Image")));
            this.refreshTestsButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.refreshTestsButton.Name = "refreshTestsButton";
            this.refreshTestsButton.Size = new System.Drawing.Size(24, 27);
            this.refreshTestsButton.ToolTipText = "Refresh";
            this.refreshTestsButton.Click += new System.EventHandler(this.RefreshTestsButtonClick);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 30);
            // 
            // runButton
            // 
            this.runButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.runAllTestsMenuItem,
            this.runSelectedTestMenuItem,
            this.toolStripSeparator3,
            this.runNotRunTestsMenuItem,
            this.runFailedTestsMenuItem,
            this.runPassedTestsMenuItem,
            this.toolStripSeparator4,
            this.runLastRunMenuItem});
            this.runButton.Image = global::Rubberduck.Properties.Resources.control_000_small;
            this.runButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(68, 27);
            this.runButton.Text = "&Run";
            // 
            // runAllTestsMenuItem
            // 
            this.runAllTestsMenuItem.Image = global::Rubberduck.Properties.Resources.flask_arrow;
            this.runAllTestsMenuItem.ImageTransparentColor = System.Drawing.Color.White;
            this.runAllTestsMenuItem.Name = "runAllTestsMenuItem";
            this.runAllTestsMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.A)));
            this.runAllTestsMenuItem.Size = new System.Drawing.Size(279, 26);
            this.runAllTestsMenuItem.Text = "&All Tests";
            // 
            // runSelectedTestMenuItem
            // 
            this.runSelectedTestMenuItem.Enabled = false;
            this.runSelectedTestMenuItem.Name = "runSelectedTestMenuItem";
            this.runSelectedTestMenuItem.Size = new System.Drawing.Size(279, 26);
            this.runSelectedTestMenuItem.Text = "&Selected Tests";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(276, 6);
            // 
            // runNotRunTestsMenuItem
            // 
            this.runNotRunTestsMenuItem.Image = global::Rubberduck.Properties.Resources.question_white;
            this.runNotRunTestsMenuItem.Name = "runNotRunTestsMenuItem";
            this.runNotRunTestsMenuItem.Size = new System.Drawing.Size(279, 26);
            this.runNotRunTestsMenuItem.Text = "&Not Run Tests";
            // 
            // runFailedTestsMenuItem
            // 
            this.runFailedTestsMenuItem.Enabled = false;
            this.runFailedTestsMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("runFailedTestsMenuItem.Image")));
            this.runFailedTestsMenuItem.Name = "runFailedTestsMenuItem";
            this.runFailedTestsMenuItem.Size = new System.Drawing.Size(279, 26);
            this.runFailedTestsMenuItem.Text = "&Failed Tests";
            // 
            // runPassedTestsMenuItem
            // 
            this.runPassedTestsMenuItem.Enabled = false;
            this.runPassedTestsMenuItem.Image = global::Rubberduck.Properties.Resources.tick_circle;
            this.runPassedTestsMenuItem.Name = "runPassedTestsMenuItem";
            this.runPassedTestsMenuItem.Size = new System.Drawing.Size(279, 26);
            this.runPassedTestsMenuItem.Text = "&Passed Tests";
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(276, 6);
            // 
            // runLastRunMenuItem
            // 
            this.runLastRunMenuItem.Enabled = false;
            this.runLastRunMenuItem.Name = "runLastRunMenuItem";
            this.runLastRunMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.L)));
            this.runLastRunMenuItem.Size = new System.Drawing.Size(279, 26);
            this.runLastRunMenuItem.Text = "Repeat &Last Run";
            // 
            // addButton
            // 
            this.addButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addTestModuleButton,
            this.toolStripSeparator2,
            this.addTestMethodButton,
            this.addExpectedErrorTestMethodButton});
            this.addButton.Image = global::Rubberduck.Properties.Resources.flask_plus;
            this.addButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(71, 27);
            this.addButton.Text = "&Add";
            // 
            // addTestModuleButton
            // 
            this.addTestModuleButton.Image = global::Rubberduck.Properties.Resources.flask_empty;
            this.addTestModuleButton.Name = "addTestModuleButton";
            this.addTestModuleButton.Size = new System.Drawing.Size(277, 26);
            this.addTestModuleButton.Text = "Test &Module";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(274, 6);
            // 
            // addTestMethodButton
            // 
            this.addTestMethodButton.Image = global::Rubberduck.Properties.Resources.flask;
            this.addTestMethodButton.ImageTransparentColor = System.Drawing.Color.Fuchsia;
            this.addTestMethodButton.Name = "addTestMethodButton";
            this.addTestMethodButton.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.T)));
            this.addTestMethodButton.Size = new System.Drawing.Size(277, 26);
            this.addTestMethodButton.Text = "&Test Method";
            // 
            // addExpectedErrorTestMethodButton
            // 
            this.addExpectedErrorTestMethodButton.Image = global::Rubberduck.Properties.Resources.flask_exclamation;
            this.addExpectedErrorTestMethodButton.Name = "addExpectedErrorTestMethodButton";
            this.addExpectedErrorTestMethodButton.Size = new System.Drawing.Size(277, 26);
            this.addExpectedErrorTestMethodButton.Text = "Test Method (Expected &Error)";
            // 
            // gotoSelectionButton
            // 
            this.gotoSelectionButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gotoSelectionButton.Image = global::Rubberduck.Properties.Resources.arrow;
            this.gotoSelectionButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.gotoSelectionButton.Name = "gotoSelectionButton";
            this.gotoSelectionButton.Size = new System.Drawing.Size(24, 27);
            this.gotoSelectionButton.ToolTipText = "Navigate to selected";
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(6, 30);
            // 
            // testProgressBar
            // 
            this.testProgressBar.ForeColor = System.Drawing.Color.LimeGreen;
            this.testProgressBar.Name = "testProgressBar";
            this.testProgressBar.Size = new System.Drawing.Size(133, 27);
            this.testProgressBar.Step = 1;
            // 
            // TotalElapsedMilisecondsLabel
            // 
            this.TotalElapsedMilisecondsLabel.Name = "TotalElapsedMilisecondsLabel";
            this.TotalElapsedMilisecondsLabel.Size = new System.Drawing.Size(40, 27);
            this.TotalElapsedMilisecondsLabel.Text = "0 ms";
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            this.toolStripSeparator6.Size = new System.Drawing.Size(6, 30);
            // 
            // passedTestsLabel
            // 
            this.passedTestsLabel.Name = "passedTestsLabel";
            this.passedTestsLabel.Size = new System.Drawing.Size(65, 27);
            this.passedTestsLabel.Text = "0 Passed";
            // 
            // failedTestsLabel
            // 
            this.failedTestsLabel.Name = "failedTestsLabel";
            this.failedTestsLabel.Size = new System.Drawing.Size(60, 27);
            this.failedTestsLabel.Text = "0 Failed";
            // 
            // inconclusiveTestsLabel
            // 
            this.inconclusiveTestsLabel.Name = "inconclusiveTestsLabel";
            this.inconclusiveTestsLabel.Size = new System.Drawing.Size(101, 27);
            this.inconclusiveTestsLabel.Text = "0 Inconclusive";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.testOutputGridView);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 30);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(800, 216);
            this.panel1.TabIndex = 2;
            // 
            // testOutputGridView
            // 
            this.testOutputGridView.AllowUserToAddRows = false;
            this.testOutputGridView.AllowUserToDeleteRows = false;
            this.testOutputGridView.AllowUserToOrderColumns = true;
            this.testOutputGridView.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
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
            this.testOutputGridView.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.testOutputGridView.Name = "testOutputGridView";
            this.testOutputGridView.ReadOnly = true;
            this.testOutputGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.testOutputGridView.Size = new System.Drawing.Size(800, 216);
            this.testOutputGridView.TabIndex = 1;
            this.testOutputGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.GridCellDoubleClicked);
            this.testOutputGridView.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.ColumnHeaderMouseClicked);
            // 
            // TestExplorerWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolStrip1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MinimumSize = new System.Drawing.Size(800, 246);
            this.Name = "TestExplorerWindow";
            this.Size = new System.Drawing.Size(800, 246);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.testOutputGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ToolStrip toolStrip1;
        private ToolStripButton refreshTestsButton;
        private Panel panel1;
        private DataGridView testOutputGridView;
        private ToolStripDropDownButton runButton;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem runSelectedTestMenuItem;
        private ToolStripMenuItem runAllTestsMenuItem;
        private ToolStripMenuItem runFailedTestsMenuItem;
        private ToolStripMenuItem runNotRunTestsMenuItem;
        private ToolStripMenuItem runPassedTestsMenuItem;
        private ToolStripMenuItem runLastRunMenuItem;
        private ToolStripDropDownButton addButton;
        private ToolStripMenuItem addTestModuleButton;
        private ToolStripMenuItem addTestMethodButton;
        private ToolStripMenuItem addExpectedErrorTestMethodButton;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripSeparator toolStripSeparator4;
        private ToolStripButton gotoSelectionButton;
        private ToolStripSeparator toolStripSeparator5;
        private ToolStripProgressBar testProgressBar;
        private ToolStripLabel passedTestsLabel;
        private ToolStripLabel failedTestsLabel;
        private ToolStripLabel inconclusiveTestsLabel;
        private ToolStripLabel TotalElapsedMilisecondsLabel;
        private ToolStripSeparator toolStripSeparator6;

    }
}