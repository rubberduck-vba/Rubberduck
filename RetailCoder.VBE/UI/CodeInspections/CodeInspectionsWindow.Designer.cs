using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.CodeInspections
{
    partial class CodeInspectionsWindow
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CodeInspectionsWindow));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.RefreshButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.StatusLabel = new System.Windows.Forms.ToolStripLabel();
            this.QuickFixButton = new System.Windows.Forms.ToolStripSplitButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.GoButton = new System.Windows.Forms.ToolStripButton();
            this.PreviousButton = new System.Windows.Forms.ToolStripButton();
            this.NextButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.CopyButton = new System.Windows.Forms.ToolStripButton();
            this.configureButton = new System.Windows.Forms.ToolStripButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.CodeIssuesGridView = new System.Windows.Forms.DataGridView();
            this.toolStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CodeIssuesGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshButton,
            this.toolStripSeparator3,
            this.StatusLabel,
            this.QuickFixButton,
            this.toolStripSeparator1,
            this.GoButton,
            this.PreviousButton,
            this.NextButton,
            this.toolStripSeparator2,
            this.CopyButton,
            this.configureButton});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(555, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // RefreshButton
            // 
            this.RefreshButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.RefreshButton.Image = ((System.Drawing.Image)(resources.GetObject("RefreshButton.Image")));
            this.RefreshButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.Size = new System.Drawing.Size(23, 22);
            this.RefreshButton.ToolTipText = "Run code inspections";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // StatusLabel
            // 
            this.StatusLabel.Image = ((System.Drawing.Image)(resources.GetObject("StatusLabel.Image")));
            this.StatusLabel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StatusLabel.Margin = new System.Windows.Forms.Padding(2, 1, 4, 2);
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(63, 22);
            this.StatusLabel.Text = "0 issues";
            // 
            // QuickFixButton
            // 
            this.QuickFixButton.Enabled = false;
            this.QuickFixButton.Image = global::Rubberduck.Properties.Resources.tick;
            this.QuickFixButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.QuickFixButton.Name = "QuickFixButton";
            this.QuickFixButton.Size = new System.Drawing.Size(53, 22);
            this.QuickFixButton.Text = "Fix";
            this.QuickFixButton.ToolTipText = "Address the issue";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // GoButton
            // 
            this.GoButton.Enabled = false;
            this.GoButton.Image = global::Rubberduck.Properties.Resources.arrow;
            this.GoButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.GoButton.Name = "GoButton";
            this.GoButton.Size = new System.Drawing.Size(42, 22);
            this.GoButton.Text = "Go";
            this.GoButton.ToolTipText = "Navigate to selected issue.";
            // 
            // PreviousButton
            // 
            this.PreviousButton.Enabled = false;
            this.PreviousButton.Image = ((System.Drawing.Image)(resources.GetObject("PreviousButton.Image")));
            this.PreviousButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.PreviousButton.Name = "PreviousButton";
            this.PreviousButton.Size = new System.Drawing.Size(72, 22);
            this.PreviousButton.Text = "Previous";
            this.PreviousButton.ToolTipText = "Navigate to previous issue.";
            // 
            // NextButton
            // 
            this.NextButton.Enabled = false;
            this.NextButton.Image = global::Rubberduck.Properties.Resources.navigation;
            this.NextButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.NextButton.Name = "NextButton";
            this.NextButton.Size = new System.Drawing.Size(51, 22);
            this.NextButton.Text = "Next";
            this.NextButton.ToolTipText = "Navigate to next issue.";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // CopyButton
            // 
            this.CopyButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.CopyButton.Enabled = false;
            this.CopyButton.Image = ((System.Drawing.Image)(resources.GetObject("CopyButton.Image")));
            this.CopyButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.CopyButton.Name = "CopyButton";
            this.CopyButton.Size = new System.Drawing.Size(23, 22);
            this.CopyButton.ToolTipText = "Copy inspection results to clipboard";
            // 
            // configureButton
            // 
            this.configureButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.configureButton.Image = global::Rubberduck.Properties.Resources.gear;
            this.configureButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.configureButton.Name = "configureButton";
            this.configureButton.Size = new System.Drawing.Size(23, 22);
            this.configureButton.Visible = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.CodeIssuesGridView);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(555, 100);
            this.panel1.TabIndex = 1;
            // 
            // CodeIssuesGridView
            // 
            this.CodeIssuesGridView.AllowUserToAddRows = false;
            this.CodeIssuesGridView.AllowUserToDeleteRows = false;
            this.CodeIssuesGridView.AllowUserToOrderColumns = true;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.Lavender;
            this.CodeIssuesGridView.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.CodeIssuesGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.CodeIssuesGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CodeIssuesGridView.Location = new System.Drawing.Point(0, 0);
            this.CodeIssuesGridView.MultiSelect = false;
            this.CodeIssuesGridView.Name = "CodeIssuesGridView";
            this.CodeIssuesGridView.ReadOnly = true;
            this.CodeIssuesGridView.Size = new System.Drawing.Size(555, 100);
            this.CodeIssuesGridView.TabIndex = 0;
            // 
            // CodeInspectionsWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolStrip1);
            this.MinimumSize = new System.Drawing.Size(400, 28);
            this.Name = "CodeInspectionsWindow";
            this.Size = new System.Drawing.Size(555, 125);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.CodeIssuesGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ToolStrip toolStrip1;
        private ToolStripButton RefreshButton;
        private Panel panel1;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton GoButton;
        public ToolStripSplitButton QuickFixButton;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton PreviousButton;
        private ToolStripButton NextButton;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripLabel StatusLabel;
        private DataGridView CodeIssuesGridView;
        private ToolStripButton CopyButton;
        private ToolStripButton configureButton;
    }
}
