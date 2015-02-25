namespace Rubberduck.UI.Settings
{
    partial class _ConfigurationTreeViewControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.settingsTreeView = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // settingsTreeView
            // 
            this.settingsTreeView.BackColor = System.Drawing.SystemColors.Control;
            this.settingsTreeView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.settingsTreeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.settingsTreeView.Location = new System.Drawing.Point(0, 0);
            this.settingsTreeView.Margin = new System.Windows.Forms.Padding(10);
            this.settingsTreeView.Name = "settingsTreeView";
            this.settingsTreeView.Size = new System.Drawing.Size(302, 314);
            this.settingsTreeView.TabIndex = 0;
            this.settingsTreeView.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.settingsTreeView_AfterSelect);
            // 
            // ConfigurationTreeView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.settingsTreeView);
            this.Name = "ConfigurationTreeView";
            this.Size = new System.Drawing.Size(302, 314);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView settingsTreeView;

    }
}
