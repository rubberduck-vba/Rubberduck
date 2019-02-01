namespace Rubberduck.UI.Refactorings
{
    partial class RefactoringDialogBase
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

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RefactoringDialogBase));
            this.elementHost = new System.Windows.Forms.Integration.ElementHost();
            this.SuspendLayout();
            // 
            // elementHost
            // 
            this.elementHost.AutoSize = true;
            this.elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost.Location = new System.Drawing.Point(0, 0);
            this.elementHost.Margin = new System.Windows.Forms.Padding(5);
            this.elementHost.Name = "elementHost";
            this.elementHost.Size = new System.Drawing.Size(384, 112);
            this.elementHost.TabIndex = 0;
            this.elementHost.Child = null;
            // 
            // RefactoringDialogBase
            // 
            this.ClientSize = new System.Drawing.Size(384, 112);
            this.Controls.Add(this.elementHost);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "RefactoringDialogBase";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost;
    }
}
