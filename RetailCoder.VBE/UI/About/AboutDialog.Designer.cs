using Rubberduck.UI.About;

namespace Rubberduck.UI.About
{
    partial class AboutDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutDialog));
            this.ElementHost = new System.Windows.Forms.Integration.ElementHost();
            this.AboutControl = new Rubberduck.UI.About.AboutControl();
            this.SuspendLayout();
            // 
            // ElementHost
            // 
            this.ElementHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ElementHost.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ElementHost.Location = new System.Drawing.Point(0, 0);
            this.ElementHost.Name = "ElementHost";
            this.ElementHost.Size = new System.Drawing.Size(682, 403);
            this.ElementHost.TabIndex = 0;
            this.ElementHost.Child = this.AboutControl;
            // 
            // AboutDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(682, 403);
            this.Controls.Add(this.ElementHost);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutDialog";
            this.ShowInTaskbar = false;
            this.Text = "About";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost ElementHost;
        private AboutControl AboutControl;
    }
}
