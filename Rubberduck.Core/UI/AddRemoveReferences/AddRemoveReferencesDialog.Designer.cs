namespace Rubberduck.UI.AddRemoveReferences
{
    partial class AddRemoveReferencesDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddRemoveReferencesDialog));
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.addRemoveReferencesWindow1 = new Rubberduck.UI.AddRemoveReferences.AddRemoveReferencesWindow();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(800, 550);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.addRemoveReferencesWindow1;
            // 
            // AddRemoveReferencesDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 550);
            this.Controls.Add(this.elementHost1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AddRemoveReferencesDialog";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Rubberduck - Add/Remove References";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private AddRemoveReferencesWindow addRemoveReferencesWindow1;
    }
}