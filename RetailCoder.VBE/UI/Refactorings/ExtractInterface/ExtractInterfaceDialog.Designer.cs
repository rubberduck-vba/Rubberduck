namespace Rubberduck.UI.Refactorings
{
    partial class ExtractInterfaceDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExtractInterfaceDialog));
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.ExtractInterfaceViewElement = new ExtractInterfaceView();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(667, 513);
            this.elementHost1.TabIndex = 1;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.ExtractInterfaceViewElement;
            // 
            // ExtractInterfaceDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(459, 339);
            this.Controls.Add(this.elementHost1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ExtractInterfaceDialog";
            this.ShowInTaskbar = false;
            this.Text = "Rubberduck - Extract Interface";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private ExtractInterfaceView ExtractInterfaceViewElement;
    }
}
