
namespace Rubberduck.UI
{
    partial class Splash2021
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Splash2021));
            this.VersionLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // VersionLabel
            // 
            this.VersionLabel.AutoSize = true;
            this.VersionLabel.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.VersionLabel.Location = new System.Drawing.Point(12, 9);
            this.VersionLabel.Name = "VersionLabel";
            this.VersionLabel.Padding = new System.Windows.Forms.Padding(2);
            this.VersionLabel.Size = new System.Drawing.Size(76, 17);
            this.VersionLabel.TabIndex = 0;
            this.VersionLabel.Text = "VersionLabel";
            this.VersionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.VersionLabel.UseMnemonic = false;
            // 
            // Splash2021
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Rubberduck.Properties.Resources.Splash2021e;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(320, 360);
            this.ControlBox = false;
            this.Controls.Add(this.VersionLabel);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Splash2021";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Rubberduck";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label VersionLabel;
    }
}