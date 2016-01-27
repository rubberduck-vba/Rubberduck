using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    partial class CloneRepositoryForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CloneRepositoryForm));
            this.RemotePathTextBox = new System.Windows.Forms.TextBox();
            this.CancelButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.InstructionsLabel = new System.Windows.Forms.Label();
            this.TitleLabel = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.InvalidRemotePathValidationIcon = new System.Windows.Forms.PictureBox();
            this.LocalDirectoryTextBox = new System.Windows.Forms.TextBox();
            this.RemotePathLabel = new System.Windows.Forms.Label();
            this.LocalDirectoryLabel = new System.Windows.Forms.Label();
            this.BrowseLocalDirectoryLocationButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidRemotePathValidationIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // RemotePathTextBox
            // 
            resources.ApplyResources(this.RemotePathTextBox, "RemotePathTextBox");
            this.RemotePathTextBox.Name = "RemotePathTextBox";
            // 
            // CancelButton
            // 
            this.CancelButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.CancelButton, "CancelButton");
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.UseVisualStyleBackColor = false;
            // 
            // OkButton
            // 
            this.OkButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            resources.ApplyResources(this.OkButton, "OkButton");
            this.OkButton.Name = "OkButton";
            this.OkButton.UseVisualStyleBackColor = false;
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelButton);
            this.flowLayoutPanel2.Controls.Add(this.OkButton);
            resources.ApplyResources(this.flowLayoutPanel2, "flowLayoutPanel2");
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            // 
            // InstructionsLabel
            // 
            resources.ApplyResources(this.InstructionsLabel, "InstructionsLabel");
            this.InstructionsLabel.Name = "InstructionsLabel";
            // 
            // TitleLabel
            // 
            resources.ApplyResources(this.TitleLabel, "TitleLabel");
            this.TitleLabel.Name = "TitleLabel";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.InstructionsLabel);
            this.panel1.Controls.Add(this.TitleLabel);
            resources.ApplyResources(this.panel1, "panel1");
            this.panel1.Name = "panel1";
            // 
            // InvalidRemotePathValidationIcon
            // 
            this.InvalidRemotePathValidationIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            resources.ApplyResources(this.InvalidRemotePathValidationIcon, "InvalidRemotePathValidationIcon");
            this.InvalidRemotePathValidationIcon.Name = "InvalidRemotePathValidationIcon";
            this.InvalidRemotePathValidationIcon.TabStop = false;
            // 
            // LocalDirectoryTextBox
            // 
            resources.ApplyResources(this.LocalDirectoryTextBox, "LocalDirectoryTextBox");
            this.LocalDirectoryTextBox.Name = "LocalDirectoryTextBox";
            // 
            // RemotePathLabel
            // 
            resources.ApplyResources(this.RemotePathLabel, "RemotePathLabel");
            this.RemotePathLabel.Name = "RemotePathLabel";
            // 
            // LocalDirectoryLabel
            // 
            resources.ApplyResources(this.LocalDirectoryLabel, "LocalDirectoryLabel");
            this.LocalDirectoryLabel.Name = "LocalDirectoryLabel";
            // 
            // BrowseLocalDirectoryLocationButton
            // 
            resources.ApplyResources(this.BrowseLocalDirectoryLocationButton, "BrowseLocalDirectoryLocationButton");
            this.BrowseLocalDirectoryLocationButton.Name = "BrowseLocalDirectoryLocationButton";
            this.BrowseLocalDirectoryLocationButton.UseVisualStyleBackColor = true;
            // 
            // CloneRepositoryForm
            // 
            this.AcceptButton = this.OkButton;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ControlBox = false;
            this.Controls.Add(this.BrowseLocalDirectoryLocationButton);
            this.Controls.Add(this.LocalDirectoryLabel);
            this.Controls.Add(this.RemotePathLabel);
            this.Controls.Add(this.LocalDirectoryTextBox);
            this.Controls.Add(this.InvalidRemotePathValidationIcon);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.RemotePathTextBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CloneRepositoryForm";
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.flowLayoutPanel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidRemotePathValidationIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TextBox RemotePathTextBox;
        private Button CancelButton;
        private Button OkButton;
        private FlowLayoutPanel flowLayoutPanel2;
        private Label InstructionsLabel;
        private Label TitleLabel;
        private Panel panel1;
        private PictureBox InvalidRemotePathValidationIcon;
        private TextBox LocalDirectoryTextBox;
        private Label RemotePathLabel;
        private Label LocalDirectoryLabel;
        private Button BrowseLocalDirectoryLocationButton;
    }
}