namespace Rubberduck.UI.SourceControl
{
    partial class MergeForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MergeForm));
            this.SourceSelector = new System.Windows.Forms.ComboBox();
            this.DestinationSelector = new System.Windows.Forms.ComboBox();
            this.OkayButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SourceLabel = new System.Windows.Forms.Label();
            this.DestinationLabel = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // SourceSelector
            // 
            this.SourceSelector.FormattingEnabled = true;
            this.SourceSelector.Location = new System.Drawing.Point(21, 32);
            this.SourceSelector.Name = "SourceSelector";
            this.SourceSelector.Size = new System.Drawing.Size(121, 21);
            this.SourceSelector.TabIndex = 0;
            this.SourceSelector.SelectedIndexChanged += new System.EventHandler(this.OnSelectedSourceBranchChanged);
            // 
            // DestinationSelector
            // 
            this.DestinationSelector.FormattingEnabled = true;
            this.DestinationSelector.Location = new System.Drawing.Point(189, 33);
            this.DestinationSelector.Name = "DestinationSelector";
            this.DestinationSelector.Size = new System.Drawing.Size(121, 21);
            this.DestinationSelector.TabIndex = 1;
            this.DestinationSelector.SelectedIndexChanged += new System.EventHandler(this.OnSelectedDestinationBranchChanged);
            // 
            // OkayButton
            // 
            this.OkayButton.Image = global::Rubberduck.Properties.Resources.tick;
            this.OkayButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.OkayButton.Location = new System.Drawing.Point(41, 81);
            this.OkayButton.Name = "OkayButton";
            this.OkayButton.Size = new System.Drawing.Size(89, 23);
            this.OkayButton.TabIndex = 2;
            this.OkayButton.Text = "Okay";
            this.OkayButton.UseVisualStyleBackColor = true;
            this.OkayButton.Click += new System.EventHandler(this.OnConfirm);
            // 
            // CancelButton
            // 
            this.CancelButton.AccessibleDescription = "Cancel";
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.CancelButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.CancelButton.Location = new System.Drawing.Point(205, 81);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(87, 23);
            this.CancelButton.TabIndex = 3;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.OnCancel);
            // 
            // SourceLabel
            // 
            this.SourceLabel.AutoSize = true;
            this.SourceLabel.Location = new System.Drawing.Point(21, 13);
            this.SourceLabel.Name = "SourceLabel";
            this.SourceLabel.Size = new System.Drawing.Size(41, 13);
            this.SourceLabel.TabIndex = 4;
            this.SourceLabel.Text = "Source";
            // 
            // DestinationLabel
            // 
            this.DestinationLabel.AutoSize = true;
            this.DestinationLabel.Location = new System.Drawing.Point(189, 13);
            this.DestinationLabel.Name = "DestinationLabel";
            this.DestinationLabel.Size = new System.Drawing.Size(60, 13);
            this.DestinationLabel.TabIndex = 5;
            this.DestinationLabel.Text = "Destination";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Rubberduck.Properties.Resources.arrow;
            this.pictureBox1.Location = new System.Drawing.Point(148, 32);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(36, 20);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // MergeForm
            // 
            this.AcceptButton = this.OkayButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelButton;
            this.ClientSize = new System.Drawing.Size(331, 132);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.DestinationLabel);
            this.Controls.Add(this.SourceLabel);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OkayButton);
            this.Controls.Add(this.DestinationSelector);
            this.Controls.Add(this.SourceSelector);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MergeForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Merge Branch";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox SourceSelector;
        private System.Windows.Forms.ComboBox DestinationSelector;
        private System.Windows.Forms.Button OkayButton;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Label SourceLabel;
        private System.Windows.Forms.Label DestinationLabel;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}