namespace Rubberduck.UI
{
    partial class AboutWindow
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.CloseButton = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.titleLabel = new System.Windows.Forms.Label();
            this.versionLabel = new System.Windows.Forms.Label();
            this.repositoryLinkLabel = new System.Windows.Forms.LinkLabel();
            this.contributorsLabel = new System.Windows.Forms.Label();
            this.retailcoderLinkLabel = new System.Windows.Forms.LinkLabel();
            this.ckuhn203LinkLabel = new System.Windows.Forms.LinkLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.mztoolsLinkLabel = new System.Windows.Forms.LinkLabel();
            this.codereviewLinkLabel = new System.Windows.Forms.LinkLabel();
            this.fugueiconsLinkLabel = new System.Windows.Forms.LinkLabel();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.White;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.CloseButton, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.pictureBox1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(486, 290);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // CloseButton
            // 
            this.CloseButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CloseButton.Location = new System.Drawing.Point(408, 265);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(75, 22);
            this.CloseButton.TabIndex = 0;
            this.CloseButton.Text = "Close";
            this.CloseButton.UseVisualStyleBackColor = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Image = global::Rubberduck.Properties.Resources.rubberduck_adsize;
            this.pictureBox1.Location = new System.Drawing.Point(2, 2);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(229, 258);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.flowLayoutPanel1.Controls.Add(this.titleLabel);
            this.flowLayoutPanel1.Controls.Add(this.versionLabel);
            this.flowLayoutPanel1.Controls.Add(this.repositoryLinkLabel);
            this.flowLayoutPanel1.Controls.Add(this.contributorsLabel);
            this.flowLayoutPanel1.Controls.Add(this.retailcoderLinkLabel);
            this.flowLayoutPanel1.Controls.Add(this.ckuhn203LinkLabel);
            this.flowLayoutPanel1.Controls.Add(this.label1);
            this.flowLayoutPanel1.Controls.Add(this.fugueiconsLinkLabel);
            this.flowLayoutPanel1.Controls.Add(this.codereviewLinkLabel);
            this.flowLayoutPanel1.Controls.Add(this.mztoolsLinkLabel);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanel1.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.flowLayoutPanel1.Location = new System.Drawing.Point(237, 4);
            this.flowLayoutPanel1.Margin = new System.Windows.Forms.Padding(4);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Padding = new System.Windows.Forms.Padding(3);
            this.flowLayoutPanel1.Size = new System.Drawing.Size(245, 254);
            this.flowLayoutPanel1.TabIndex = 2;
            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLabel.Location = new System.Drawing.Point(6, 3);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(96, 17);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "assemblyname";
            // 
            // versionLabel
            // 
            this.versionLabel.AutoSize = true;
            this.versionLabel.Location = new System.Drawing.Point(6, 20);
            this.versionLabel.Name = "versionLabel";
            this.versionLabel.Size = new System.Drawing.Size(69, 17);
            this.versionLabel.TabIndex = 1;
            this.versionLabel.Text = "versioninfo";
            // 
            // repositoryLinkLabel
            // 
            this.repositoryLinkLabel.ActiveLinkColor = System.Drawing.Color.Navy;
            this.repositoryLinkLabel.AutoSize = true;
            this.repositoryLinkLabel.LinkColor = System.Drawing.Color.Blue;
            this.repositoryLinkLabel.Location = new System.Drawing.Point(6, 47);
            this.repositoryLinkLabel.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
            this.repositoryLinkLabel.Name = "repositoryLinkLabel";
            this.repositoryLinkLabel.Size = new System.Drawing.Size(111, 17);
            this.repositoryLinkLabel.TabIndex = 2;
            this.repositoryLinkLabel.TabStop = true;
            this.repositoryLinkLabel.Text = "GitHub Repository";
            this.repositoryLinkLabel.VisitedLinkColor = System.Drawing.Color.DarkSlateBlue;
            // 
            // contributorsLabel
            // 
            this.contributorsLabel.AutoSize = true;
            this.contributorsLabel.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.contributorsLabel.Location = new System.Drawing.Point(6, 74);
            this.contributorsLabel.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
            this.contributorsLabel.Name = "contributorsLabel";
            this.contributorsLabel.Size = new System.Drawing.Size(84, 17);
            this.contributorsLabel.TabIndex = 3;
            this.contributorsLabel.Text = "Contributors";
            // 
            // retailcoderLinkLabel
            // 
            this.retailcoderLinkLabel.ActiveLinkColor = System.Drawing.Color.Navy;
            this.retailcoderLinkLabel.AutoSize = true;
            this.retailcoderLinkLabel.LinkColor = System.Drawing.Color.Blue;
            this.retailcoderLinkLabel.Location = new System.Drawing.Point(6, 91);
            this.retailcoderLinkLabel.Name = "retailcoderLinkLabel";
            this.retailcoderLinkLabel.Size = new System.Drawing.Size(106, 17);
            this.retailcoderLinkLabel.TabIndex = 4;
            this.retailcoderLinkLabel.TabStop = true;
            this.retailcoderLinkLabel.Text = "Mathieu Guindon";
            this.retailcoderLinkLabel.VisitedLinkColor = System.Drawing.Color.DarkSlateBlue;
            // 
            // ckuhn203LinkLabel
            // 
            this.ckuhn203LinkLabel.ActiveLinkColor = System.Drawing.Color.Navy;
            this.ckuhn203LinkLabel.AutoSize = true;
            this.ckuhn203LinkLabel.LinkColor = System.Drawing.Color.Blue;
            this.ckuhn203LinkLabel.Location = new System.Drawing.Point(6, 108);
            this.ckuhn203LinkLabel.Name = "ckuhn203LinkLabel";
            this.ckuhn203LinkLabel.Size = new System.Drawing.Size(142, 17);
            this.ckuhn203LinkLabel.TabIndex = 5;
            this.ckuhn203LinkLabel.TabStop = true;
            this.ckuhn203LinkLabel.Text = "Christopher J. McClellan";
            this.ckuhn203LinkLabel.VisitedLinkColor = System.Drawing.Color.DarkSlateBlue;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 135);
            this.label1.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 17);
            this.label1.TabIndex = 6;
            this.label1.Text = "Special thanks";
            // 
            // mztoolsLinkLabel
            // 
            this.mztoolsLinkLabel.ActiveLinkColor = System.Drawing.Color.Navy;
            this.mztoolsLinkLabel.AutoSize = true;
            this.mztoolsLinkLabel.LinkColor = System.Drawing.Color.Blue;
            this.mztoolsLinkLabel.Location = new System.Drawing.Point(6, 186);
            this.mztoolsLinkLabel.Name = "mztoolsLinkLabel";
            this.mztoolsLinkLabel.Size = new System.Drawing.Size(104, 17);
            this.mztoolsLinkLabel.TabIndex = 7;
            this.mztoolsLinkLabel.TabStop = true;
            this.mztoolsLinkLabel.Text = "MZ-Tools Articles";
            this.mztoolsLinkLabel.VisitedLinkColor = System.Drawing.Color.DarkSlateBlue;
            // 
            // codereviewLinkLabel
            // 
            this.codereviewLinkLabel.ActiveLinkColor = System.Drawing.Color.Navy;
            this.codereviewLinkLabel.AutoSize = true;
            this.codereviewLinkLabel.LinkColor = System.Drawing.Color.Blue;
            this.codereviewLinkLabel.Location = new System.Drawing.Point(6, 169);
            this.codereviewLinkLabel.Name = "codereviewLinkLabel";
            this.codereviewLinkLabel.Size = new System.Drawing.Size(170, 17);
            this.codereviewLinkLabel.TabIndex = 8;
            this.codereviewLinkLabel.TabStop = true;
            this.codereviewLinkLabel.Text = "Code Review Stack Exchange";
            this.codereviewLinkLabel.VisitedLinkColor = System.Drawing.Color.DarkSlateBlue;
            // 
            // fugueiconsLinkLabel
            // 
            this.fugueiconsLinkLabel.ActiveLinkColor = System.Drawing.Color.Navy;
            this.fugueiconsLinkLabel.AutoSize = true;
            this.fugueiconsLinkLabel.LinkColor = System.Drawing.Color.Blue;
            this.fugueiconsLinkLabel.Location = new System.Drawing.Point(6, 152);
            this.fugueiconsLinkLabel.Name = "fugueiconsLinkLabel";
            this.fugueiconsLinkLabel.Size = new System.Drawing.Size(199, 17);
            this.fugueiconsLinkLabel.TabIndex = 10;
            this.fugueiconsLinkLabel.TabStop = true;
            this.fugueiconsLinkLabel.Text = "Yusuke Kamiyamane (Fugue Icons)";
            this.fugueiconsLinkLabel.VisitedLinkColor = System.Drawing.Color.DarkSlateBlue;
            // 
            // AboutWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(486, 290);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "About";
            this.TopMost = true;
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Label versionLabel;
        private System.Windows.Forms.LinkLabel repositoryLinkLabel;
        private System.Windows.Forms.Label contributorsLabel;
        private System.Windows.Forms.LinkLabel retailcoderLinkLabel;
        private System.Windows.Forms.LinkLabel ckuhn203LinkLabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.LinkLabel fugueiconsLinkLabel;
        private System.Windows.Forms.LinkLabel codereviewLinkLabel;
        private System.Windows.Forms.LinkLabel mztoolsLinkLabel;
    }
}