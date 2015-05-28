using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    partial class ReorderParametersDialog
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(ReorderParametersDialog));
            this.MoveDownButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.InstructionsLabel = new System.Windows.Forms.Label();
            this.TitleLabel = new System.Windows.Forms.Label();
            this.MethodParametersGrid = new System.Windows.Forms.DataGridView();
            this.MoveUpButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MethodParametersGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // MoveDownButton
            // 
            this.MoveDownButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.MoveDownButton.Image = global::Rubberduck.Properties.Resources.arrow_270;
            this.MoveDownButton.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.MoveDownButton.Location = new System.Drawing.Point(316, 154);
            this.MoveDownButton.Name = "MoveDownButton";
            this.MoveDownButton.Size = new System.Drawing.Size(75, 72);
            this.MoveDownButton.TabIndex = 2;
            this.MoveDownButton.Text = "Move down";
            this.MoveDownButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.MoveDownButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            this.MoveDownButton.UseVisualStyleBackColor = true;
            this.MoveDownButton.Click += new System.EventHandler(this.MoveDownButtonClicked);
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelButton);
            this.flowLayoutPanel2.Controls.Add(this.OkButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 232);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(8, 8, 0, 8);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(402, 43);
            this.flowLayoutPanel2.TabIndex = 3;
            // 
            // CancelButton
            // 
            this.CancelButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(316, 11);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 0;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = false;
            // 
            // OkButton
            // 
            this.OkButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point(235, 11);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = false;
            this.OkButton.Click += new System.EventHandler(this.OkButtonClick);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.InstructionsLabel);
            this.panel1.Controls.Add(this.TitleLabel);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(402, 71);
            this.panel1.TabIndex = 4;
            // 
            // InstructionsLabel
            // 
            this.InstructionsLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.InstructionsLabel.Location = new System.Drawing.Point(9, 30);
            this.InstructionsLabel.Name = "InstructionsLabel";
            this.InstructionsLabel.Padding = new System.Windows.Forms.Padding(4);
            this.InstructionsLabel.Size = new System.Drawing.Size(383, 34);
            this.InstructionsLabel.TabIndex = 6;
            this.InstructionsLabel.Text = "Select a parameter and drag it or use buttons to move it up or down.";
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TitleLabel.Location = new System.Drawing.Point(12, 9);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Padding = new System.Windows.Forms.Padding(2);
            this.TitleLabel.Size = new System.Drawing.Size(140, 19);
            this.TitleLabel.TabIndex = 4;
            this.TitleLabel.Text = "Reorder parameters";
            // 
            // MethodParametersGrid
            // 
            this.MethodParametersGrid.AllowUserToAddRows = false;
            this.MethodParametersGrid.AllowUserToDeleteRows = false;
            this.MethodParametersGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MethodParametersGrid.BackgroundColor = System.Drawing.Color.White;
            this.MethodParametersGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.MethodParametersGrid.Location = new System.Drawing.Point(9, 77);
            this.MethodParametersGrid.Margin = new System.Windows.Forms.Padding(8, 3, 8, 3);
            this.MethodParametersGrid.MultiSelect = false;
            this.MethodParametersGrid.Name = "MethodParametersGrid";
            this.MethodParametersGrid.RowHeadersVisible = false;
            this.MethodParametersGrid.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.MethodParametersGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.MethodParametersGrid.ShowCellErrors = false;
            this.MethodParametersGrid.ShowCellToolTips = false;
            this.MethodParametersGrid.ShowEditingIcon = false;
            this.MethodParametersGrid.ShowRowErrors = false;
            this.MethodParametersGrid.Size = new System.Drawing.Size(301, 149);
            this.MethodParametersGrid.TabIndex = 8;
            // 
            // MoveUpButton
            // 
            this.MoveUpButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.MoveUpButton.Image = global::Rubberduck.Properties.Resources.arrow_090;
            this.MoveUpButton.Location = new System.Drawing.Point(316, 77);
            this.MoveUpButton.Name = "MoveUpButton";
            this.MoveUpButton.Size = new System.Drawing.Size(75, 72);
            this.MoveUpButton.TabIndex = 1;
            this.MoveUpButton.Text = "Move up";
            this.MoveUpButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.MoveUpButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.MoveUpButton.UseVisualStyleBackColor = true;
            this.MoveUpButton.Click += new System.EventHandler(this.MoveUpButtonClicked);
            // 
            // ReorderParametersDialog
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(402, 275);
            this.Controls.Add(this.MethodParametersGrid);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.MoveDownButton);
            this.Controls.Add(this.MoveUpButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReorderParametersDialog";
            this.Text = "Rubberduck - Reorder Parameters";
            this.flowLayoutPanel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MethodParametersGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Button MoveUpButton;
        private Button MoveDownButton;
        private FlowLayoutPanel flowLayoutPanel2;
        private Button CancelButton;
        private Button OkButton;
        private Panel panel1;
        private Label TitleLabel;
        private DataGridView MethodParametersGrid;
        private Label InstructionsLabel;
    }
}