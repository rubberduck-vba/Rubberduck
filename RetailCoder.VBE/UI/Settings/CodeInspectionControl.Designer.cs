using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.Settings
{
    partial class CodeInspectionSettingsControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.codeInspectionsGrid = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.codeInspectionsGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // codeInspectionsGrid
            // 
            this.codeInspectionsGrid.AllowUserToAddRows = false;
            this.codeInspectionsGrid.AllowUserToDeleteRows = false;
            this.codeInspectionsGrid.AllowUserToResizeRows = false;
            this.codeInspectionsGrid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.codeInspectionsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.PowderBlue;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.MediumBlue;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.codeInspectionsGrid.DefaultCellStyle = dataGridViewCellStyle1;
            this.codeInspectionsGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.codeInspectionsGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2;
            this.codeInspectionsGrid.GridColor = System.Drawing.Color.LightGray;
            this.codeInspectionsGrid.Location = new System.Drawing.Point(0, 0);
            this.codeInspectionsGrid.MultiSelect = false;
            this.codeInspectionsGrid.Name = "codeInspectionsGrid";
            this.codeInspectionsGrid.RowHeadersVisible = false;
            this.codeInspectionsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.codeInspectionsGrid.ShowCellErrors = false;
            this.codeInspectionsGrid.ShowCellToolTips = false;
            this.codeInspectionsGrid.ShowRowErrors = false;
            this.codeInspectionsGrid.Size = new System.Drawing.Size(476, 333);
            this.codeInspectionsGrid.TabIndex = 0;
            // 
            // CodeInspectionSettingsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.codeInspectionsGrid);
            this.Name = "CodeInspectionSettingsControl";
            this.Size = new System.Drawing.Size(476, 333);
            ((System.ComponentModel.ISupportInitialize)(this.codeInspectionsGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DataGridView codeInspectionsGrid;

    }
}
