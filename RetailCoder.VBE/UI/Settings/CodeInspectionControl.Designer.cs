namespace Rubberduck.UI.Settings
{
    partial class CodeInspectionControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.InspectionName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.InspectionType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.InspectionSeverity = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.InspectionName,
            this.InspectionType,
            this.InspectionSeverity});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(476, 333);
            this.dataGridView1.TabIndex = 0;
            // 
            // InspectionName
            // 
            this.InspectionName.Frozen = true;
            this.InspectionName.HeaderText = "Inspection";
            this.InspectionName.Name = "InspectionName";
            this.InspectionName.ReadOnly = true;
            // 
            // InspectionType
            // 
            this.InspectionType.Frozen = true;
            this.InspectionType.HeaderText = "Type";
            this.InspectionType.Name = "InspectionType";
            this.InspectionType.ReadOnly = true;
            // 
            // InspectionSeverity
            // 
            this.InspectionSeverity.HeaderText = "Severity";
            this.InspectionSeverity.Name = "InspectionSeverity";
            // 
            // CodeInspectionControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridView1);
            this.Name = "CodeInspectionControl";
            this.Size = new System.Drawing.Size(476, 333);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn InspectionName;
        private System.Windows.Forms.DataGridViewTextBoxColumn InspectionType;
        private System.Windows.Forms.DataGridViewComboBoxColumn InspectionSeverity;

    }
}
