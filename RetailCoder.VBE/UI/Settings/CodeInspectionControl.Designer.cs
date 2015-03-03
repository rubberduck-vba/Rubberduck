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
            this.codeInspectionsGrid = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.codeInspectionsGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // codeInspectionsGrid
            // 
            this.codeInspectionsGrid.AllowUserToAddRows = false;
            this.codeInspectionsGrid.AllowUserToDeleteRows = false;
            this.codeInspectionsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.codeInspectionsGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.codeInspectionsGrid.Location = new System.Drawing.Point(0, 0);
            this.codeInspectionsGrid.Name = "codeInspectionsGrid";
            this.codeInspectionsGrid.Size = new System.Drawing.Size(476, 333);
            this.codeInspectionsGrid.TabIndex = 0;
            // 
            // CodeInspectionControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.codeInspectionsGrid);
            this.Name = "CodeInspectionControl";
            this.Size = new System.Drawing.Size(476, 333);
            ((System.ComponentModel.ISupportInitialize)(this.codeInspectionsGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView codeInspectionsGrid;

    }
}
