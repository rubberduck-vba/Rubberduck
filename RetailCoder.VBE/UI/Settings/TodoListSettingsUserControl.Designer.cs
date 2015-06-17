using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.Settings
{
    partial class TodoListSettingsUserControl
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
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.AddButton = new System.Windows.Forms.ToolStripButton();
            this.RemoveButton = new System.Windows.Forms.ToolStripButton();
            this.TodoMarkersGridView = new System.Windows.Forms.DataGridView();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TodoMarkersGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddButton,
            this.RemoveButton});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.toolStrip1.Size = new System.Drawing.Size(419, 27);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // AddButton
            // 
            this.AddButton.Image = global::Rubberduck.Properties.Resources.plus_circle;
            this.AddButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AddButton.Name = "AddButton";
            this.AddButton.Size = new System.Drawing.Size(61, 24);
            this.AddButton.Text = "Add";
            this.AddButton.ToolTipText = "Add todo marker";
            this.AddButton.Click += new System.EventHandler(this.addButton_Click);
            // 
            // RemoveButton
            // 
            this.RemoveButton.Image = global::Rubberduck.Properties.Resources.minus_circle;
            this.RemoveButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RemoveButton.Margin = new System.Windows.Forms.Padding(5, 1, 0, 2);
            this.RemoveButton.Name = "RemoveButton";
            this.RemoveButton.Size = new System.Drawing.Size(87, 24);
            this.RemoveButton.Text = "Remove";
            this.RemoveButton.ToolTipText = "Remove todo marker";
            this.RemoveButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // TodoMarkersGridView
            // 
            this.TodoMarkersGridView.AllowUserToAddRows = false;
            this.TodoMarkersGridView.AllowUserToDeleteRows = false;
            this.TodoMarkersGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TodoMarkersGridView.Location = new System.Drawing.Point(0, 28);
            this.TodoMarkersGridView.Name = "TodoMarkersGridView";
            this.TodoMarkersGridView.RowTemplate.Height = 24;
            this.TodoMarkersGridView.Size = new System.Drawing.Size(419, 331);
            this.TodoMarkersGridView.TabIndex = 1;
            // 
            // TodoListSettingsUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.TodoMarkersGridView);
            this.Controls.Add(this.toolStrip1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MinimumSize = new System.Drawing.Size(419, 362);
            this.Name = "TodoListSettingsUserControl";
            this.Size = new System.Drawing.Size(419, 362);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TodoMarkersGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ToolStrip toolStrip1;
        private DataGridView TodoMarkersGridView;
        private ToolStripButton AddButton;
        private ToolStripButton RemoveButton;

    }
}
