using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.ToDoItems
{
    partial class ToDoExplorerWindow
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ToDoExplorerWindow));
            this.todoItemsGridView = new System.Windows.Forms.DataGridView();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.refreshButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.configureButton = new System.Windows.Forms.ToolStripButton();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.todoItemsGridView)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // todoItemsGridView
            // 
            this.todoItemsGridView.AllowUserToAddRows = false;
            this.todoItemsGridView.AllowUserToDeleteRows = false;
            this.todoItemsGridView.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.todoItemsGridView.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.todoItemsGridView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.todoItemsGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.todoItemsGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.todoItemsGridView.Location = new System.Drawing.Point(0, 0);
            this.todoItemsGridView.Name = "todoItemsGridView";
            this.todoItemsGridView.ReadOnly = true;
            this.todoItemsGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.todoItemsGridView.Size = new System.Drawing.Size(319, 206);
            this.todoItemsGridView.TabIndex = 0;
            this.todoItemsGridView.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.ColumnHeaderMouseClicked);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshButton,
            this.toolStripSeparator1,
            this.configureButton});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(319, 25);
            this.toolStrip1.TabIndex = 2;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // refreshButton
            // 
            this.refreshButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.refreshButton.Image = ((System.Drawing.Image)(resources.GetObject("refreshButton.Image")));
            this.refreshButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.Size = new System.Drawing.Size(23, 22);
            this.refreshButton.Click += new System.EventHandler(this.RefreshButtonClicked);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // configureButton
            // 
            this.configureButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.configureButton.Image = global::Rubberduck.Properties.Resources.gear;
            this.configureButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.configureButton.Name = "configureButton";
            this.configureButton.Size = new System.Drawing.Size(23, 22);
            this.configureButton.Visible = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.todoItemsGridView);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(319, 206);
            this.panel1.TabIndex = 3;
            // 
            // ToDoExplorerWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "ToDoExplorerWindow";
            this.Size = new System.Drawing.Size(319, 231);
            ((System.ComponentModel.ISupportInitialize)(this.todoItemsGridView)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DataGridView todoItemsGridView;
        private ToolStrip toolStrip1;
        private ToolStripButton refreshButton;
        private ToolStripSeparator toolStripSeparator1;
        private Panel panel1;
        private ToolStripButton configureButton;


    }
}
