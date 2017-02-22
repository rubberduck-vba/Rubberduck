using System.ComponentModel;
using System.Windows.Forms;
using Rubberduck.UI.Inspections;

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
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.TodoExplorerControl = new Rubberduck.UI.ToDoItems.ToDoExplorerControl();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(150, 150);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.TodoExplorerControl;
            // 
            // ToDoExplorerWindow
            // 
            this.Controls.Add(this.elementHost1);
            this.Name = "ToDoExplorerWindow";
            this.ResumeLayout(false);

        }


        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private ToDoExplorerControl TodoExplorerControl;
    }
}
