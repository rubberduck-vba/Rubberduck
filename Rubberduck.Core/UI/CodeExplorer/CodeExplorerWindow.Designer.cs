using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.CodeExplorer
{
    partial class CodeExplorerWindow
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
            this.components = new System.ComponentModel.Container();
            this.CodeExplorerContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.RefreshContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddClassContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.AddStdModuleContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.AddFormContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.AddTestModuleContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.NavigateContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.FindAllReferencesContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.FindAllImplementationsContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.RenameContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.ShowDesignerContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.InspectContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.RunAllTestsContextButton = new System.Windows.Forms.ToolStripMenuItem();
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.codeExplorerControl1 = new Rubberduck.UI.CodeExplorer.CodeExplorerControl();
            this.CodeExplorerContextMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // CodeExplorerContextMenu
            // 
            this.CodeExplorerContextMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.CodeExplorerContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshContextButton,
            this.newToolStripMenuItem,
            this.toolStripSeparator4,
            this.NavigateContextButton,
            this.FindAllReferencesContextButton,
            this.FindAllImplementationsContextButton,
            this.RenameContextButton,
            this.ShowDesignerContextButton,
            this.toolStripSeparator5,
            this.InspectContextButton,
            this.RunAllTestsContextButton});
            this.CodeExplorerContextMenu.Name = "CodeExplorerContextMenu";
            this.CodeExplorerContextMenu.Size = new System.Drawing.Size(215, 214);
            // 
            // RefreshContextButton
            // 
            this.RefreshContextButton.Image = global::Rubberduck.Properties.Resources.arrow_circle_double;
            this.RefreshContextButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.RefreshContextButton.Name = "RefreshContextButton";
            this.RefreshContextButton.Size = new System.Drawing.Size(214, 22);
            this.RefreshContextButton.Text = "Refresh";
            // 
            // newToolStripMenuItem
            // 
            this.newToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddClassContextButton,
            this.AddStdModuleContextButton,
            this.AddFormContextButton,
            this.toolStripSeparator6,
            this.AddTestModuleContextButton});
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.newToolStripMenuItem.Text = "&New";
            // 
            // AddClassContextButton
            // 
            this.AddClassContextButton.Image = global::Rubberduck.Properties.Resources.AddClass;
            this.AddClassContextButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.AddClassContextButton.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.AddClassContextButton.Name = "AddClassContextButton";
            this.AddClassContextButton.Size = new System.Drawing.Size(165, 22);
            this.AddClassContextButton.Text = "&Class module";
            // 
            // AddStdModuleContextButton
            // 
            this.AddStdModuleContextButton.Image = global::Rubberduck.Properties.Resources.AddModule;
            this.AddStdModuleContextButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.AddStdModuleContextButton.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.AddStdModuleContextButton.Name = "AddStdModuleContextButton";
            this.AddStdModuleContextButton.Size = new System.Drawing.Size(165, 22);
            this.AddStdModuleContextButton.Text = "Standard &module";
            // 
            // AddFormContextButton
            // 
            this.AddFormContextButton.Image = global::Rubberduck.Properties.Resources.AddForm;
            this.AddFormContextButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.AddFormContextButton.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.AddFormContextButton.Name = "AddFormContextButton";
            this.AddFormContextButton.Size = new System.Drawing.Size(165, 22);
            this.AddFormContextButton.Text = "User &form";
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            this.toolStripSeparator6.Size = new System.Drawing.Size(162, 6);
            // 
            // AddTestModuleContextButton
            // 
            this.AddTestModuleContextButton.Name = "AddTestModuleContextButton";
            this.AddTestModuleContextButton.Size = new System.Drawing.Size(165, 22);
            this.AddTestModuleContextButton.Text = "&Test module";
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(211, 6);
            // 
            // NavigateContextButton
            // 
            this.NavigateContextButton.Name = "NavigateContextButton";
            this.NavigateContextButton.Size = new System.Drawing.Size(214, 22);
            this.NavigateContextButton.Text = "Navi&gate";
            // 
            // FindAllReferencesContextButton
            // 
            this.FindAllReferencesContextButton.Name = "FindAllReferencesContextButton";
            this.FindAllReferencesContextButton.Size = new System.Drawing.Size(214, 22);
            this.FindAllReferencesContextButton.Text = "&Find all references...";
            // 
            // FindAllImplementationsContextButton
            // 
            this.FindAllImplementationsContextButton.Name = "FindAllImplementationsContextButton";
            this.FindAllImplementationsContextButton.Size = new System.Drawing.Size(214, 22);
            this.FindAllImplementationsContextButton.Text = "Find all &implementations...";
            // 
            // RenameContextButton
            // 
            this.RenameContextButton.Name = "RenameContextButton";
            this.RenameContextButton.Size = new System.Drawing.Size(214, 22);
            this.RenameContextButton.Text = "Re&name";
            // 
            // ShowDesignerContextButton
            // 
            this.ShowDesignerContextButton.Enabled = false;
            this.ShowDesignerContextButton.Image = global::Rubberduck.Properties.Resources.ProjectForm;
            this.ShowDesignerContextButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ShowDesignerContextButton.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.ShowDesignerContextButton.Name = "ShowDesignerContextButton";
            this.ShowDesignerContextButton.ShortcutKeys = System.Windows.Forms.Keys.F7;
            this.ShowDesignerContextButton.Size = new System.Drawing.Size(214, 22);
            this.ShowDesignerContextButton.Text = "Show &designer";
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(211, 6);
            // 
            // InspectContextButton
            // 
            this.InspectContextButton.Image = global::Rubberduck.Properties.Resources.light_bulb_code;
            this.InspectContextButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.InspectContextButton.Name = "InspectContextButton";
            this.InspectContextButton.Size = new System.Drawing.Size(214, 22);
            this.InspectContextButton.Text = "&Inspect";
            // 
            // RunAllTestsContextButton
            // 
            this.RunAllTestsContextButton.Name = "RunAllTestsContextButton";
            this.RunAllTestsContextButton.Size = new System.Drawing.Size(214, 22);
            this.RunAllTestsContextButton.Text = "&Run all tests";
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(280, 368);
            this.elementHost1.TabIndex = 1;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.codeExplorerControl1;
            // 
            // CodeExplorerWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.elementHost1);
            this.Name = "CodeExplorerWindow";
            this.Size = new System.Drawing.Size(280, 368);
            this.CodeExplorerContextMenu.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private ContextMenuStrip CodeExplorerContextMenu;
        private ToolStripSeparator toolStripSeparator4;
        private ToolStripMenuItem NavigateContextButton;
        private ToolStripMenuItem ShowDesignerContextButton;
        private ToolStripSeparator toolStripSeparator5;
        private ToolStripMenuItem RunAllTestsContextButton;
        private ToolStripMenuItem newToolStripMenuItem;
        private ToolStripMenuItem AddClassContextButton;
        private ToolStripMenuItem AddStdModuleContextButton;
        private ToolStripMenuItem AddFormContextButton;
        private ToolStripSeparator toolStripSeparator6;
        private ToolStripMenuItem AddTestModuleContextButton;
        private ToolStripMenuItem RefreshContextButton;
        private ToolStripMenuItem InspectContextButton;
        private ToolStripMenuItem RenameContextButton;
        private ToolStripMenuItem FindAllReferencesContextButton;
        private ToolStripMenuItem FindAllImplementationsContextButton;
        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private CodeExplorerControl codeExplorerControl1;
    }
}
