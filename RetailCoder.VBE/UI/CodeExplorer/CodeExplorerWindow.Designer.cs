namespace Rubberduck.UI.CodeExplorer
{
    partial class CodeExplorerWindow
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CodeExplorerWindow));
            this.CodeExplorerToolbar = new System.Windows.Forms.ToolStrip();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.panel1 = new System.Windows.Forms.Panel();
            this.SolutionTree = new System.Windows.Forms.TreeView();
            this.TreeNodeIcons = new System.Windows.Forms.ImageList(this.components);
            this.RefreshButton = new System.Windows.Forms.ToolStripButton();
            this.ShowFoldersToggleButton = new System.Windows.Forms.ToolStripButton();
            this.ShowDesignerButton = new System.Windows.Forms.ToolStripButton();
            this.AddButton = new System.Windows.Forms.ToolStripDropDownButton();
            this.AddClassButton = new System.Windows.Forms.ToolStripMenuItem();
            this.AddStdModuleButton = new System.Windows.Forms.ToolStripMenuItem();
            this.AddFormButton = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.AddTestModuleButton = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.DisplayModeButton = new System.Windows.Forms.ToolStripDropDownButton();
            this.DisplayMemberNamesButton = new System.Windows.Forms.ToolStripMenuItem();
            this.DisplaySignaturesButton = new System.Windows.Forms.ToolStripMenuItem();
            this.CodeExplorerToolbar.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CodeExplorerToolbar
            // 
            this.CodeExplorerToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshButton,
            this.ShowFoldersToggleButton,
            this.toolStripSeparator2,
            this.ShowDesignerButton,
            this.toolStripSeparator3,
            this.DisplayModeButton,
            this.AddButton});
            this.CodeExplorerToolbar.Location = new System.Drawing.Point(0, 0);
            this.CodeExplorerToolbar.Name = "CodeExplorerToolbar";
            this.CodeExplorerToolbar.Size = new System.Drawing.Size(297, 25);
            this.CodeExplorerToolbar.TabIndex = 0;
            this.CodeExplorerToolbar.Text = "toolStrip1";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.SolutionTree);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(297, 343);
            this.panel1.TabIndex = 1;
            // 
            // SolutionTree
            // 
            this.SolutionTree.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SolutionTree.Location = new System.Drawing.Point(0, 0);
            this.SolutionTree.Name = "SolutionTree";
            this.SolutionTree.Size = new System.Drawing.Size(297, 343);
            this.SolutionTree.TabIndex = 0;
            // 
            // TreeNodeIcons
            // 
            this.TreeNodeIcons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("TreeNodeIcons.ImageStream")));
            this.TreeNodeIcons.TransparentColor = System.Drawing.Color.Fuchsia;
            this.TreeNodeIcons.Images.SetKeyName(0, "ClosedFolder");
            this.TreeNodeIcons.Images.SetKeyName(1, "OpenFolder");
            this.TreeNodeIcons.Images.SetKeyName(2, "Form");
            this.TreeNodeIcons.Images.SetKeyName(3, "ClassModule");
            this.TreeNodeIcons.Images.SetKeyName(4, "PrivateClass");
            this.TreeNodeIcons.Images.SetKeyName(5, "Option");
            this.TreeNodeIcons.Images.SetKeyName(6, "Implements");
            this.TreeNodeIcons.Images.SetKeyName(7, "StandardModule");
            this.TreeNodeIcons.Images.SetKeyName(8, "PrivateModule");
            this.TreeNodeIcons.Images.SetKeyName(9, "PublicField");
            this.TreeNodeIcons.Images.SetKeyName(10, "PrivateField");
            this.TreeNodeIcons.Images.SetKeyName(11, "FriendField");
            this.TreeNodeIcons.Images.SetKeyName(12, "PublicMethod");
            this.TreeNodeIcons.Images.SetKeyName(13, "FriendMethod");
            this.TreeNodeIcons.Images.SetKeyName(14, "PrivateMethod");
            this.TreeNodeIcons.Images.SetKeyName(15, "TestMethod");
            this.TreeNodeIcons.Images.SetKeyName(16, "PublicProperty");
            this.TreeNodeIcons.Images.SetKeyName(17, "FriendProperty");
            this.TreeNodeIcons.Images.SetKeyName(18, "PrivateProperty");
            this.TreeNodeIcons.Images.SetKeyName(19, "PublicConst");
            this.TreeNodeIcons.Images.SetKeyName(20, "FriendConst");
            this.TreeNodeIcons.Images.SetKeyName(21, "PrivateConst");
            this.TreeNodeIcons.Images.SetKeyName(22, "PublicEnum");
            this.TreeNodeIcons.Images.SetKeyName(23, "FriendEnum");
            this.TreeNodeIcons.Images.SetKeyName(24, "PrivateEnum");
            this.TreeNodeIcons.Images.SetKeyName(25, "EnumItem");
            this.TreeNodeIcons.Images.SetKeyName(26, "PublicEvent");
            this.TreeNodeIcons.Images.SetKeyName(27, "FriendEvent");
            this.TreeNodeIcons.Images.SetKeyName(28, "PrivateEvent");
            this.TreeNodeIcons.Images.SetKeyName(29, "PublicType");
            this.TreeNodeIcons.Images.SetKeyName(30, "FriendType");
            this.TreeNodeIcons.Images.SetKeyName(31, "PrivateType");
            this.TreeNodeIcons.Images.SetKeyName(32, "Operation");
            this.TreeNodeIcons.Images.SetKeyName(33, "CodeBlock");
            this.TreeNodeIcons.Images.SetKeyName(34, "Identifier");
            this.TreeNodeIcons.Images.SetKeyName(35, "Parameter");
            this.TreeNodeIcons.Images.SetKeyName(36, "Assignment");
            this.TreeNodeIcons.Images.SetKeyName(37, "PublicInterface");
            this.TreeNodeIcons.Images.SetKeyName(38, "PrivateInterface");
            this.TreeNodeIcons.Images.SetKeyName(39, "Label");
            this.TreeNodeIcons.Images.SetKeyName(40, "Hourglass");
            this.TreeNodeIcons.Images.SetKeyName(41, "Locked");
            this.TreeNodeIcons.Images.SetKeyName(42, "OfficeDocument");
            // 
            // RefreshButton
            // 
            this.RefreshButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.RefreshButton.Image = global::Rubberduck.Properties.Resources.arrow_circle_double;
            this.RefreshButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.Size = new System.Drawing.Size(23, 22);
            this.RefreshButton.ToolTipText = "Refresh";
            // 
            // ShowFoldersToggleButton
            // 
            this.ShowFoldersToggleButton.Checked = true;
            this.ShowFoldersToggleButton.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ShowFoldersToggleButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ShowFoldersToggleButton.Image = global::Rubberduck.Properties.Resources.VSFolder_closed;
            this.ShowFoldersToggleButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ShowFoldersToggleButton.Name = "ShowFoldersToggleButton";
            this.ShowFoldersToggleButton.Size = new System.Drawing.Size(23, 22);
            this.ShowFoldersToggleButton.ToolTipText = "Toggle folders";
            // 
            // ShowDesignerButton
            // 
            this.ShowDesignerButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ShowDesignerButton.Image = global::Rubberduck.Properties.Resources.VSProject_form;
            this.ShowDesignerButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ShowDesignerButton.Name = "ShowDesignerButton";
            this.ShowDesignerButton.Size = new System.Drawing.Size(23, 22);
            this.ShowDesignerButton.ToolTipText = "Open designer";
            // 
            // AddButton
            // 
            this.AddButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddClassButton,
            this.AddStdModuleButton,
            this.AddFormButton,
            this.toolStripSeparator1,
            this.AddTestModuleButton});
            this.AddButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AddButton.Name = "AddButton";
            this.AddButton.Size = new System.Drawing.Size(42, 22);
            this.AddButton.Text = "Add";
            // 
            // AddClassButton
            // 
            this.AddClassButton.Image = global::Rubberduck.Properties.Resources.AddClass_5561_32;
            this.AddClassButton.Name = "AddClassButton";
            this.AddClassButton.Size = new System.Drawing.Size(197, 22);
            this.AddClassButton.Text = "&Class module (.cls)";
            // 
            // AddStdModuleButton
            // 
            this.AddStdModuleButton.Image = global::Rubberduck.Properties.Resources.AddModule_368_32;
            this.AddStdModuleButton.Name = "AddStdModuleButton";
            this.AddStdModuleButton.Size = new System.Drawing.Size(247, 22);
            this.AddStdModuleButton.Text = "&Standard module (.bas)";
            // 
            // AddFormButton
            // 
            this.AddFormButton.Image = global::Rubberduck.Properties.Resources.AddForm_369_32;
            this.AddFormButton.Name = "AddFormButton";
            this.AddFormButton.Size = new System.Drawing.Size(247, 22);
            this.AddFormButton.Text = "User &form";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(244, 6);
            // 
            // AddTestModuleButton
            // 
            this.AddTestModuleButton.Name = "AddTestModuleButton";
            this.AddTestModuleButton.Size = new System.Drawing.Size(197, 22);
            this.AddTestModuleButton.Text = "&Test module";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // DisplayModeButton
            // 
            this.DisplayModeButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.DisplayModeButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.DisplayMemberNamesButton,
            this.DisplaySignaturesButton});
            this.DisplayModeButton.Image = global::Rubberduck.Properties.Resources.DisplayName_13394_32;
            this.DisplayModeButton.ImageTransparentColor = System.Drawing.Color.White;
            this.DisplayModeButton.Name = "DisplayModeButton";
            this.DisplayModeButton.Size = new System.Drawing.Size(29, 22);
            this.DisplayModeButton.Text = "toolStripSplitButton1";
            // 
            // DisplayMemberNamesButton
            // 
            this.DisplayMemberNamesButton.Checked = true;
            this.DisplayMemberNamesButton.CheckState = System.Windows.Forms.CheckState.Checked;
            this.DisplayMemberNamesButton.Image = global::Rubberduck.Properties.Resources.DisplayName_13394_32;
            this.DisplayMemberNamesButton.Name = "DisplayMemberNamesButton";
            this.DisplayMemberNamesButton.Size = new System.Drawing.Size(198, 22);
            this.DisplayMemberNamesButton.Text = "Display member &names";
            // 
            // DisplaySignaturesButton
            // 
            this.DisplaySignaturesButton.Image = global::Rubberduck.Properties.Resources.DisplayFullSignature_13393_32;
            this.DisplaySignaturesButton.ImageTransparentColor = System.Drawing.Color.White;
            this.DisplaySignaturesButton.Name = "DisplaySignaturesButton";
            this.DisplaySignaturesButton.Size = new System.Drawing.Size(198, 22);
            this.DisplaySignaturesButton.Text = "Display full &signatures";
            // 
            // CodeExplorerWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.CodeExplorerToolbar);
            this.Name = "CodeExplorerWindow";
            this.Size = new System.Drawing.Size(297, 368);
            this.CodeExplorerToolbar.ResumeLayout(false);
            this.CodeExplorerToolbar.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip CodeExplorerToolbar;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ToolStripButton RefreshButton;
        public System.Windows.Forms.TreeView SolutionTree;
        private System.Windows.Forms.ImageList TreeNodeIcons;
        private System.Windows.Forms.ToolStripButton ShowFoldersToggleButton;
        public System.Windows.Forms.ToolStripButton ShowDesignerButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripDropDownButton AddButton;
        private System.Windows.Forms.ToolStripMenuItem AddClassButton;
        private System.Windows.Forms.ToolStripMenuItem AddStdModuleButton;
        private System.Windows.Forms.ToolStripMenuItem AddFormButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem AddTestModuleButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripDropDownButton DisplayModeButton;
        private System.Windows.Forms.ToolStripMenuItem DisplayMemberNamesButton;
        private System.Windows.Forms.ToolStripMenuItem DisplaySignaturesButton;
    }
}
