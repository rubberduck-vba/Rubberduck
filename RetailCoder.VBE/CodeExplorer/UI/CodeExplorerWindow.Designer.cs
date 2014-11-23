namespace Rubberduck.CodeExplorer.UI
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
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.panel1 = new System.Windows.Forms.Panel();
            this.SolutionTree = new System.Windows.Forms.TreeView();
            this.RefreshButton = new System.Windows.Forms.ToolStripButton();
            this.TreeNodeIcons = new System.Windows.Forms.ImageList(this.components);
            this.CodeExplorerToolbar.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CodeExplorerToolbar
            // 
            this.CodeExplorerToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshButton,
            this.toolStripSeparator1});
            this.CodeExplorerToolbar.Location = new System.Drawing.Point(0, 0);
            this.CodeExplorerToolbar.Name = "CodeExplorerToolbar";
            this.CodeExplorerToolbar.Size = new System.Drawing.Size(297, 25);
            this.CodeExplorerToolbar.TabIndex = 0;
            this.CodeExplorerToolbar.Text = "toolStrip1";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
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
            // RefreshButton
            // 
            this.RefreshButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.RefreshButton.Image = global::Rubberduck.Properties.Resources.arrow_circle_double;
            this.RefreshButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.Size = new System.Drawing.Size(23, 22);
            this.RefreshButton.ToolTipText = "Refresh";
            // 
            // TreeNodeIcons
            // 
            this.TreeNodeIcons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("TreeNodeIcons.ImageStream")));
            this.TreeNodeIcons.TransparentColor = System.Drawing.Color.Transparent;
            this.TreeNodeIcons.Images.SetKeyName(0, "ClosedFolder");
            this.TreeNodeIcons.Images.SetKeyName(1, "OpenFolder");
            this.TreeNodeIcons.Images.SetKeyName(2, "ClassModule");
            this.TreeNodeIcons.Images.SetKeyName(3, "PrivateClass");
            this.TreeNodeIcons.Images.SetKeyName(4, "Option");
            this.TreeNodeIcons.Images.SetKeyName(5, "Implements");
            this.TreeNodeIcons.Images.SetKeyName(6, "StandardModule");
            this.TreeNodeIcons.Images.SetKeyName(7, "PrivateModule");
            this.TreeNodeIcons.Images.SetKeyName(8, "PublicField");
            this.TreeNodeIcons.Images.SetKeyName(9, "PrivateField");
            this.TreeNodeIcons.Images.SetKeyName(10, "FriendField");
            this.TreeNodeIcons.Images.SetKeyName(11, "PublicMethod");
            this.TreeNodeIcons.Images.SetKeyName(12, "FriendMethod");
            this.TreeNodeIcons.Images.SetKeyName(13, "PrivateMethod");
            this.TreeNodeIcons.Images.SetKeyName(14, "PublicProperty");
            this.TreeNodeIcons.Images.SetKeyName(15, "FriendProperty");
            this.TreeNodeIcons.Images.SetKeyName(16, "PrivateProperty");
            this.TreeNodeIcons.Images.SetKeyName(17, "PublicConst");
            this.TreeNodeIcons.Images.SetKeyName(18, "FriendConst");
            this.TreeNodeIcons.Images.SetKeyName(19, "PrivateConst");
            this.TreeNodeIcons.Images.SetKeyName(20, "PublicEnum");
            this.TreeNodeIcons.Images.SetKeyName(21, "FriendEnum");
            this.TreeNodeIcons.Images.SetKeyName(22, "PrivateEnum");
            this.TreeNodeIcons.Images.SetKeyName(23, "PublicEnumItem");
            this.TreeNodeIcons.Images.SetKeyName(24, "FriendEnumItem");
            this.TreeNodeIcons.Images.SetKeyName(25, "PrivateEnumItem");
            this.TreeNodeIcons.Images.SetKeyName(26, "PublicType.bmp");
            this.TreeNodeIcons.Images.SetKeyName(27, "FriendType");
            this.TreeNodeIcons.Images.SetKeyName(28, "PrivateType");
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
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        public System.Windows.Forms.TreeView SolutionTree;
        private System.Windows.Forms.ImageList TreeNodeIcons;
    }
}
