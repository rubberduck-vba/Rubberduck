using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.Settings
{
    partial class ConfigurationTreeViewControl
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfigurationTreeViewControl));
            this.settingsTreeView = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.SuspendLayout();
            // 
            // settingsTreeView
            // 
            this.settingsTreeView.BackColor = System.Drawing.Color.White;
            this.settingsTreeView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.settingsTreeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.settingsTreeView.ImageIndex = 0;
            this.settingsTreeView.ImageList = this.imageList1;
            this.settingsTreeView.LineColor = System.Drawing.Color.LightGray;
            this.settingsTreeView.Location = new System.Drawing.Point(0, 0);
            this.settingsTreeView.Margin = new System.Windows.Forms.Padding(10);
            this.settingsTreeView.Name = "settingsTreeView";
            this.settingsTreeView.SelectedImageIndex = 0;
            this.settingsTreeView.Size = new System.Drawing.Size(302, 314);
            this.settingsTreeView.TabIndex = 0;
            this.settingsTreeView.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.settingsTreeView_AfterSelect);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Default");
            this.imageList1.Images.SetKeyName(1, "Ducky");
            this.imageList1.Images.SetKeyName(2, "CodeInspections");
            this.imageList1.Images.SetKeyName(3, "Navigation");
            this.imageList1.Images.SetKeyName(4, "UnitTesting");
            // 
            // ConfigurationTreeViewControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.settingsTreeView);
            this.Name = "ConfigurationTreeViewControl";
            this.Size = new System.Drawing.Size(302, 314);
            this.ResumeLayout(false);

        }

        #endregion

        private TreeView settingsTreeView;
        private ImageList imageList1;

    }
}
