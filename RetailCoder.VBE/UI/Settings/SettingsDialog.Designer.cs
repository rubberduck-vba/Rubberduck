namespace Rubberduck.UI.Settings
{
    partial class SettingsDialog
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.configurationTreeView1 = new Rubberduck.UI.Settings.ConfigurationTreeView();
            this.todoListSettingsControl1 = new Rubberduck.UI.Settings.TodoListSettingsControl();
            this.SuspendLayout();
            // 
            // configurationTreeView1
            // 
            this.configurationTreeView1.Location = new System.Drawing.Point(12, 12);
            this.configurationTreeView1.Name = "configurationTreeView1";
            this.configurationTreeView1.Size = new System.Drawing.Size(259, 256);
            this.configurationTreeView1.TabIndex = 0;
            this.configurationTreeView1.Load += new System.EventHandler(this.configurationTreeView1_Load);
            // 
            // todoListSettingsControl1
            // 
            this.todoListSettingsControl1.Location = new System.Drawing.Point(277, 2);
            this.todoListSettingsControl1.Name = "todoListSettingsControl1";
            this.todoListSettingsControl1.Size = new System.Drawing.Size(530, 294);
            this.todoListSettingsControl1.TabIndex = 1;
            // 
            // SettingsDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(805, 291);
            this.Controls.Add(this.todoListSettingsControl1);
            this.Controls.Add(this.configurationTreeView1);
            this.Name = "SettingsDialog";
            this.Text = "SettingsDialog";
            this.ResumeLayout(false);

        }

        #endregion

        private ConfigurationTreeView configurationTreeView1;
        private TodoListSettingsControl todoListSettingsControl1;
    }
}