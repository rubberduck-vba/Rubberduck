namespace Rubberduck.UI.RegexAssistant
{
    partial class RegexAssistantDialog
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
            this.ElementHost = new System.Windows.Forms.Integration.ElementHost();
            this.RegexAssistant = new Rubberduck.UI.RegexAssistant.RegexAssistant();
            this.SuspendLayout();
            // 
            // AssistantControl
            // 
            this.ElementHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ElementHost.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ElementHost.Location = new System.Drawing.Point(0, 0);
            this.ElementHost.Name = "AssistantControl";
            this.ElementHost.Size = new System.Drawing.Size(509, 544);
            this.ElementHost.TabIndex = 0;
            this.ElementHost.Child = this.RegexAssistant;
            // 
            // RegexAssistantDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(509, 544);
            this.Controls.Add(this.ElementHost);
            this.Name = "RegexAssistantDialog";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost ElementHost;
        private RegexAssistant RegexAssistant;
    }
}