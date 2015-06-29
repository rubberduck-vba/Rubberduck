namespace Rubberduck.UI.SourceControl
{
    partial class FailedActionControl
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.DismissMessageButton = new System.Windows.Forms.Button();
            this.ActionFailedMessage = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.DismissMessageButton);
            this.panel1.Controls.Add(this.ActionFailedMessage);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(243, 193);
            this.panel1.TabIndex = 3;
            // 
            // DismissMessageButton
            // 
            this.DismissMessageButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.DismissMessageButton.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.DismissMessageButton.Location = new System.Drawing.Point(0, 169);
            this.DismissMessageButton.Name = "DismissMessageButton";
            this.DismissMessageButton.Size = new System.Drawing.Size(243, 24);
            this.DismissMessageButton.TabIndex = 3;
            this.DismissMessageButton.Text = "Dismiss";
            this.DismissMessageButton.UseVisualStyleBackColor = true;
            // 
            // ActionFailedMessage
            // 
            this.ActionFailedMessage.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ActionFailedMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ActionFailedMessage.Location = new System.Drawing.Point(0, 0);
            this.ActionFailedMessage.Name = "ActionFailedMessage";
            this.ActionFailedMessage.Size = new System.Drawing.Size(243, 193);
            this.ActionFailedMessage.TabIndex = 2;
            // 
            // FailedActionControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.panel1);
            this.Name = "FailedActionControl";
            this.Size = new System.Drawing.Size(243, 193);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button DismissMessageButton;
        private System.Windows.Forms.Label ActionFailedMessage;

    }
}
