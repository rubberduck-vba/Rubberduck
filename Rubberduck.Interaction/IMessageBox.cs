using System;
using Forms = System.Windows.Forms;

namespace Rubberduck.Interaction
{
    public interface IMessageBox
    {
        void Notify(string text);
        void NotifyError(string text, string caption);
        void NotifyWarn(string text, string caption);
        bool Question(string text, string caption);
        bool Prompt(string text, string caption);
        bool Confirm(string text, string caption);
    }

    public class MessageBox : IMessageBox
    {
        public void Notify(string text)
        {
            Forms.MessageBox.Show(text);
        }

        public void NotifyError(string text, string caption)
        {
            Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
        }

        public void NotifyWarn(string text, string caption)
        {
            Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
        }

        public bool Question(string text, string caption)
        {
            return Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.YesNo, Forms.MessageBoxIcon.Question) == Forms.DialogResult.Yes;
        }

        // FIXME check whether to merge with Question
        public bool Prompt(string text, string caption)
        {
            return Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.YesNo, Forms.MessageBoxIcon.Information) == Forms.DialogResult.Yes;
        }

        public bool Confirm(string text, string caption)
        {
            return Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.YesNo, Forms.MessageBoxIcon.Exclamation) == Forms.DialogResult.Yes;
        }
    }
}
