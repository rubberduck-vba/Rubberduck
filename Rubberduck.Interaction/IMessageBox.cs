using System;
using Forms = System.Windows.Forms;

namespace Rubberduck.Interaction
{
    public interface IMessageBox
    {
        void Message(string text);
        void NotifyError(string text, string caption);
        void NotifyWarn(string text, string caption);
        bool Question(string text, string caption);
        bool ConfirmYesNo(string text, string caption);
        bool ConfirmYesNo(string text, string caption, bool suggestion);
        bool? Confirm(string text, string caption, bool? suggestion);
    }

    public class MessageBox : IMessageBox
    {
        public void Message(string text)
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

        public bool ConfirmYesNo(string text, string caption)
        {
            return Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.YesNo, Forms.MessageBoxIcon.Exclamation) == Forms.DialogResult.Yes;
        }

        public bool ConfirmYesNo(string text, string caption, bool suggestion)
        {
            return Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.YesNo, Forms.MessageBoxIcon.Exclamation, suggestion ? Forms.MessageBoxDefaultButton.Button1 : Forms.MessageBoxDefaultButton.Button2) == Forms.DialogResult.Yes;
        }

        public bool? Confirm(string text, string caption, bool? suggestion)
        {
            var suggestionButton = suggestion.HasValue ? (suggestion.Value ? Forms.MessageBoxDefaultButton.Button1 : Forms.MessageBoxDefaultButton.Button2) : Forms.MessageBoxDefaultButton.Button3;
            var result = Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.YesNoCancel, Forms.MessageBoxIcon.Exclamation, suggestionButton);

            switch (result)
            {
                case Forms.DialogResult.Cancel:
                    return null;
                case Forms.DialogResult.Yes:
                    return true;
                case Forms.DialogResult.No:
                    return false;
                default:
                    return suggestion;

            }
        }
    }
}
