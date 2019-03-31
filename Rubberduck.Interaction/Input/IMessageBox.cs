using Forms = System.Windows.Forms;

namespace Rubberduck.Interaction
{
    public interface IMessageBox
    {
        /// <summary>
        /// Show a message to the user. Will only return after the user has acknowledged the message.
        /// </summary>
        /// <param name="text">The message to show to the user</param>
        void Message(string text);
        /// <summary>
        /// Notify the user of a warning. Will only return after the user has acknowledged the warning.
        /// </summary>
        /// <param name="text">The Warning text to show the user</param>
        /// <param name="caption">The caption of the dialog window</param>
        void NotifyWarn(string text, string caption);
        /// <summary>
        /// Ask the user a question. Neither user selection must have any non-reversible consequences.
        /// Will only return on user-input.
        /// </summary>
        /// <param name="text">The Question to ask the user</param>
        /// <param name="caption">The caption of the dialog window</param>
        /// <returns>true, if the user selects "Yes", false if the user selects "No"</returns>
        bool Question(string text, string caption);
        /// <summary>
        /// Ask the user for a simple confirmation. If the user selects an option, non-reversible consequences are acceptable.
        /// Will only return on user-input.
        /// </summary>
        /// <param name="text">The question to ask the user</param>
        /// <param name="caption">The caption of the dialog window</param>
        /// <param name="suggestion">The pre-selected result for the user, defaults to <b>Yes</b></param>
        /// <returns>true, if the user selects "Yes", false if the user selects "No"</returns>
        bool ConfirmYesNo(string text, string caption, bool suggestion = true);
        /// <summary>
        /// Ask the user for a confirmation. If the user selects an option that is not "Cancel", 
        /// non-reversible consequences are acceptable.
        /// Will only return on user-input.
        /// </summary>
        /// <param name="text">The question to ask the user</param>
        /// <param name="caption">The caption of the dialog window</param>
        /// <param name="suggestion">The pre-selected result for the user, defaults to <b>Cancel</b></param>
        /// <returns>Yes, No or Cancel respectively, according to the user's input</returns>
        ConfirmationOutcome Confirm(string text, string caption, ConfirmationOutcome suggestion = ConfirmationOutcome.Cancel);
    }

    public enum ConfirmationOutcome
    {
        Yes, No, Cancel
    }

    public class MessageBox : IMessageBox
    {
        public void Message(string text)
        {
            Forms.MessageBox.Show(text);
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

        public ConfirmationOutcome Confirm(string text, string caption, ConfirmationOutcome suggestion = ConfirmationOutcome.Yes)
        {
            Forms.MessageBoxDefaultButton suggestionButton;
            switch (suggestion)
            {
                // default required to shut the compiler up about "unassigned variable"
                default:
                case ConfirmationOutcome.Yes:
                    suggestionButton = Forms.MessageBoxDefaultButton.Button1;
                    break;
                case ConfirmationOutcome.No:
                    suggestionButton = Forms.MessageBoxDefaultButton.Button2;
                    break;
                case ConfirmationOutcome.Cancel:
                    suggestionButton = Forms.MessageBoxDefaultButton.Button3;
                    break;
            }
            var result = Forms.MessageBox.Show(text, caption, Forms.MessageBoxButtons.YesNoCancel, Forms.MessageBoxIcon.Exclamation, suggestionButton);

            switch (result)
            {
                case Forms.DialogResult.Cancel:
                    return ConfirmationOutcome.Cancel;
                case Forms.DialogResult.Yes:
                    return ConfirmationOutcome.Yes;
                case Forms.DialogResult.No:
                    return ConfirmationOutcome.No;
                default:
                    return suggestion;
            }
        }
    }
}
