using System.Windows.Forms;

namespace Rubberduck.UI
{
    public interface IMessageBox
    {
        DialogResult Show(string text);
        DialogResult Show(IWin32Window owner, string text);
        DialogResult Show(string text, string caption);
        DialogResult Show(IWin32Window owner, string text, string caption);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, bool displayHelpButton);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, string keyword);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, string keyword);
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator, object param);
        DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator, object param);
    }

    public class MessageBox : IMessageBox
    {
        public DialogResult Show(string text)
        {
            return System.Windows.Forms.MessageBox.Show(text);
        }

        public DialogResult Show(IWin32Window owner, string text)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text);
        }

        public DialogResult Show(string text, string caption)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons, icon);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons, icon, defaultButton);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton, options);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons, icon, defaultButton, options);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, bool displayHelpButton)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton, options, displayHelpButton);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton, options, helpFilePath);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons, icon, defaultButton, options, helpFilePath);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, string keyword)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton, options, helpFilePath, keyword);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, string keyword)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons, icon, defaultButton, options, helpFilePath, keyword);
        }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator,
            object param)
        {
            return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator, param);
        }

        public DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton, MessageBoxOptions options, string helpFilePath, HelpNavigator navigator,
            object param)
        {
            return System.Windows.Forms.MessageBox.Show(owner, text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator, param);
        }
    }
}
