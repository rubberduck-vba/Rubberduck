
using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class FailedActionControl : UserControl, IFailedMessageView
    {
        public FailedActionControl()
        {
            InitializeComponent();
            DismissMessageButton.Click += DismissMessageButton_Click;
        }

        public string Message
        {
            get { return ActionFailedMessage.Text; }
            set { ActionFailedMessage.Text = value; }
        }

        public event EventHandler<EventArgs> DismissSecondaryPanel;
        void DismissMessageButton_Click(object sender, EventArgs e)
        {
            var handler = DismissSecondaryPanel;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
