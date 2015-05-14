using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public partial class ReorderParametersDialog : Form, IReorderParametersView
    {
        public ReorderParametersDialog()
        {
            InitializeComponent();

            OkButton.Click += OkButtonClicked;
            button1.Click += button1_Click;
            button2.Click += button2_Click;
        }

        public event EventHandler CancelButtonClicked;
        public void OnCancelButtonClicked()
        {
            Hide();
        }

        public event EventHandler OkButtonClicked;
        public void OnOkButtonClicked()
        {
            var handler = OkButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private Declaration _target;
        public Declaration Target
        {
            get { return _target; }
            set
            {
                _target = value;
                if (_target == null)
                {
                    return;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // TODO - implement move up functionality
            // simple swap should do it
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // TODO - implement move down functionality
        }
    }
}
