using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public partial class CodeInspectionsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "D3B2A683-9856-4246-BDC8-6B0795DC875B";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return "Code Inspections"; } }

        public CodeInspectionsWindow()
        {
            InitializeComponent();
            RefreshButton.Click += RefreshButtonClicked;
        }

        public event EventHandler RefreshCodeInspections;
        private void RefreshButtonClicked(object sender, EventArgs e)
        {
            var handler = RefreshCodeInspections;
            if (handler == null)
            {
                return;
            }

            handler(this, EventArgs.Empty);
        }

        public event EventHandler NavigateCodeIssue;
    }
}
