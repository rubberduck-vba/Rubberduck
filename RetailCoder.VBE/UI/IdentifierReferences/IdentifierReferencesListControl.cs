using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.IdentifierReferences
{
    public partial class IdentifierReferencesListControl : UserControl, IDockableUserControl
    {
        public IdentifierReferencesListControl()
        {
            InitializeComponent();
        }

        public string IdentifierName { get; set; }

        private const string ClassId = "972A7CE8-55A0-48F5-B607-2035E81D28CF";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return string.Format(RubberduckUI.AllReferences_Caption, IdentifierName); } }
    }
}
