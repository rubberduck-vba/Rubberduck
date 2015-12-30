using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;

namespace Rubberduck.UI.CodeExplorer
{
    [ExcludeFromCodeCoverage]
    public partial class CodeExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B59-172F-417C-88E3-B377CDA2D809";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeExplorerDockablePresenter_Caption; } }

        public CodeExplorerWindow()
        {
            InitializeComponent();
        }

    }
}
