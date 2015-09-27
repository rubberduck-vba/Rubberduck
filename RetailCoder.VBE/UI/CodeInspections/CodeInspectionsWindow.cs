using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.Properties;

namespace Rubberduck.UI.CodeInspections
{
    public partial class CodeInspectionsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "D3B2A683-9856-4246-BDC8-6B0795DC875B";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeInspections; } }
        
        public CodeInspectionsWindow()
            : this(null)
        {
        }

        public CodeInspectionsWindow(InspectionResultsViewModel viewModel)
        {
            InitializeComponent();
            inspectionResultsControl1.DataContext = viewModel;
        }
    }
}
