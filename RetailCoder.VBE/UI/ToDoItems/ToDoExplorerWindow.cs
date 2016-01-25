using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Rubberduck.ToDoItems;
using Rubberduck.UI.CodeInspections;

namespace Rubberduck.UI.ToDoItems
{
    [ExcludeFromCodeCoverage]
    public partial class ToDoExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "8B071EDA-2C9C-4009-9A22-A1958BF98B28";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.ToDoExplorer_Caption; } }

        public ToDoExplorerWindow()
        {
            InitializeComponent();
        }

        private ToDoExplorerViewModel _viewModel;
        public ToDoExplorerViewModel ViewModel
        {
            get { return _viewModel; }
            set
            {
                _viewModel = value;
                this.toDoExplorerControl1.DataContext = _viewModel;
                if (_viewModel != null)
                {
                    _viewModel.RefreshCommand.Execute(null);
                }
            }
        }

    }
}
