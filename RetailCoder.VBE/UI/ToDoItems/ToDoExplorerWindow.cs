using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

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

            TodoExplorerControl.TodoDoubleClick += toDoExplorerControl1_TodoDoubleClick;
        }

        void toDoExplorerControl1_TodoDoubleClick(object sender, EventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.NavigateToToDo.Execute(new NavigateCodeEventArgs(_viewModel.SelectedToDo.GetSelection()));
            }
        }

        private ToDoExplorerViewModel _viewModel;
        public ToDoExplorerViewModel ViewModel
        {
            get { return _viewModel; }
            set
            {
                _viewModel = value;
                TodoExplorerControl.DataContext = _viewModel;
                if (_viewModel != null)
                {
                    _viewModel.RefreshCommand.Execute(null);
                }
            }
        }

    }
}
