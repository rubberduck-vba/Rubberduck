using System.Windows.Input;
using Rubberduck.Refactorings.RemoveParameters;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public partial class RemoveParametersView
    {
        public RemoveParametersView()
        {
            InitializeComponent();
        }

        private RemoveParametersViewModel ViewModel => (RemoveParametersViewModel) DataContext;

        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount != 2) { return; }

            var target = (Parameter)ParameterGrid.SelectedItem;
            if (target.IsRemoved)
            {
                ViewModel.RestoreParameterCommand.Execute(target);
            }
            else
            {
                ViewModel.RemoveParameterCommand.Execute(target);
            }
        }
    }
}
