using System.Windows.Controls;
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

        private void ListBoxItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (ListBoxItem)sender;
            var target = (Parameter)item.Content;

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
