using System.Runtime.InteropServices;
using Rubberduck.Navigation.RegexSearchReplace;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class RegexSearchReplaceCommand : CommandBase
    {
        private readonly RegexSearchReplacePresenter _presenter;

        public RegexSearchReplaceCommand(RegexSearchReplacePresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}