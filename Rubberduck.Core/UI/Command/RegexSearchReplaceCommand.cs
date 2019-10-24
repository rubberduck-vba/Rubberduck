using System.Runtime.InteropServices;
using Rubberduck.Navigation.RegexSearchReplace;
using Rubberduck.Parsing.Common;

namespace Rubberduck.UI.Command
{
    [Disabled]
    [ComVisible(false)]
    public class RegexSearchReplaceCommand : CommandBase
    {
        private readonly RegexSearchReplacePresenter _presenter;

        public RegexSearchReplaceCommand(RegexSearchReplacePresenter presenter)
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            //_presenter.Show();
        }
    }
}
