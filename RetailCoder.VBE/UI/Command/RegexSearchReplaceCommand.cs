using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.RegexSearchReplace;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class RegexSearchReplaceCommand : CommandBase
    {
        private readonly RegexSearchReplacePresenter _presenter;

        public RegexSearchReplaceCommand(RegexSearchReplacePresenter presenter) : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        protected override void ExecuteImpl(object parameter)
        {
            //_presenter.Show();
        }
    }
}
