using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.RegexSearchReplace;
using Rubberduck.Parsing.Common;

namespace Rubberduck.UI.Command
{
#if !DEBUG
    [Experimental]
#endif
    [ComVisible(false)]
    public class RegexSearchReplaceCommand : CommandBase
    {
        private readonly RegexSearchReplacePresenter _presenter;

        public RegexSearchReplaceCommand(RegexSearchReplacePresenter presenter) : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            //_presenter.Show();
        }
    }
}
