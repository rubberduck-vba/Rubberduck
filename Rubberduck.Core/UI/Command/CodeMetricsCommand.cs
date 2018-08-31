using NLog;
using System.Runtime.InteropServices;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class CodeMetricsCommand : CommandBase
    {
        private readonly IDockablePresenter _presenter;

        public CodeMetricsCommand(IDockablePresenter presenter)
            : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
