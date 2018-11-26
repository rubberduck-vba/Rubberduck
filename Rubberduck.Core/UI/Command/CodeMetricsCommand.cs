using NLog;
using Rubberduck.UI.CodeMetrics;
using System.Runtime.InteropServices;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class CodeMetricsCommand : CommandBase
    {
        private readonly CodeMetricsDockablePresenter _presenter;

        public CodeMetricsCommand(CodeMetricsDockablePresenter presenter)
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
