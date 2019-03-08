using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.CodeMetrics;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class CodeMetricsCommand : ComCommandBase
    {
        private readonly CodeMetricsDockablePresenter _presenter;

        public CodeMetricsCommand(CodeMetricsDockablePresenter presenter, IVbeEvents vbeEvents)
            : base(LogManager.GetCurrentClassLogger(), vbeEvents)
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
