using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete.Service
{
    public interface IShowIntelliSenseCommand
    {
        void Execute();
    }

    public class ShowIntelliSenseCommand : CommandBase, IShowIntelliSenseCommand
    {
        private readonly IVBE _vbe;

        public ShowIntelliSenseCommand(IVBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        public void Execute()
        {
            OnExecute(null);
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            try
            {
                return parameter != null || _vbe.ProjectsCount == 1;
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            const int showIntelliSenseId = 2531;
            using (var commandBars = _vbe.CommandBars)
            {
                using (var command = commandBars.FindControl(showIntelliSenseId))
                {
                    command.Execute();
                }
            }
        }
    }
}
