using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete
{
    public interface IShowIntelliSenseCommand
    {
        /// <summary>
        /// WARNING! Makes an utterly annoying DING! in the VBE if the "QuickInfo" command is unavailable.
        /// </summary>
        void Execute();
    }

    public class ShowIntelliSenseCommand : CommandBase, IShowIntelliSenseCommand
    {
        private readonly IVBE _vbe;

        public ShowIntelliSenseCommand(IVBE vbe)
        {
            _vbe = vbe;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
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

        public void Execute()
        {
            OnExecute(null);
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
