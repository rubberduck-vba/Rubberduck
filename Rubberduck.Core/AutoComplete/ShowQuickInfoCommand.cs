using System.Runtime.InteropServices;
using Rubberduck.Runtime;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete
{
    public interface IShowQuickInfoCommand
    {
        /// <summary>
        /// Displays the quick info in the VBE.
        /// NOTE: By default, it makes an utterly annoying DING! in the VBE if the "QuickInfo" command is unavailable. Hooking is used to suppress the ding.
        /// </summary>
        void Execute();
    }

    public class ShowQuickInfoCommand : ComCommandBase, IShowQuickInfoCommand
    {
        private readonly IVBE _vbe;
        private readonly IBeepInterceptor _beepInterceptor;

        public ShowQuickInfoCommand(IVBE vbe, IVbeEvents vbeEvents, IBeepInterceptor beepInterceptor) : base(vbeEvents)
        {
            _vbe = vbe;
            _beepInterceptor = beepInterceptor;
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
            const int showQuickInfoId = 2531;
            using (var commandBars = _vbe.CommandBars)
            {
                using (var command = commandBars.FindControl(showQuickInfoId))
                {
                    // Ensures that the queued beep message (if any) is suppressed
                    _beepInterceptor.SuppressBeep(100);

                    command.Execute();
                }
            }
        }
    }
}