using System.Runtime.InteropServices;
using EasyHook;
using NLog;
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

        public ShowQuickInfoCommand(
            IVBE vbe, 
            IVbeEvents vbeEvents, 
            IBeepInterceptor beepInterceptor) 
            : base(LogManager.GetCurrentClassLogger(), vbeEvents)
        {
            _vbe = vbe;
            _beepInterceptor = beepInterceptor;
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
                    // Ensures that the queued beep message (if any) is suppressed
                    _beepInterceptor.SuppressBeep(100);

                    command.Execute();
                }
            }
        }
    }
}
