using System.Runtime.InteropServices;
using EasyHook;
using NLog;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.VbeRuntime;

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
        private readonly IComMessagePumper _pumper;
        private readonly IVbeNativeApi _vbeApi;

        public ShowQuickInfoCommand(
            IVBE vbe, 
            IVBEEvents vbeEvents, 
            IVbeNativeApi vbeApi, 
            IComMessagePumper pumper) 
            : base(LogManager.GetCurrentClassLogger(), vbeEvents)
        {
            _vbe = vbe;
            _vbeApi = vbeApi;
            _pumper = pumper;
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
                using (HookVbaBeep())
                {
                    command.Execute();

                    // Ensures that the queued beep message (if any) is cleared
                    // while we have it hooked
                    _pumper.PumpMessages();
                }
            }
        }

        private LocalHook HookVbaBeep()
        {
            var processAddress = LocalHook.GetProcAddress(_vbeApi.DllName, "rtcBeep");
            var callbackDelegate = new VbaBeepDelegate(VbaBeepCallback);
            var hook = LocalHook.Create(processAddress, callbackDelegate, null);
            hook.ThreadACL.SetInclusiveACL(new[] {0});
            return hook;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void VbaBeepDelegate();

        public void VbaBeepCallback()
        {
            //Ignore the beep command
        }
    }
}
