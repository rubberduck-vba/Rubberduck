using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test module to the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class AddTestModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly NewUnitTestModuleCommand _command;

        public AddTestModuleCommand(VBE vbe, NewUnitTestModuleCommand command) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _command = command;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _vbe.HostSupportsUnitTests();
        }

        protected override void ExecuteImpl(object parameter)
        {
            _command.NewUnitTestModule(_vbe.ActiveVBProject);
        }
    }
}
