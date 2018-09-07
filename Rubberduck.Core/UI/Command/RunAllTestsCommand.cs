using NLog;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all Rubberduck unit tests in the VBE.
    /// </summary>
    public class RunAllTestsCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly ITestEngine _engine;

        public RunAllTestsCommand(IVBE vbe, ITestEngine engine)
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _engine = engine;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            // the vbe design mode requirement could also be encapsulated into the engine
            return _vbe.IsInDesignMode && _engine.CanRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (_engine.CanRun)
            {
                _engine.Run(_engine.Tests);
            }
        }
    }
}
