using System.Linq;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all Rubberduck unit tests in the VBE.
    /// </summary>
    public class RunAllTestsCommand : CommandBase
    {
        private readonly ITestEngine _engine;

        public RunAllTestsCommand(ITestEngine engine)
        {
            _engine = engine;
        }

        public override void Execute(object parameter)
        {
            _engine.Run();
        }
    }
}