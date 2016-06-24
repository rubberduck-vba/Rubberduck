using System;
using System.Diagnostics;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all Rubberduck unit tests in the VBE.
    /// </summary>
    public class RunAllTestsCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly ITestEngine _engine;
        private readonly TestExplorerModel _model;
        private readonly RubberduckParserState _state;
        
        public RunAllTestsCommand(VBE vbe, RubberduckParserState state, ITestEngine engine, TestExplorerModel model)
        {
            _vbe = vbe;
            _engine = engine;
            _model = model;
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            return _vbe.IsInDesignMode();
        }

        public override void Execute(object parameter)
        {
            if (!_state.IsDirty())
            {
                RunTests();
            }
            else
            {
                _model.TestsRefreshed += TestsRefreshed;
                _model.Refresh();
            }
        }

        private void TestsRefreshed(object sender, EventArgs e)
        {
            RunTests();
        }

        private void RunTests()
        {
            _model.TestsRefreshed -= TestsRefreshed;

            var stopwatch = new Stopwatch();

            _model.ClearLastRun();
            _model.IsBusy = true;

            stopwatch.Start();
            _engine.Run(_model.Tests);
            stopwatch.Stop();

            _model.IsBusy = false;

            OnRunCompleted(new TestRunEventArgs(stopwatch.ElapsedMilliseconds));
        }

        public event EventHandler<TestRunEventArgs> RunCompleted;
        protected virtual void OnRunCompleted(TestRunEventArgs e)
        {
            var handler = RunCompleted;
            if (handler != null)
            {
                handler.Invoke(this, e);
            }
        }
    }
    
    public class TestRunEventArgs : EventArgs
    {
        public long Duration { get; private set; }

        public TestRunEventArgs(long duration)
        {
            Duration = duration;
        }
    }
}
