using System;
using System.Diagnostics;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using System.Linq;
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
        private readonly TestExplorerModel _model;
        private readonly IDockablePresenter _presenter;
        private readonly RubberduckParserState _state;

        public RunAllTestsCommand(IVBE vbe, RubberduckParserState state, ITestEngine engine, TestExplorerModel model, IDockablePresenter presenter) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _engine = engine;
            _model = model;
            _state = state;
            _presenter = presenter;
        }

        private static readonly ParserState[] AllowedRunStates = { ParserState.ResolvedDeclarations, ParserState.ResolvingReferences, ParserState.Ready };

        protected override bool CanExecuteImpl(object parameter)
        {
            return _vbe.IsInDesignMode && AllowedRunStates.Contains(_state.Status);
        }

        protected override void ExecuteImpl(object parameter)
        {
            EnsureRubberduckIsReferencedForEarlyBoundTests();

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

        private void EnsureRubberduckIsReferencedForEarlyBoundTests()
        {
            foreach (var member in _state.AllUserDeclarations)
            {
                if (member.AsTypeName == "Rubberduck.PermissiveAssertClass" ||
                    member.AsTypeName == "Rubberduck.AssertClass")
                {
                    member.Project.EnsureReferenceToAddInLibrary();
                }
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

            _presenter?.Show();

            stopwatch.Start();
            try
            {
                _engine.Run(_model.Tests);
            }
            finally
            {
                stopwatch.Stop();
                _model.IsBusy = false;
            }

            Logger.Info($"Test run completed in {stopwatch.ElapsedMilliseconds}.");
            OnRunCompleted(new TestRunEventArgs(stopwatch.ElapsedMilliseconds));
        }

        public event EventHandler<TestRunEventArgs> RunCompleted;
        protected virtual void OnRunCompleted(TestRunEventArgs e)
        {
            var handler = RunCompleted;
            handler?.Invoke(this, e);
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
