using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using System.Windows.Threading;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerModel : ViewModelBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly Dispatcher _dispatcher;

        public TestExplorerModel(IVBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;
            _state.StateChanged += HandleStateChanged;

            _dispatcher = Dispatcher.CurrentDispatcher;
        }

        private void HandleStateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.ResolvedDeclarations) { return; }

            _dispatcher.Invoke(() =>
            {
                var tests = UnitTestUtils.GetAllTests(_vbe, _state).ToList();

                var removedTests = Tests.Where(test =>
                             !tests.Any(t =>
                                     t.Declaration.ComponentName == test.Declaration.ComponentName &&
                                     t.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                                     t.Declaration.ProjectId == test.Declaration.ProjectId)).ToList();

                // remove old tests
                foreach (var test in removedTests)
                {
                    Tests.Remove(test);
                }

                // update declarations for existing tests--declarations are immutable
                foreach (var test in Tests.Except(removedTests))
                {
                    var declaration = tests.First(t =>
                        t.Declaration.ComponentName == test.Declaration.ComponentName &&
                        t.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                        t.Declaration.ProjectId == test.Declaration.ProjectId).Declaration;

                    test.SetDeclaration(declaration);
                }

                // add new tests
                foreach (var test in tests)
                {
                    if (!Tests.Any(t =>
                        t.Declaration.ComponentName == test.Declaration.ComponentName &&
                        t.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                        t.Declaration.ProjectId == test.Declaration.ProjectId))
                    {
                        Tests.Add(test);
                    }
                }
            });

            OnTestsRefreshed();
        }

        public ObservableCollection<TestMethod> Tests { get; } = new ObservableCollection<TestMethod>();

        public List<TestMethod> LastRun { get; } = new List<TestMethod>();

        public void ClearLastRun()
        {
            LastRun.Clear();
        }

        public void AddExecutedTest(TestMethod test)
        {
            LastRun.Add(test);
            ExecutedCount = Tests.Count(t => t.Result.Outcome != TestOutcome.Unknown);

            if (Tests.Any(t => t.Result.Outcome == TestOutcome.Failed))
            {
                ProgressBarColor = Colors.Red;
            }
            else
            {
                ProgressBarColor = Tests.Any(t => t.Result.Outcome == TestOutcome.Inconclusive)
                    ? Colors.Gold
                    : Colors.LimeGreen;
            }

            if (!Tests.Any(t =>
                        t.Declaration.ComponentName == test.Declaration.ComponentName &&
                        t.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                        t.Declaration.ProjectId == test.Declaration.ProjectId))
            {
                Tests.Add(test);
            }
        }

        public void Refresh()
        {
            _state.OnParseRequested(this);
        }

        private int _executedCount;
        public int ExecutedCount
        {
            get { return _executedCount; }
            private set
            {
                _executedCount = value;
                OnPropertyChanged();
            }
        }

        private Color _progressBarColor = Colors.DimGray;
        public Color ProgressBarColor
        {
            get { return _progressBarColor; }
            set
            {
                _progressBarColor = value;
                OnPropertyChanged();
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value;
                OnPropertyChanged();

                IsReady = !_isBusy;
            }
        }

        private bool _isReady = true;
        public bool IsReady
        {
            get { return _isReady; }
            private set
            {
                _isReady = value;
                OnPropertyChanged();
            }
        }

        public event EventHandler<EventArgs> TestsRefreshed;
        private void OnTestsRefreshed()
        {
            var handler = TestsRefreshed;
            handler?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= HandleStateChanged;
            }
        }
    }
}