using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using System.Windows.Threading;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting.ViewModels;
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
                Tests.Clear();
                foreach (var test in UnitTestUtils.GetAllTests(_state))
                {
                    // FIXME this shouldn't be necessary
                    if (!Tests.Any(t =>
                        t.Method.Declaration.ComponentName == test.Declaration.ComponentName &&
                        t.Method.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                        t.Method.Declaration.ProjectId == test.Declaration.ProjectId))
                    {
                        Tests.Add(new TestMethodViewModel(test));
                    }
                }
            });

            OnTestsRefreshed();
        }

        internal ObservableCollection<TestMethodViewModel> Tests { get; } = new ObservableCollection<TestMethodViewModel>();

        internal List<TestMethodViewModel> LastRun { get; } = new List<TestMethodViewModel>();

        public void ClearLastRun()
        {
            LastRun.Clear();
        }

        public void AddExecutedTest(TestMethod test)
        {
            if (!Tests.Any(t =>
            t.Method.Declaration.ComponentName == test.Declaration.ComponentName &&
            t.Method.Declaration.IdentifierName == test.Declaration.IdentifierName &&
            t.Method.Declaration.ProjectId == test.Declaration.ProjectId))
            {
                Tests.Add(new TestMethodViewModel(test));
            }

            LastRun.Add(Tests.First(m => m.Method == test));
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
        }

        public void Refresh()
        {
            _state.OnParseRequested(this);
        }

        private int _executedCount;
        public int ExecutedCount
        {
            get => _executedCount;
            private set
            {
                _executedCount = value;
                OnPropertyChanged();
            }
        }

        private Color _progressBarColor = Colors.DimGray;
        public Color ProgressBarColor
        {
            get => _progressBarColor;
            set
            {
                _progressBarColor = value;
                OnPropertyChanged();
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
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
            get => _isReady;
            private set
            {
                _isReady = value;
                OnPropertyChanged();
            }
        }

        public event EventHandler<EventArgs> TestsRefreshed;
        private void OnTestsRefreshed()
        {
            TestsRefreshed?.Invoke(this, EventArgs.Empty);
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