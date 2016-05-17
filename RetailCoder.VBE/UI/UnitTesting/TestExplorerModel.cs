using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerModel : ViewModelBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly Dispatcher _uiDispatcher;

        public TestExplorerModel(VBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;
            _state.StateChanged += State_StateChanged;

            _uiDispatcher = Dispatcher.CurrentDispatcher;
        }

        private void State_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready) { return; }

            var tests = UnitTestHelpers.GetAllTests(_vbe, _state).ToList();

            UpdateTestList addTest = AddTest;
            UpdateTestList removeTest = RemoveTest;

            // add new tests
            foreach (var test in tests)
            {
                if (!_tests.Any(t =>
                    t.Declaration.ComponentName == test.Declaration.ComponentName &&
                    t.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                    t.Declaration.ProjectId == test.Declaration.ProjectId))
                {
                    _uiDispatcher.Invoke(addTest, test);
                }
            }

            var removedTests =_tests.Where(test =>
                        !tests.Any(t =>
                                t.Declaration.ComponentName == test.Declaration.ComponentName &&
                                t.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                                t.Declaration.ProjectId == test.Declaration.ProjectId)).ToList();

            // remove old tests
            foreach (var test in removedTests)
            {
                _uiDispatcher.Invoke(removeTest, test);
            }
        }

        private delegate void UpdateTestList(TestMethod test);
        private void AddTest(TestMethod test)
        {
            _tests.Add(test);
        }

        private void RemoveTest(TestMethod test)
        {
            _tests.Remove(test);
        }

        private readonly ObservableCollection<TestMethod> _tests = new ObservableCollection<TestMethod>();
        public ObservableCollection<TestMethod> Tests { get { return _tests; } }

        private readonly List<TestMethod> _lastRun = new List<TestMethod>();
        public List<TestMethod> LastRun { get { return _lastRun; } } 

        public void ClearLastRun()
        {
            _lastRun.Clear();
        }

        public void AddExecutedTest(TestMethod test)
        {
            _lastRun.Add(test);
            ExecutedCount = _tests.Count(t => t.Result.Outcome != TestOutcome.Unknown);

            ProgressBarColor = _tests.Any(t => t.Result.Outcome == TestOutcome.Failed)
                ? Colors.Red
                : _tests.Any(t => t.Result.Outcome == TestOutcome.Inconclusive) 
                    ? Colors.Gold
                    : Colors.LimeGreen;

            if (!_tests.Any(t =>
                        t.Declaration.ComponentName == test.Declaration.ComponentName &&
                        t.Declaration.IdentifierName == test.Declaration.IdentifierName &&
                        t.Declaration.ProjectId == test.Declaration.ProjectId))
            {
                _tests.Add(test);
            }

            // ReSharper disable once ExplicitCallerInfoArgument
            OnPropertyChanged("Tests");
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
    }
}