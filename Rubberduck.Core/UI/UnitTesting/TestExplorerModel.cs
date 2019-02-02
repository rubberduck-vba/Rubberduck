using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using System.Windows.Threading;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting
{
    internal class TestExplorerModel : ViewModelBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly Dispatcher _dispatcher;
        private readonly ITestEngine _testEngine;

        public TestExplorerModel(IVBE vbe, ITestEngine testEngine)
        {
            _vbe = vbe;
            _testEngine = testEngine;

            _testEngine.TestsRefreshed += HandleTestsRefreshed;
            _testEngine.TestRunCompleted += HandleRunCompletion;
            _testEngine.TestCompleted += HandleTestCompletion;
            _dispatcher = Dispatcher.CurrentDispatcher;
        }

        private void HandleRunCompletion(object sender, TestRunCompletedEventArgs e)
        {
            TotalDuration = e.Duration;
            ExecutedCount = Tests.Count(t => t.Result.Outcome != TestOutcome.Unknown);
        }

        private void HandleTestCompletion(object sender, TestCompletedEventArgs e)
        {
            var test = e.Test;
            var vmTest = new TestMethodViewModel(test);

            if (!Tests.Contains(vmTest))
            {
                Tests.Add(vmTest);
            }
            else
            {
                vmTest = Tests.First(inside => inside.Equals(vmTest));
            }
            vmTest.Result = e.Result;

            RefreshProgressBarColor();
        }

        private void HandleTestsRefreshed(object sender, EventArgs args)
        {
            var previous = Tests.ToList();
            
            Tests.Clear();
            foreach (var test in _testEngine.Tests)
            {
                var adding = new TestMethodViewModel(test);
                var match = previous.FirstOrDefault(ut => ut.Method.Equals(test));
                if (match != null)
                {
                    adding.Result = match.Result;
                }
                Tests.Add(adding);
            }
            RefreshProgressBarColor();
        }

        private void RefreshProgressBarColor()
        {
            var overallOutcome = _testEngine.CurrentAggregateOutcome;
            switch (overallOutcome)
            {
                case TestOutcome.Failed:
                    ProgressBarColor = Colors.Red;
                    break;
                case TestOutcome.Inconclusive:
                    ProgressBarColor = Colors.Gold;
                    break;
                case TestOutcome.Succeeded:
                    ProgressBarColor = Colors.LimeGreen;
                    break;
                default:
                    ProgressBarColor = Colors.DimGray;
                    break;
            }
        }
        
        public ObservableCollection<TestMethodViewModel> Tests { get; } = new ObservableCollection<TestMethodViewModel>();
        
        private long _totalDuration;
        public long TotalDuration
        {
            get => _totalDuration;
            private set
            {
                _totalDuration = value;
                OnPropertyChanged();
            }
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
            }
        }

        public void Dispose()
        {
            if (_testEngine != null)
            {
                _testEngine.TestCompleted -= HandleTestCompletion;
                _testEngine.TestsRefreshed -= HandleTestsRefreshed;
                _testEngine.TestRunCompleted -= HandleRunCompletion;
            }
        }
    }
}