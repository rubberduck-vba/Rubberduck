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
        private readonly ITestEngine testEngine;

        public TestExplorerModel(IVBE vbe, ITestEngine testEngine)
        {
            _vbe = vbe;
            this.testEngine = testEngine;

            testEngine.TestsRefreshed += HandleTestsRefreshed;
            testEngine.TestRunCompleted += HandleRunCompletion;
            testEngine.TestCompleted += HandleTestCompletion;
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
            Tests.Clear();
            foreach (var test in testEngine.Tests.Select(test => new TestMethodViewModel(test)))
            {
                Tests.Add(test);
            }
            RefreshProgressBarColor();
        }

        private void RefreshProgressBarColor()
        {
            var overallOutcome = testEngine.CurrentAggregateOutcome;
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
            get { return _totalDuration; } private set
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
            if (testEngine != null)
            {
                testEngine.TestCompleted -= HandleTestCompletion;
                testEngine.TestsRefreshed -= HandleTestsRefreshed;
                testEngine.TestRunCompleted -= HandleRunCompletion;
            }
        }
    }
}