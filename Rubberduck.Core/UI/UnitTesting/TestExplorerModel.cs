using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using System.Windows.Threading;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting
{

    public class TestExplorerModel : ViewModelBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly Dispatcher _dispatcher;
        private readonly ITestEngine testEngine;

        public TestExplorerModel(IVBE vbe, ITestEngine testEngine)
        {
            _vbe = vbe;
            this.testEngine = testEngine;

            testEngine.TestsRefreshed += RefreshTests;
            testEngine.TestRunCompleted += (_, time) => TotalDuration = time;
            testEngine.TestCompleted += HandleTestCompletion;
            _dispatcher = Dispatcher.CurrentDispatcher;
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
                vmTest = Tests.First(vm => vm.Method == test);
            }
            LastRun.Add(vmTest);
            vmTest.Result = e.Result;

            ExecutedCount = Tests.Count(t => t.Result.Outcome != TestOutcome.Unknown);

            RefreshProgressBarColor();
        }

        private void RefreshTests(object sender, EventArgs args)
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

        // FIXME can this really not be internal?
        public ObservableCollection<TestMethodViewModel> Tests { get; } = new ObservableCollection<TestMethodViewModel>();

        public List<TestMethodViewModel> LastRun { get; } = new List<TestMethodViewModel>();
        private long _totalDuration;
        public long TotalDuration
        {
            get { return _totalDuration; } private set
            {
                _totalDuration = value;
                OnPropertyChanged();
            }
        }

        public void ClearLastRun()
        {
            LastRun.Clear();
        }

        //public void AddExecutedTest(TestMethod test)
        //{
        //    if (!Tests.Any(t =>
        //    t.Method.Declaration.ComponentName == test.Declaration.ComponentName &&
        //    t.Method.Declaration.IdentifierName == test.Declaration.IdentifierName &&
        //    t.Method.Declaration.ProjectId == test.Declaration.ProjectId))
        //    {
        //        Tests.Add(new TestMethodViewModel(test));
        //    }

        //    LastRun.Add(Tests.First(m => m.Method == test));
        //    ExecutedCount = Tests.Count(t => t.Result.Outcome != TestOutcome.Unknown);

        //    if (Tests.Any(t => t.Result.Outcome == TestOutcome.Failed))
        //    {
        //        ProgressBarColor = Colors.Red;
        //    }
        //    else
        //    {
        //        ProgressBarColor = Tests.Any(t => t.Result.Outcome == TestOutcome.Inconclusive)
        //            ? Colors.Gold
        //            : Colors.LimeGreen;
        //    }
        //}
        
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
                testEngine.TestsRefreshed -= RefreshTests;
            }
        }
    }
}