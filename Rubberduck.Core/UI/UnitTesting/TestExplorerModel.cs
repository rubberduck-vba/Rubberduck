using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using NLog;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    internal class TestExplorerModel : ViewModelBase, IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public event EventHandler<TestCompletedEventArgs> TestCompleted;
        private readonly ITestEngine _testEngine;

        public TestExplorerModel(ITestEngine testEngine)
        {
            _testEngine = testEngine;

            _testEngine.TestsRefreshStarted += HandleTestRefreshStarted;
            _testEngine.TestsRefreshed += HandleTestsRefreshed;
            _testEngine.TestRunStarted += HandleTestRunStarted;
            _testEngine.TestStarted += HandleTestStarted;
            _testEngine.TestRunCompleted += HandleRunCompletion;
            _testEngine.TestCompleted += HandleTestCompletion;
        }


        public ObservableCollection<TestMethodViewModel> Tests { get; } = new ObservableCollection<TestMethodViewModel>();

        private int _runCount;
        public int CurrentRunTestCount
        {
            get => _runCount;
            set
            {
                _runCount = value;
                OnPropertyChanged();
            }
        }

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

        private bool _isRefreshing;
        public bool IsRefreshing
        {
            get => _isRefreshing;
            set
            {
                _isRefreshing = value;
                OnPropertyChanged();
            }
        }

        public string LastTestRunSummary =>
            string.Format(Resources.UnitTesting.TestExplorer.TestOutcome_RunSummaryFormat, CurrentRunTestCount, Tests.Count, TimeSpan.FromMilliseconds(TotalDuration));

        public int LastTestFailedCount => Tests.Count(test => test.Result.Outcome == TestOutcome.Failed && _testEngine.LastRunTests.Contains(test.Method));

        public int LastTestInconclusiveCount => Tests.Count(test => test.Result.Outcome == TestOutcome.Inconclusive && _testEngine.LastRunTests.Contains(test.Method));

        public int LastTestIgnoredCount => Tests.Count(test => test.Result.Outcome == TestOutcome.Ignored && _testEngine.LastRunTests.Contains(test.Method));

        public int LastTestSucceededCount => Tests.Count(test => test.Result.Outcome == TestOutcome.Succeeded && _testEngine.LastRunTests.Contains(test.Method));

        public void ExecuteTests(IReadOnlyCollection<TestMethodViewModel> tests)
        {
            if (!tests.Any())
            {
                return;
            }

            IsBusy = true;

            try
            {
                _testEngine.Run(tests.Select(test => test.Method));
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "Test engine exception caught in TestExplorerModel.");
                IsBusy = false;
            }
        }

        public void CancelTestRun()
        {
            _testEngine.RequestCancellation();
        }

        private void HandleTestRunStarted(object sender, TestRunStartedEventArgs e)
        {
            if (e.Tests is null)
            {
                return;
            }

            var running = Tests.Where(test => e.Tests.Contains(test.Method)).ToList();

            // This also needs to be set in the handler - the test engine has other entry points.
            IsBusy = true;
            UpdateProgressBar(TestOutcome.Unknown, true);
            CurrentRunTestCount = running.Count;

            foreach (var test in running)
            {
                test.RunState = TestRunState.Queued;
            }
        }

        private void HandleTestStarted(object sender, TestStartedEventArgs e)
        {
            var running = Tests.FirstOrDefault(test => test?.Method?.Equals(e.Test) ?? false);

            if (running is null)
            {
                Logger.Warn($"{(e.Test is null ? "Null" : "Unexpected")} test result handled by TestExplorerModel.");
                return;
            }

            running.RunState = TestRunState.Running;
        }

        private void HandleRunCompletion(object sender, TestRunCompletedEventArgs e)
        {
            TotalDuration = e.Duration;

            foreach (var test in Tests)
            {
                test.RunState = TestRunState.Stopped;
            }

            ExecutedCount = Tests.Count(t => t.Result.Outcome != TestOutcome.Unknown);
            IsBusy = false;

            OnPropertyChanged(nameof(LastTestRunSummary));
            OnPropertyChanged(nameof(LastTestFailedCount));
            OnPropertyChanged(nameof(LastTestIgnoredCount));
            OnPropertyChanged(nameof(LastTestInconclusiveCount));
            OnPropertyChanged(nameof(LastTestSucceededCount));
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

            vmTest.RunState = TestRunState.Stopped;
            vmTest.Result = e.Result;
            
            UpdateProgressBar(vmTest.Result.Outcome);
            // Propagate the event.
            OnTestCompleted(e);
        }

        private void HandleTestRefreshStarted(object sender, EventArgs args)
        {
            IsRefreshing = true;
        }

        private void HandleTestsRefreshed(object sender, EventArgs args)
        {
            IsRefreshing = false;

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
        }

        private void OnTestCompleted(TestCompletedEventArgs args)
        {
            TestCompleted?.Invoke(this, args);
        }

        private static readonly Dictionary<TestOutcome, Color> OutcomeColors = new Dictionary<TestOutcome, Color>
        {
            { TestOutcome.Unknown, Colors.DimGray },
            { TestOutcome.Succeeded, Colors.LimeGreen },
            { TestOutcome.Inconclusive, Colors.Gold },
            { TestOutcome.Ignored, Colors.Orange },
            { TestOutcome.Failed, Colors.Red },
        };

        private TestOutcome _worstOutcome = TestOutcome.Unknown;
        private void UpdateProgressBar(TestOutcome output, bool reset = false)
        {
            if (reset)
            {
                _worstOutcome = TestOutcome.Unknown;
                ExecutedCount = 0;
                CurrentRunTestCount = 0;
                ProgressBarColor = OutcomeColors[TestOutcome.Unknown];
                return;
            }

            ExecutedCount++;

            if (_worstOutcome != TestOutcome.Unknown && output >= _worstOutcome)
            {
                return;
            }

            _worstOutcome = output;
            ProgressBarColor = OutcomeColors[output];
        }

        public void Dispose()
        {
            if (_testEngine != null)
            {
                _testEngine.TestsRefreshStarted -= HandleTestRefreshStarted;
                _testEngine.TestRunStarted -= HandleTestRunStarted;
                _testEngine.TestCompleted -= HandleTestCompletion;
                _testEngine.TestStarted -= HandleTestStarted;
                _testEngine.TestsRefreshed -= HandleTestsRefreshed;
                _testEngine.TestRunCompleted -= HandleRunCompletion;
            }
        }
    }
}