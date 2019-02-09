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

        private readonly ITestEngine _testEngine;

        public TestExplorerModel(ITestEngine testEngine)
        {
            _testEngine = testEngine;

            _testEngine.TestsRefreshed += HandleTestsRefreshed;
            _testEngine.TestRunStarted += HandleTestRunStarted;
            _testEngine.TestStarted += HandleTestStarted;
            _testEngine.TestRunCompleted += HandleRunCompletion;
            _testEngine.TestCompleted += HandleTestCompletion;
        }

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
            }

            IsBusy = false;
        }

        private void HandleTestRunStarted(object sender, TestRunStartedEventArgs e)
        {
            if (e.Tests is null)
            {
                return;
            }

            UpdateProgressBar(TestOutcome.Unknown, true);

            var running = Tests.Where(test => e.Tests.Contains(test.Method)).ToList();

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
        }

        private static readonly Dictionary<TestOutcome, Color> OutcomeColors = new Dictionary<TestOutcome, Color>
        {
            { TestOutcome.Unknown, Colors.DimGray },
            { TestOutcome.Succeeded, Colors.LimeGreen },
            { TestOutcome.Inconclusive, Colors.Gold },
            { TestOutcome.Ignored, Colors.Gold },
            { TestOutcome.Failed, Colors.Red }
        };

        private TestOutcome _bestOutcome = TestOutcome.Unknown;
        private void UpdateProgressBar(TestOutcome output, bool reset = false)
        {
            if (reset)
            {
                ExecutedCount = 0;
                CurrentRunTestCount = 0;
                ProgressBarColor = OutcomeColors[TestOutcome.Unknown];
                return;
            }

            ExecutedCount++;

            if (output <= _bestOutcome)
            {
                return;
            }

            _bestOutcome = output;
            ProgressBarColor = OutcomeColors[output];
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

        public void Dispose()
        {
            if (_testEngine != null)
            {
                _testEngine.TestRunStarted -= HandleTestRunStarted;
                _testEngine.TestCompleted -= HandleTestCompletion;
                _testEngine.TestStarted -= HandleTestStarted;
                _testEngine.TestsRefreshed -= HandleTestsRefreshed;
                _testEngine.TestRunCompleted -= HandleRunCompletion;
            }
        }
    }
}