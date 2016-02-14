using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using Rubberduck.Parsing.Reflection;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public abstract class TestExplorerModelBase : ViewModelBase
    {
        public abstract void Refresh();

        private readonly ObservableCollection<TestMethod> _tests = new ObservableCollection<TestMethod>();
        public ObservableCollection<TestMethod> Tests { get { return _tests; } }

        private static readonly string[] ReservedTestAttributeNames =
        {
            "ModuleInitialize",
            "TestInitialize", 
            "TestCleanup",
            "ModuleCleanup"
        };

        private readonly IList<TestMethod> _lastRun = new List<TestMethod>();
        public IEnumerable<TestMethod> LastRun { get { return _lastRun; } } 

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
        }

        private int _executedCount;
        public int ExecutedCount
        {
            get { return _executedCount; }
            protected set
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

        /// <summary>
        /// A method that determines whether a <see cref="Member"/> is a test method or not.
        /// </summary>
        /// <param name="member">The <see cref="Member"/> to evaluate.</param>
        /// <returns>Returns <c>true</c> if specified member is a test method.</returns>
        protected static bool IsTestMethod(Member member)
        {
            // todo: reimplement using declarations/annotations
            var isIgnoredMethod = member.HasAttribute<TestInitializeAttribute>()
                                  || member.HasAttribute<TestCleanupAttribute>()
                                  || member.HasAttribute<ModuleInitializeAttribute>()
                                  || member.HasAttribute<ModuleCleanupAttribute>()
                                  || (ReservedTestAttributeNames.Any(attribute =>
                                      member.QualifiedMemberName.MemberName.StartsWith(attribute)));

            var result = !isIgnoredMethod &&
                         (member.QualifiedMemberName.MemberName.StartsWith("Test") || member.HasAttribute<TestMethodAttribute>())
                         && member.Signature.Contains(member.QualifiedMemberName.MemberName + "()")
                         && member.MemberType == MemberType.Sub
                         && member.MemberVisibility == MemberVisibility.Public;

            return result;
        }
    }
}