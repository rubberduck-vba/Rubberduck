using Rubberduck.Interaction.Navigation;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.UnitTesting.ViewModels
{
    public enum TestRunState
    {
        Stopped,
        Queued,
        Running
    }

    public class TestMethodViewModel : ViewModelBase, INavigateSource
    {
        public TestMethodViewModel(TestMethod test)
        {
            Method = test;
        }

        public TestMethod Method { get; }

        private TestResult _result = new TestResult(TestOutcome.Unknown);

        public TestResult Result
        {
            get => _result;
            set { _result = value; OnPropertyChanged(); }
        }

        private TestRunState _state;
        public TestRunState RunState
        {
            get => _state;
            set { _state = value; OnPropertyChanged(); }
        }

        public QualifiedMemberName QualifiedName => Method.Declaration.QualifiedName;

        // Delegate Navigability to encapsulated TestMethod
        public NavigateCodeEventArgs GetNavigationArgs() => Method.GetNavigationArgs();

        public override string ToString() => $"{Method.Declaration.QualifiedName}: {Result.Outcome} ({Result.Duration}ms) {Result.Output}";
        public override bool Equals(object obj) => obj is TestMethodViewModel other && Method.Equals(other.Method);
        public override int GetHashCode() => Method.GetHashCode();

        public object[] ToArray()
        {
            var declaration = Method.Declaration;
            return new object[] {
                declaration.QualifiedName.QualifiedModuleName.ProjectName,
                declaration.QualifiedName.QualifiedModuleName.ComponentName,
                declaration.IdentifierName, 
                _result.Outcome.ToString(),
                _result.Output,
                _result.Duration
            };
        }
    }
}
