using Rubberduck.Interaction.Navigation;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting.ViewModels
{
    internal class TestMethodViewModel : ViewModelBase, INavigateSource
    {
        public TestMethod Method { get; private set; }

        // Delegate Navigability to encapsulated TestMethod
        public NavigateCodeEventArgs GetNavigationArgs() => Method.GetNavigationArgs();

        private TestResult _result = new TestResult(TestOutcome.Unknown);

        public TestMethodViewModel(TestMethod test)
        {
            Method = test;
        }

        public TestResult Result
        {
            get => _result;
            set { _result = value; OnPropertyChanged(); }
        }

        public override string ToString()
        {
            return $"{Method.Declaration.QualifiedName}: {Result.Outcome} ({Result.Duration}ms) {Result.Output}";
        }

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
