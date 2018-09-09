using Rubberduck.Interaction.Navigation;
using Rubberduck.UnitTesting;
using System.Collections.Generic;

namespace Rubberduck.UI.UnitTesting.ViewModels
{
    public class TestMethodViewModel : ViewModelBase, INavigateSource
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

        public override bool Equals(object obj)
        {
            var model = obj as TestMethodViewModel;
            return model != null &&
                   EqualityComparer<TestMethod>.Default.Equals(Method, model.Method);
        }

        public override int GetHashCode()
        {
            return 1003453392 + EqualityComparer<TestMethod>.Default.GetHashCode(Method);
        }
    }
}
