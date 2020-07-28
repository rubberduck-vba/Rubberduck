using System.Linq;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class RunSelectedTestModuleCommand : CommandBase
    {
        private readonly ITestEngine _engine;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public RunSelectedTestModuleCommand(ITestEngine engine, ISelectedDeclarationProvider selectedDeclarationProvider)
        {
            _engine = engine;
            _selectedDeclarationProvider = selectedDeclarationProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return (parameter ?? FindModuleFromSelection()) is Declaration candidate &&
                   candidate.DeclarationType == DeclarationType.ProceduralModule &&
                   candidate.Annotations.Any(annotation => annotation is TestModuleAnnotation) &&
                   _engine.CanRun &&
                   _engine.Tests.Any(test => test.Declaration.QualifiedModuleName.Equals(candidate.QualifiedModuleName));
        }

        protected override void OnExecute(object parameter)
        {
            if (!((parameter ?? FindModuleFromSelection()) is Declaration candidate) ||
                candidate.DeclarationType != DeclarationType.ProceduralModule ||
                !candidate.Annotations.Any(annotation => annotation is TestModuleAnnotation) ||
                !_engine.CanRun)
            {
                return;
            }

            var tests = _engine.Tests.Where(test => test.Declaration.QualifiedModuleName.Equals(candidate.QualifiedModuleName)).ToList();

            if (!tests.Any())
            {
                return;
            }

            _engine.Run(tests);
        }

        private Declaration FindModuleFromSelection()
        {
            return _selectedDeclarationProvider.SelectedModule();
        }
    }
}
