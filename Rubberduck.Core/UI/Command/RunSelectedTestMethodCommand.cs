using System.Linq;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class RunSelectedTestMethodCommand : CommandBase
    {
        private readonly ITestEngine _engine;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IDeclarationFinderProvider _finderProvider;

        public RunSelectedTestMethodCommand(ITestEngine engine, ISelectedDeclarationProvider selectedDeclarationProvider, IDeclarationFinderProvider finderProvider) 
        {
            _engine = engine;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _finderProvider = finderProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return (parameter ?? FindDeclarationFromSelection()) is Declaration candidate &&
                   !(_engine.Tests.FirstOrDefault(test => test.Declaration.Equals(candidate)) is null) &&
                   _engine.CanRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (!((parameter ?? FindDeclarationFromSelection()) is Declaration candidate) ||
                !(_engine.Tests.FirstOrDefault(test => test.Declaration.Equals(candidate)) is TestMethod selectedTest) ||
                !_engine.CanRun)
            {
                return;
            }

            _engine.Run(new [] { selectedTest });
        }

        private Declaration FindDeclarationFromSelection()
        {
            var selectedMember = _selectedDeclarationProvider.SelectedMember();
            return IsTestMethod(selectedMember)
                ? selectedMember
                : null;
        }

        private bool IsTestMethod(Declaration member)
        {
            return member != null 
                   && member.DeclarationType == DeclarationType.Procedure
                   && member.Annotations.Any(parseTreeAnnotation =>
                       parseTreeAnnotation.Annotation is TestMethodAnnotation);
        }
    }
}
