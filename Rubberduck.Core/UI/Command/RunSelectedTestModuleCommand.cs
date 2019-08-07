using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command
{
    public class RunSelectedTestModuleCommand : CommandBase
    {
        private readonly ITestEngine _engine;
        private readonly ISelectionService _selectionService;
        private readonly IDeclarationFinderProvider _finderProvider;

        public RunSelectedTestModuleCommand(ITestEngine engine, ISelectionService selectionService, IDeclarationFinderProvider finderProvider)
        {
            _engine = engine;
            _selectionService = selectionService;
            _finderProvider = finderProvider;

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
            var active = _selectionService?.ActiveSelection();
            return !active.HasValue
                ? null
                : _finderProvider.DeclarationFinder.ModuleDeclaration(active.Value.QualifiedName);
        }
    }
}
