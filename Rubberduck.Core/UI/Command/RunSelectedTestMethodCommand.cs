using System.Linq;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command
{
    public class RunSelectedTestMethodCommand : CommandBase
    {
        private readonly ITestEngine _engine;
        private readonly ISelectionService _selectionService;
        private readonly IDeclarationFinderProvider _finderProvider;

        public RunSelectedTestMethodCommand(ITestEngine engine, ISelectionService selectionService, IDeclarationFinderProvider finderProvider) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _engine = engine;
            _selectionService = selectionService;
            _finderProvider = finderProvider;
        }

        protected override bool EvaluateCanExecute(object parameter)
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
            var active = _selectionService?.ActiveSelection();
            if (!active.HasValue)
            {
                return null;
            }

            return _finderProvider.DeclarationFinder.FindDeclarationsForSelection(active.Value)
                .SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Procedure &&
                                                declaration.Annotations.Any(annotation => annotation is TestMethodAnnotation));
        }
    }
}
