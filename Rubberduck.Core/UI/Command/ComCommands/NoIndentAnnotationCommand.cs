using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class NoIndentAnnotationCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectionService _selectionService;
        private readonly IAnnotationUpdater _annotationUpdater;
        private readonly IRewritingManager _rewritingManager;

        public NoIndentAnnotationCommand(
            IVBE vbe, 
            IDeclarationFinderProvider declarationFinderProvider, 
            ISelectionService selectionService,
            IRewritingManager rewritingManager,
            IAnnotationUpdater annotationUpdater,
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _vbe = vbe;
            _declarationFinderProvider = declarationFinderProvider;
            _selectionService = selectionService;
            _rewritingManager = rewritingManager;
            _annotationUpdater = annotationUpdater;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            var target = FindTarget(parameter);
            return target != null
                   && target.DeclarationType.HasFlag(DeclarationType.Module)
                   && !target.Annotations.Any(a => a.Annotation is NoIndentAnnotation);
        }

        protected override void OnExecute(object parameter)
        {
            var target = FindTarget(parameter);
            if (target == null)
            {
                return;
            }

            var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
            _annotationUpdater.AddAnnotation(rewriteSession, target, new NoIndentAnnotation());
            rewriteSession.TryRewrite();
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return null;
            }

            return _declarationFinderProvider.DeclarationFinder?
                .UserDeclarations(DeclarationType.Module)
                .FirstOrDefault(module => module.QualifiedModuleName.Equals(activeSelection.Value.QualifiedName));
        }
    }
}
