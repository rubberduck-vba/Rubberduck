using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class NoIndentAnnotationCommand : ComCommandBase
    {
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IAnnotationUpdater _annotationUpdater;
        private readonly IRewritingManager _rewritingManager;

        public NoIndentAnnotationCommand(
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IRewritingManager rewritingManager,
            IAnnotationUpdater annotationUpdater,
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _selectedDeclarationProvider = selectedDeclarationProvider;
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

            return _selectedDeclarationProvider.SelectedModule();
        }
    }
}
