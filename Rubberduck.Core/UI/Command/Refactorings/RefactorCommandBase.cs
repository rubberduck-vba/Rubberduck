using NLog;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly IRewritingManager RewritingManager;
        protected readonly ISelectionService SelectionService;

        protected RefactorCommandBase(IRewritingManager rewritingManager, ISelectionService selectionService)
            : base (LogManager.GetCurrentClassLogger())
        {
            RewritingManager = rewritingManager;
            SelectionService = selectionService;
        }
    }
}