using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringStrategy
    {
        IMoveMemberRewriteSession RefactorRewrite(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider movedContent, bool asPreview = false);
        INewContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove);
        bool IsApplicable(MoveMemberModel model);
        bool IsAnExecutableStrategy { get; }
    }

    public class MoveMemberStrategyBase
    {
    }
}
