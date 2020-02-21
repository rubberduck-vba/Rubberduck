using Rubberduck.Parsing.Rewriter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberToUndefined : IMoveMemberRefactoringStrategy
    {
        public bool IsApplicable(MoveMemberModel model)
        {
            if (!model.SelectedDeclarations.Any()) { return false; }

            if (model.Destination.IsExistingModule(out _)) { return false; }

            return string.IsNullOrEmpty(model.Destination.ModuleName);
        }

        public bool IsAnExecutableStrategy => false;


        public IMoveMemberRewriteSession RefactorRewrite(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove, bool asPreview = false)
        {
            return model.MoveMemberFactory.CreateMoveMemberRewriteSession(rewritingManager.CheckOutCodePaneSession());
        }

        public INewContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove)
        {
            contentToMove.AddFieldOrConstantDeclaration("'Undefined Destination Module");
            return contentToMove;
        }
    }

    public class MoveMemberEmptySet : IMoveMemberRefactoringStrategy
    {
        public bool IsApplicable(MoveMemberModel model)
        {
            if (!model.SelectedDeclarations.Any()) { return true; }

            return false;
        }

        public bool IsAnExecutableStrategy => false;

        public IMoveMemberRewriteSession RefactorRewrite(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove, bool asPreview = false)
        {
            return model.MoveMemberFactory.CreateMoveMemberRewriteSession(rewritingManager.CheckOutCodePaneSession());
        }

        public INewContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove)
        {
            contentToMove.AddFieldOrConstantDeclaration("'No Declarations Selected to Move");
            return contentToMove;
        }
    }
}
