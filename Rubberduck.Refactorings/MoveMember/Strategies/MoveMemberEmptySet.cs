using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.MoveMember.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    /// <summary>
    /// MoveMemberEmptySet supports the scenario where
    /// the user has unselected all moveable declarations
    /// </summary>
    public class MoveMemberEmptySet : IMoveMemberRefactoringStrategy
    {
        public bool IsApplicable(MoveMemberModel model)
        {
            if (!model.SelectedDeclarations.Any())
            {
                return true;
            }

            return false;
        }

        public bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage)
        {
            nonExecutableMessage = MoveMemberResources.NoDeclarationsSelectedToMove;
            return false;
        }

        public IMoveMemberRewriteSession RefactorRewrite(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove, bool asPreview = false)
        {
            if (string.IsNullOrEmpty(model.Destination.ModuleName))
            {
                return model.MoveMemberFactory.CreateMoveMemberRewriteSession(rewritingManager.CheckOutCodePaneSession());
            }

            var moveSession = model.MoveMemberFactory.CreateMoveMemberRewriteSession(rewritingManager.CheckOutCodePaneSession());
            if (asPreview)
            {

                contentToMove.AddFieldOrConstantDeclaration(NothingSelectedPreviewMessage);
                var isExistingDestination = model.Destination.IsExistingModule(out var module);
                if (isExistingDestination)
                {
                    var rewriter = moveSession.CheckOutModuleRewriter(module.QualifiedModuleName);
                    if (model.Destination.TryGetCodeSectionStartIndex(out var insertIndex))
                    {
                        rewriter.InsertBefore(insertIndex, contentToMove.AsSingleBlockWithinDemarcationComments());
                    }
                    else
                    {
                        rewriter.InsertAtEndOfFile(contentToMove.AsSingleBlockWithinDemarcationComments());
                    }
                }
            }
            return moveSession;
        }

        public INewContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove)
        {
            contentToMove.AddFieldOrConstantDeclaration(NothingSelectedPreviewMessage);
            return contentToMove;
        }

        private string NothingSelectedPreviewMessage 
            => $"{Environment.NewLine}{Environment.NewLine}'****  {MoveMemberResources.NoDeclarationsSelectedToMove}{Environment.NewLine}{Environment.NewLine}";
    }
}
