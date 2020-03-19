using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Refactorings.Rename;
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
            nonExecutableMessage = Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove; // MoveMemberResources.NoDeclarationsSelectedToMove;
            return false;
        }

        public void RefactorRewrite(MoveMemberModel model, IRewriteSession session, IRewritingManager rewritingManager, bool asPreview = false)
        {
            if (string.IsNullOrEmpty(model.Destination.ModuleName))
            {
                return;
            }

            if (asPreview)
            {
                var contentToMove = new MovedContentProvider();
                contentToMove.AddFieldOrConstantDeclaration(NothingSelectedPreviewMessage);
                var isExistingDestination = model.Destination.IsExistingModule(out var module);
                if (isExistingDestination)
                {
                    var rewriter = session.CheckOutModuleRewriter(module.QualifiedModuleName);
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
        }

        public IMovedContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, IMovedContentProvider contentToMove)
        {
            contentToMove.AddFieldOrConstantDeclaration(NothingSelectedPreviewMessage);
            return contentToMove;
        }

        private string NothingSelectedPreviewMessage 
            => $"{Environment.NewLine}{Environment.NewLine}'****  {Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove /*MoveMemberResources.NoDeclarationsSelectedToMove*/}{Environment.NewLine}{Environment.NewLine}";
    }
}
