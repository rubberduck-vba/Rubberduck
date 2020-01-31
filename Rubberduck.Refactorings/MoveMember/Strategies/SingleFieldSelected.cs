using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class SingleFieldToStdModule : IMoveMemberRefactoringStrategy
    {
        public static bool IsApplicable(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
        {
            if (!scenario.MoveDefinition.IsStdModuleDestination) { return false; }

            if (!MoveMemberStrategyCommon.IsSingleDeclarationSelection(groups, DeclarationType.Variable)) { return false; }

            if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneral(scenario, groups)) { return false; }

            //We do not move fields of Private UserDefinedTypes
            if (scenario.MoveDefinition.SelectedElements.Any(nm => (nm.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
                && nm.AsTypeDeclaration.HasPrivateAccessibility())) { return false; }

            //We do not move fields of Private Enums
            if (scenario.MoveDefinition.SelectedElements.Any(nm => (nm.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.Enumeration) ?? false)
                && nm.AsTypeDeclaration.HasPrivateAccessibility())) { return false; }

            return true;
        }

        public static IMoveMemberRefactoringStrategy CreateStrategy(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
            => new SingleFieldToStdModule(scenario, rewritingManager);

        private readonly IMoveScenario _scenario;
        private readonly MoveMemberStrategyCommon _helper;

        private SingleFieldToStdModule(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
        {
            _scenario = scenario;
            _helper = new MoveMemberStrategyCommon(scenario, rewritingManager)
            {
                PrepareNewDestinationCodeElementsAction = PrepareNewDestinationCodeElements
            };
        }

        public void ModifyContent() => _helper.ModifyContent(ModifySource);

        public string PreviewDestination() => _helper.PreviewDestination();

        public string DestinationMemberCodeBlock(Declaration member) => _helper.DestinationMemberCodeBlockDefault(member);

        public string DestinationNewModuleContent => _helper.DestinationNewModuleContent;

        public int DestinationNewContentLineCount => _helper.DestinationNewContentLineCount;

        private string MemberAccessLExpr => _helper.ForwardToModuleLExpression();

        private void ModifySource(IMoveEndpointRewriter sourceRewriter)
        {
            _helper.RemoveDeclarations(sourceRewriter);

            _helper.ReplaceMovedOrRenamedReferenceIdentifiers(sourceRewriter);

            _helper.UpdateSourceReferencesToMovedElements(sourceRewriter);

            _helper.InsertNewSourceContent(sourceRewriter);
        }

        private void PrepareNewDestinationCodeElements(IMoveEndpointRewriter tempRewriter)
        {

            var groups = _scenario as IProvideMoveDeclarationGroups;
            var nonMember = groups.SelectedElements.NonMembers.First();

            var variableStmt = nonMember.Context.GetAncestor<VBAParser.VariableStmtContext>();
            Debug.Assert(variableStmt != null);

            var visibility = _scenario.IsOnlyReferencedByMovedElements(nonMember) ? Tokens.Private : Tokens.Public;
            var declarationBlock = $"{visibility} {tempRewriter.GetModifiedText(nonMember)}";
            _scenario.DestinationContentProvider.AddDeclarationBlock(declarationBlock);
        }
    }
}
