using Rubberduck.Common;
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
    public class SingleConstantToStdModule : IMoveMemberRefactoringStrategy
    {
        public static bool IsApplicable(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
        {
            if (!scenario.MoveDefinition.IsStdModuleDestination) { return false; }

            if (!MoveMemberStrategyCommon.IsSingleDeclarationSelection(groups, DeclarationType.Constant)) { return false; }

            if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneral(scenario, groups)) { return false; }

            return true;
        }

        public static IMoveMemberRefactoringStrategy CreateStrategy(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
            => new SingleConstantToStdModule(scenario, rewritingManager);

        private readonly IMoveScenario _scenario;
        private readonly MoveMemberStrategyCommon _helper;

        private SingleConstantToStdModule(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
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

            var constStmt = nonMember.Context.GetAncestor<VBAParser.ConstStmtContext>();
            Debug.Assert(constStmt != null);

            var visibility = _scenario.IsOnlyReferencedByMovedElements(nonMember) ? Tokens.Private : Tokens.Public;
            var declarationBlock = $"{visibility} {Tokens.Const} {tempRewriter.GetModifiedText(nonMember)}"; //.Context.Start.TokenIndex, nonMember.Context.Stop.TokenIndex)}";
            _scenario.DestinationContentProvider.AddDeclarationBlock(declarationBlock);
        }
    }
}
