using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class SingleMemberNonPropertyToStdModule : MoveMemberStrategyBase
    {
        public static bool IsApplicable(IMoveScenario scenario, IProvideMoveDeclarationGroups groups) //, DeclarationType declarationType)
        {
            if (!scenario.MoveDefinition.IsStdModuleDestination) { return false; }

            if (!(MoveMemberStrategyCommon.IsSingleDeclarationSelection(groups, DeclarationType.Function)
                || MoveMemberStrategyCommon.IsSingleDeclarationSelection(groups, DeclarationType.Procedure)))
            {
                return false;
            }

            if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneral(scenario, groups)) { return false; }

            if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneralMethod(scenario, groups)) { return false; }

            var theSelectedMember = groups.SelectedElements.Single();

            var exclusiveVariables = groups.SupportingElements.NonMembers
                .Where((se => se.References.All(rf => rf.ParentScoping.Equals(theSelectedMember))
                        || se.References.All(rf => groups.SupportingElements.Members.Contains(rf.ParentScoping))));

            var nonExclusiveVariables = groups.Participants.NonMembers.Except(exclusiveVariables);

            var allMembers = groups.SupportingElements.Members.Concat(groups.SelectedElements.Members);

            var exclusiveMembers = groups.SupportingElements.Members
                .Where(se => se.References.All(seRefs => allMembers.Contains(seRefs.ParentScoping)));

            var nonExclusiveMembers = groups.Participants.Members.Except(groups.SelectedElements.AllDeclarations).Except(exclusiveMembers);

            var unmoveableDeclarations = nonExclusiveMembers.Except(exclusiveMembers)
                        .Concat(nonExclusiveVariables.Except(exclusiveVariables));

            if (scenario.MoveDefinition.IsStdModuleSource)
            {
                return !unmoveableDeclarations.Any(ud => ud.HasPrivateAccessibility());
            }

            var externalMemberRefs = groups.SelectedElements.AllReferences().Where(rf => rf.QualifiedModuleName != scenario.QualifiedModuleNameSource);

            //External references to the moved elements not supported for Classes and Forms in this strategy
            if (externalMemberRefs.Any())
            {
                return false;
            }

            return !unmoveableDeclarations.Any();
        }

        public static IMoveMemberRefactoringStrategy CreateStrategy(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
            => new SingleMemberNonPropertyToStdModule(scenario, rewritingManager);

        private readonly IMoveScenario _scenario;
        private readonly MoveMemberStrategyCommon _helper;

        private SingleMemberNonPropertyToStdModule(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
        {
            _scenario = scenario;
            _helper = new MoveMemberStrategyCommon(scenario, rewritingManager);
        }

        public override void ModifyContent() => _helper.ModifyContent(ModifySource);

        public override string PreviewDestination() => _helper.PreviewDestination();

        public override string DestinationMemberCodeBlock(Declaration member) => _helper.DestinationMemberCodeBlockDefault(member);

        public override string DestinationNewModuleContent => _helper.DestinationNewModuleContent;

        public override int DestinationNewContentLineCount => _helper.DestinationNewContentLineCount;

        private void ModifySource(IMoveEndpointRewriter sourceRewriter)
        {
            _helper.RemoveDeclarations(sourceRewriter);

            _helper.UpdateSourceReferencesToMovedElements(sourceRewriter);

            _helper.InsertNewSourceContent(sourceRewriter);
        }
    }
}
