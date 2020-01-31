using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    //public abstract class SingleMemberToStdModule : MoveMemberStrategyBase
    //{
    //    protected static bool IsApplicable(IMoveScenario scenario, IProvideMoveDeclarationGroups groups, DeclarationType declarationType)
    //    {
    //        if (!scenario.MoveDefinition.IsStdModuleDestination) { return false; }

    //        if (!MoveMemberStrategyCommon.IsSingleDeclarationSelection(groups, declarationType)) { return false; }

    //        if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneral(scenario, groups)) { return false; }

    //        if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneralMethod(scenario, groups)) { return false; }

    //        var unmoveableDeclarations = UnMoveableElements(scenario, groups);

    //        if (scenario.MoveDefinition.IsStdModuleSource)
    //        {
    //            return !unmoveableDeclarations.Any(ud => ud.HasPrivateAccessibility());
    //        }

    //        var externalMemberRefs = groups.SelectedElements.AllReferences().Where(rf => rf.QualifiedModuleName != scenario.QualifiedModuleNameSource);

    //        //External references to the moved elements not supported for Classes and Forms in this strategy
    //        if (externalMemberRefs.Any())
    //        {
    //            return false;
    //        }

    //        return !unmoveableDeclarations.Any();
    //    }

    //    private static IEnumerable<Declaration> UnMoveableElements(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
    //    {
    //        var theSelectedMember = groups.SelectedElements.Single();

    //        var exclusiveVariables = groups.SupportingElements.NonMembers
    //            .Where((se => se.References.All(rf => rf.ParentScoping.Equals(theSelectedMember))
    //                    || se.References.All(rf => groups.SupportingElements.Members.Contains(rf.ParentScoping))));

    //        var nonExclusiveVariables = groups.Participants.NonMembers.Except(exclusiveVariables);

    //        var allMembers = groups.SupportingElements.Members.Concat(groups.SelectedElements.Members);

    //        var exclusiveMembers = groups.SupportingElements.Members
    //            .Where(se => se.References.All(seRefs => allMembers.Contains(seRefs.ParentScoping)));

    //        var nonExclusiveMembers = groups.Participants.Members.Except(groups.SelectedElements.AllDeclarations).Except(exclusiveMembers);

    //        var unmoveableDeclarations = nonExclusiveMembers.Except(exclusiveMembers)
    //                    .Concat(nonExclusiveVariables.Except(exclusiveVariables));

    //        return unmoveableDeclarations;
    //    }
    //}

    public abstract class MoveMemberStrategyBase : IMoveMemberRefactoringStrategy
    {
        public abstract string DestinationNewModuleContent { get; } // => throw new NotImplementedException();

        public abstract int DestinationNewContentLineCount { get; } //=> throw new NotImplementedException();

        public abstract string DestinationMemberCodeBlock(Declaration member);
        //{
        //    throw new NotImplementedException();
        //}

        public abstract void ModifyContent();
        //{
        //    throw new NotImplementedException();
        //}

        public abstract string PreviewDestination();
        //{
        //    throw new NotImplementedException();
        //}
    }
}
