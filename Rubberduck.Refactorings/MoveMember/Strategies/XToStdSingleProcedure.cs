using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Refactorings.MoveMember
{
    public class XToStdSingleProcedure : IMoveMemberRefactoringStrategy
    {
        public static bool IsApplicable(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
        {
            if (!scenario.MoveDefinition.IsStdModuleDestination) { return false; }

            if (!MoveMemberStrategyCommon.IsSingleDeclarationSelection(groups, DeclarationType.Procedure)) { return false; }

            if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneral(scenario, groups)) { return false; }

            if (MoveMemberStrategyCommon.IsUnsupportedMoveGeneralMethod(scenario, groups)) { return false; }

            //Note: this checks if the object reference becomes self-referential after the move
            if (groups.SupportingElements.AllDeclarations.Any(mv => mv.IsObject
                && mv.IsVariable() && !mv.AsTypeName.Equals(scenario.MoveDefinition.Destination.Module.IdentifierName)))
            {
                //The strategy does not support moving objects with the exception of
                //variables that are existing instances of the destination module
                return false;
            }

            //var moveRequiresForwardingInternalState = false;

            var theSelectedMember = groups.SelectedElements.Single();

            var allMembers = groups.SupportingElements.Members.Concat(groups.SelectedElements.Members);

            var exclusiveMembers = groups.SupportingElements.Members
                .Where(m => m.References.All(seRefs => allMembers.Contains(seRefs.ParentScoping)));

            var nonExclusiveMembers = groups.SupportingElements.Members.Except(exclusiveMembers);

            var supportingFunctions = groups.SupportingElements.Members.Where(m => m.DeclarationType.HasFlag(DeclarationType.Function));
            foreach (var supportingFunction in supportingFunctions)
            {
                var memberWithParams = supportingFunction as IParameterizedDeclaration;
                if (memberWithParams.Parameters.Any() && nonExclusiveMembers.Contains(supportingFunction))
                {
                    return false;
                }
            }

            var callbackFunctions = Enumerable.Empty<Declaration>();
            var publicAccessibleProcedures = Enumerable.Empty<Declaration>();
            var functionsWithResultToPassAsArgumentValue = Enumerable.Empty<Declaration>();

            //Non-exclusive members can be part of the move if:
            //1. It is a publicly available member of a standard module
            //2. If it is a Function DeclarationType, then the return type must be a Value Type
            if (nonExclusiveMembers.Any())
            {
                functionsWithResultToPassAsArgumentValue = nonExclusiveMembers.Where(m => m.DeclarationType.HasFlag(DeclarationType.Function)
                    && ValueTypes.Contains(m.AsTypeName));

                callbackFunctions = nonExclusiveMembers.Where(m => m.DeclarationType.HasFlag(DeclarationType.Function)
                    && ValueTypes.Contains(m.AsTypeName)
                    && MoveMemberStrategyCommon.IsPubliclyAccessibleInStdModule(m)
                    && !functionsWithResultToPassAsArgumentValue.Contains(m));

                publicAccessibleProcedures = nonExclusiveMembers.Where(m => m.DeclarationType.Equals(DeclarationType.Procedure)
                    && MoveMemberStrategyCommon.IsPubliclyAccessibleInStdModule(m));

                if (callbackFunctions.Count() == 0 && publicAccessibleProcedures.Count() == 0 && functionsWithResultToPassAsArgumentValue.Count() == 0)
                {
                    //Unable to handle non-exclusive supporting members
                    return false;
                }
            }

            var supportConstants = groups.SupportingElements.AllDeclarations
                .Where(supportElement => supportElement.IsConstant());

            var exclusiveConstants = supportConstants
                .Where(constant => constant.References.All(rf => allMembers.Contains(rf.ParentScoping)));

            var constantsToPass = Enumerable.Empty<Declaration>();
            if (supportConstants.Count() != exclusiveConstants.Count()
                 || !supportConstants.All(smv => exclusiveConstants.Contains(smv)))
            {
                //moveRequiresForwardingInternalState = true;
                constantsToPass = supportConstants.Where(sv => sv.References.Any(rf => rf.ParentScoping.Equals(theSelectedMember))
                    && !sv.References.All(rf => rf.ParentScoping.Equals(theSelectedMember)));
            }

            var supportMemberVariables = groups.SupportingElements.AllDeclarations
                .Where(supportElement => supportElement.IsVariable());

            var exclusiveMemberVariables = supportMemberVariables
                .Where(memberVariable => memberVariable.References.All(rf => allMembers.Contains(rf.ParentScoping)));

            var nonExclusiveMemberVariables = supportMemberVariables
                        .Except(exclusiveMemberVariables);

            //if (nonExclusiveMemberVariables.Count() > 0 || 
            //    !(nonExclusiveMemberVariables.All(smv => !smv.HasPrivateAccessibility() && scenario.MoveDefinition.IsStdModuleSource)))
            if (!scenario.MoveDefinition.IsStdModuleSource) // nonExclusiveMemberVariables.Any(smv => smv.HasPrivateAccessibility()))
            {
                if (nonExclusiveMemberVariables.Count() > 0) { return false; }
            }
            else if(nonExclusiveMemberVariables.Any(smv => smv.HasPrivateAccessibility()))
            {
                return false;
            }

            //var retainedPublicVariablesOfStdSource = supportMemberVariables.Where(smv => scenario.MoveDefinition.IsStdModuleSource && !smv.HasPrivateAccessibility());

            var destinationInstanceVariables = Enumerable.Empty<Declaration>();
            var variablesToPass = Enumerable.Empty<Declaration>();
            var callbackAccessVariables = Enumerable.Empty<Declaration>();
            var accessibleNonExclusivePublicVariables = Enumerable.Empty<Declaration>();
            var nonExclusiveVariablesReferencedByFunctionsPassedAsResult = Enumerable.Empty<Declaration>();

            if (supportMemberVariables.Count() != exclusiveMemberVariables.Count()
                || !supportMemberVariables.All(smv => exclusiveMemberVariables.Contains(smv)))
            {
               // moveRequiresForwardingInternalState = true;
                accessibleNonExclusivePublicVariables = nonExclusiveMemberVariables
                    .Where(nemv => MoveMemberStrategyCommon.IsPubliclyAccessibleInStdModule(nemv));

                destinationInstanceVariables = supportMemberVariables
                     .Where(sv => sv.IsVariable() && sv.AsTypeName.Equals(scenario.MoveDefinition.Destination.Module.IdentifierName));

                variablesToPass = supportMemberVariables.Where(sv => sv.References.Any(rf => rf.ParentScoping.Equals(theSelectedMember))
                    && !sv.References.All(rf => rf.ParentScoping.Equals(theSelectedMember))
                    && !sv.AsTypeName.Equals(scenario.MoveDefinition.Destination.Module.IdentifierName));

                callbackAccessVariables = supportMemberVariables.Except(variablesToPass)
                    .Where(smv => smv.References.Any(rf => rf.ParentScoping == publicAccessibleProcedures)
                    && !smv.References.Any(rf => exclusiveMembers.Contains(rf.ParentScoping)));

                nonExclusiveVariablesReferencedByFunctionsPassedAsResult = nonExclusiveMemberVariables
                    .Where(nev => nev.References.Any(rf => functionsWithResultToPassAsArgumentValue.Contains(rf.ParentScoping)));

            }

            scenario.ForwardSelectedMemberCalls = false;// moveRequiresForwardingInternalState/* || !scenario.MoveDefinition.Endpoints.Equals(MoveEndpoints.StdToStd)*/;

            //MoveElementGroups MoveAndDelete { get; }
            //MoveElementGroups Retain { get; }
            //IEnumerable<Declaration> Remove { get; }
            //IEnumerable<Declaration> Forward { get; }
            var unaccountedElements = groups.SupportingElements.AllDeclarations
                .Except(exclusiveMemberVariables) //MoveAndDelete
                .Except(variablesToPass) //Retain..and Add argument and Add parameter to moved member signature
                .Except(exclusiveConstants) //MoveAndDelete
                .Except(constantsToPass) ////Retain...and Add argument and Add parameter to moved member signature
                .Except(exclusiveMembers) //MoveAndDelete
                .Except(callbackFunctions) //Retain
                .Except(functionsWithResultToPassAsArgumentValue) //Retain...and Add argument and Add parameter to moved member signature
                .Except(nonExclusiveVariablesReferencedByFunctionsPassedAsResult) //Retain
                .Except(callbackAccessVariables) //Retain
                .Except(accessibleNonExclusivePublicVariables)//Retain
                .Except(destinationInstanceVariables) //Retain
                .Except(publicAccessibleProcedures); //Retain
                                                     //.Except(groups.MoveAndDelete.AllDeclarations);

            var localMoveAndDelete = exclusiveMemberVariables
                .Concat(exclusiveConstants)
                .Concat(exclusiveMembers);

            //Debug.Assert(localMoveAndDelete.Count() == groups.AllDeclarations.Count());

            return !unaccountedElements.Any();

        }

        private static string[] ValueTypes = new string[]
        {
            Tokens.Boolean,
            Tokens.Currency,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.Single,
            Tokens.Double,
            Tokens.String,
            Tokens.Byte
        };

        public static IMoveMemberRefactoringStrategy CreateStrategy(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
            => new XToStdSingleProcedure(scenario, rewritingManager);

        private readonly IMoveScenario _scenario;
        private readonly MoveMemberStrategyCommon _helper;

        private XToStdSingleProcedure(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
        {
            _scenario = scenario;
            _helper = new MoveMemberStrategyCommon(scenario, rewritingManager);
        }

        public void ModifyContent() => _helper.ModifyContent(ModifySource);

        public string PreviewDestination() => _helper.PreviewDestination();

        public string DestinationMemberCodeBlock(Declaration member) => _helper.DestinationMemberCodeBlockDefault(member);

        public string DestinationNewModuleContent => _helper.DestinationNewModuleContent;

        public int DestinationNewContentLineCount => _helper.DestinationNewContentLineCount;

        private void ModifySource(IMoveEndpointRewriter sourceRewriter)
        {
            _helper.RemoveDeclarations(sourceRewriter);

            if (_scenario.ForwardSelectedMemberCalls)
            {
                var procedure = _scenario.SelectedDeclarations.Where(d => d.IsMember()).Single();
                var _PROCEDURE_BODY_FORMAT = "    {0}.{1} {2}";
                var newBody = string.Format(_PROCEDURE_BODY_FORMAT, MemberAccessLExpr, procedure.IdentifierName, _helper.CallSiteArguments(procedure));
                sourceRewriter.ReplaceDescendentContext<VBAParser.BlockContext>(procedure, $"{newBody}{Environment.NewLine}");
            }

            _helper.ReplaceMovedOrRenamedReferenceIdentifiers(sourceRewriter);

            _helper.UpdateSourceReferencesToMovedElements(sourceRewriter);

            _helper.EnsureClassIsValidWhereReferenced(sourceRewriter);

            _helper.InsertNewSourceContent(sourceRewriter);
        }

        private string MemberAccessLExpr
                => _scenario.MoveDefinition.IsClassModuleDestination ?
                    _helper.ForwardToModuleLExpression()
                        : _scenario.MoveDefinition.Destination.ModuleName;
    }
}
