using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ParameterCanBeByValInspection : InspectionBase
    {
        public ParameterCanBeByValInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var parameters = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Parameter)
                .OfType<ParameterDeclaration>().ToList();
            var parametersThatCanBeChangedToBePassedByVal = new List<ParameterDeclaration>();

            var interfaceDeclarationMembers = State.DeclarationFinder.FindAllInterfaceMembers().ToList();
            var interfaceScopeDeclarations = State.DeclarationFinder
                .FindAllInterfaceImplementingMembers()
                .Concat(interfaceDeclarationMembers)
                .ToHashSet();

            parametersThatCanBeChangedToBePassedByVal.AddRange(InterFaceMembersThatCanBeChangedToBePassedByVal(interfaceDeclarationMembers));

            var eventMembers = State.DeclarationFinder.UserDeclarations(DeclarationType.Event).ToList();
            var formEventHandlerScopeDeclarations = State.FindFormEventHandlers();
            var eventHandlerScopeDeclarations = State.DeclarationFinder.FindEventHandlers().Concat(parameters.FindUserEventHandlers());
            var eventScopeDeclarations = eventMembers
                .Concat(formEventHandlerScopeDeclarations)
                .Concat(eventHandlerScopeDeclarations)
                .ToHashSet();

            parametersThatCanBeChangedToBePassedByVal.AddRange(EventMembersThatCanBeChangedToBePassedByVal(eventMembers));

            parametersThatCanBeChangedToBePassedByVal
                .AddRange(parameters.Where(parameter => CanBeChangedToBePassedByVal(parameter, eventScopeDeclarations, interfaceScopeDeclarations)));

            return parametersThatCanBeChangedToBePassedByVal
                .Where(parameter => !IsIgnoringInspectionResultFor(parameter, AnnotationName))
                .Select(parameter => new DeclarationInspectionResult(this, string.Format(InspectionResults.ParameterCanBeByValInspection, parameter.IdentifierName), parameter));
        }

        private bool CanBeChangedToBePassedByVal(ParameterDeclaration parameter, HashSet<Declaration> eventScopeDeclarations, HashSet<Declaration> interfaceScopeDeclarations)
        {
            var enclosingMember = parameter.ParentScopeDeclaration;
            var isIssue = !interfaceScopeDeclarations.Contains(enclosingMember)
                          && !eventScopeDeclarations.Contains(enclosingMember)
                          && CanBeChangedToBePassedByValIndividually(parameter);
            return isIssue;
        }

        private bool CanBeChangedToBePassedByValIndividually(ParameterDeclaration parameter)
        {
            var canPossiblyBeChangedToBePassedByVal =
                !parameter.IsArray
                && !parameter.IsParamArray
                && (parameter.IsByRef || parameter.IsImplicitByRef)
                && !IsParameterOfDeclaredLibraryFunction(parameter)
                && (parameter.AsTypeDeclaration == null 
                    || (parameter.AsTypeDeclaration.DeclarationType != DeclarationType.ClassModule 
                        && parameter.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType 
                        && parameter.AsTypeDeclaration.DeclarationType != DeclarationType.Enumeration))
                && !parameter.References.Any(reference => reference.IsAssignment)
                && !IsPotentiallyUsedAsByRefParameter(parameter);
            return canPossiblyBeChangedToBePassedByVal;
        }

        private static bool IsParameterOfDeclaredLibraryFunction(ParameterDeclaration parameter)
        {
            var parentMember = parameter.ParentScopeDeclaration;
            return parentMember.DeclarationType == DeclarationType.LibraryFunction
                   || parentMember.DeclarationType == DeclarationType.LibraryProcedure;
        }

        private IEnumerable<ParameterDeclaration> InterFaceMembersThatCanBeChangedToBePassedByVal(List<Declaration> interfaceMembers)
        {
            foreach (var memberDeclaration in interfaceMembers.OfType<ModuleBodyElementDeclaration>())
            {
                var interfaceParameters = memberDeclaration.Parameters.ToList();
                if (!interfaceParameters.Any())
                {
                    continue;
                }

                var parameterCanBeChangedToBeByVal = interfaceParameters.Select(parameter => CanBeChangedToBePassedByValIndividually(parameter)).ToList();

                var implementingMembers = State.DeclarationFinder.FindInterfaceImplementationMembers(memberDeclaration);
                foreach (var implementingMember in implementingMembers)
                {
                    var implementationParameters = implementingMember.Parameters.ToList();

                    //If you hit this assert, reopen https://github.com/rubberduck-vba/Rubberduck/issues/3906
                    Debug.Assert(parameterCanBeChangedToBeByVal.Count == implementationParameters.Count);

                    for (var i = 0; i < implementationParameters.Count; i++)
                    {
                        parameterCanBeChangedToBeByVal[i] = parameterCanBeChangedToBeByVal[i]
                                                            && CanBeChangedToBePassedByValIndividually(implementationParameters[i]);
                    }
                }

                for (var i = 0; i < parameterCanBeChangedToBeByVal.Count; i++)
                {
                    if (parameterCanBeChangedToBeByVal[i])
                    {
                        yield return interfaceParameters[i];
                    }
                }
            }
        }

        private IEnumerable<ParameterDeclaration> EventMembersThatCanBeChangedToBePassedByVal(IEnumerable<Declaration> eventMembers)
        {
            foreach (var memberDeclaration in eventMembers)
            {
                var eventParameters = (memberDeclaration as IParameterizedDeclaration)?.Parameters.ToList();
                if (!eventParameters?.Any() ?? false)
                {
                    continue;
                }

                var parameterCanBeChangedToBeByVal = eventParameters.Select(parameter => parameter.IsByRef).ToList();

                //todo: Find a better way to find the handlers.
                var eventHandlers = State.DeclarationFinder
                    .AllUserDeclarations
                    .FindHandlersForEvent(memberDeclaration)
                    .Select(s => s.Item2)
                    .ToList();

                foreach (var eventHandler in eventHandlers.OfType<IParameterizedDeclaration>())
                {
                    var handlerParameters = eventHandler.Parameters.ToList();

                    //If you hit this assert, reopen https://github.com/rubberduck-vba/Rubberduck/issues/3906
                    Debug.Assert(parameterCanBeChangedToBeByVal.Count == handlerParameters.Count);

                    for (var i = 0; i < handlerParameters.Count; i++)
                    {
                        parameterCanBeChangedToBeByVal[i] = parameterCanBeChangedToBeByVal[i] 
                                                            && CanBeChangedToBePassedByValIndividually(handlerParameters[i]);
                    }
                }

                for (var i = 0; i < parameterCanBeChangedToBeByVal.Count; i++)
                {
                    if (parameterCanBeChangedToBeByVal[i])
                    {
                        yield return eventParameters[i];
                    }
                }
            }
        }

        private bool IsPotentiallyUsedAsByRefParameter(ParameterDeclaration parameter)
        {
            return IsPotentiallyUsedAsByRefMethodParameter(parameter) 
                   || IsPotentiallyUsedAsByRefEventParameter(parameter);
        }

        private bool IsPotentiallyUsedAsByRefMethodParameter(ParameterDeclaration parameter)
        {
            //The condition on the text of the argument context excludes the cases where the argument is either passed explicitly by value 
            //or used inside a non-trivial expression, e.g. an arithmetic expression.
            var argumentsBeingTheParameter = parameter.References
                .Select(reference => reference.Context.GetAncestor<VBAParser.ArgumentExpressionContext>())
                .Where(context => context != null && context.GetText().Equals(parameter.IdentifierName, StringComparison.OrdinalIgnoreCase));

            foreach (var argument in argumentsBeingTheParameter)
            {
                var parameterCorrespondingToArgument = State.DeclarationFinder
                    .FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argument, parameter.QualifiedModuleName);

                if (parameterCorrespondingToArgument == null)
                {
                    //We have no idea what parameter it is passed to ar argument. So, we have to err on the safe side and assume it is passed by reference.
                    return true;
                }

                if (parameterCorrespondingToArgument.IsByRef)
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsPotentiallyUsedAsByRefEventParameter(ParameterDeclaration parameter)
        {
            //The condition on the text of the eventArgument context excludes the cases where the argument is either passed explicitly by value 
            //or used inside a non-trivial expression, e.g. an arithmetic expression.
            var argumentsBeingTheParameter = parameter.References
                .Select(reference => reference.Context.GetAncestor<VBAParser.EventArgumentContext>())
                .Where(context => context != null && context.GetText().Equals(parameter.IdentifierName, StringComparison.OrdinalIgnoreCase));

            foreach (var argument in argumentsBeingTheParameter)
            {
                var parameterCorrespondingToArgument = State.DeclarationFinder
                    .FindParameterFromSimpleEventArgumentNotPassedByValExplicitly(argument, parameter.QualifiedModuleName);

                if (parameterCorrespondingToArgument == null)
                {
                    //We have no idea what parameter it is passed to ar argument. So, we have to err on the safe side and assume it is passed by reference.
                    return true;
                }

                if (parameterCorrespondingToArgument.IsByRef)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
