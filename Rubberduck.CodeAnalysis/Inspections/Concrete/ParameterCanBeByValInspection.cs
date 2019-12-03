using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags parameters that are passed by reference (ByRef), but could be passed by value (ByVal).
    /// </summary>
    /// <why>
    /// Explicitly specifying a ByVal modifier on a parameter makes the intent explicit: this parameter is not meant to be assigned. In contrast, 
    /// a parameter that is passed by reference (implicitly, or explicitly ByRef) makes it ambiguous from the calling code's standpoint, whether the 
    /// procedure might re-assign these ByRef values and introduce a bug.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long, bar As Long)
    ///     Debug.Print foo, bar
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething(ByVal foo As long, ByRef bar As Long)
    ///     bar = foo * 2 ' ByRef parameter assignment: passing it ByVal could introduce a bug.
    ///     Debug.Print foo, bar
    /// End Sub
    /// ]]>
    /// </example>
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
            var formEventHandlerScopeDeclarations = State.DeclarationFinder.FindFormEventHandlers();
            var eventHandlerScopeDeclarations = State.DeclarationFinder.FindEventHandlers();
            var eventScopeDeclarations = eventMembers
                .Concat(formEventHandlerScopeDeclarations)
                .Concat(eventHandlerScopeDeclarations)
                .ToHashSet();

            parametersThatCanBeChangedToBePassedByVal.AddRange(EventMembersThatCanBeChangedToBePassedByVal(eventMembers));

            parametersThatCanBeChangedToBePassedByVal
                .AddRange(parameters.Where(parameter => CanBeChangedToBePassedByVal(parameter, eventScopeDeclarations, interfaceScopeDeclarations)));

            return parametersThatCanBeChangedToBePassedByVal
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
                    || (!parameter.AsTypeDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule)
                        && parameter.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType))
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

                var eventHandlers = State.DeclarationFinder
                    .FindEventHandlers(memberDeclaration)
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
