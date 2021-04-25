using System.Diagnostics;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags parameters that are passed by reference (ByRef), but could be passed by value (ByVal).
    /// </summary>
    /// <why>
    /// Explicitly specifying a ByVal modifier on a parameter makes the intent explicit: this parameter is not meant to be assigned. In contrast, 
    /// a parameter that is passed by reference (implicitly, or explicitly ByRef) makes it ambiguous from the calling code's standpoint, whether the 
    /// procedure might re-assign these ByRef values and introduce a bug.
    /// </why>
    /// <remarks>For performance reasons, this inspection will not flag a parameter that is passed as an argument to a procedure that also accepts it ByRef.</remarks>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long, bar As Long)
    ///     Debug.Print foo, bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething(ByVal foo As long, ByRef bar As Long)
    ///     bar = foo * 2 ' ByRef parameter assignment: passing it ByVal could introduce a bug.
    ///     Debug.Print foo, bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething(ByVal foo As long, ByRef bar As Long)
    ///     DoSomethingElse bar ' ByRef argument will not be flagged
    ///     Debug.Print foo, bar
    /// End Sub
    ///
    /// Private Sub DoSomethingElse(ByRef wouldNeedRecursiveLogic As Long)
    ///    Debug.Print wouldNeedRecursiveLogic
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ParameterCanBeByValInspection : DeclarationInspectionBase
    {
        public ParameterCanBeByValInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Parameter)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (!(declaration is ParameterDeclaration parameter)
                || !CanBeChangedToBePassedByValIndividually(parameter, finder))
            {
                return false;
            }

            var enclosingMember = declaration.ParentDeclaration;
            if (IsLibraryMethod(enclosingMember))
            {
                return false;
            }

            if (enclosingMember is EventDeclaration eventDeclaration)
            {
                return AllHandlerParametersCanBeChangedToByVal(parameter, eventDeclaration, finder);
            }

            if (enclosingMember is ModuleBodyElementDeclaration member)
            {
                if(member.IsInterfaceMember)
                {
                    return AllImplementationParametersCanBeChangedToByVal(parameter, member, finder);
                }

                if (member.IsInterfaceImplementation
                    || IsEventHandler(member, finder))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsLibraryMethod(Declaration declaration)
        {
            return declaration.DeclarationType == DeclarationType.LibraryProcedure
                   || declaration.DeclarationType == DeclarationType.LibraryFunction;
        }

        private static bool IsEventHandler(Declaration declaration, DeclarationFinder finder)
        {
            return finder.FindEventHandlers().Contains(declaration)
                   || finder.FindFormEventHandlers().Contains(declaration);
        }

        private static bool CanBeChangedToBePassedByValIndividually(ParameterDeclaration parameter, DeclarationFinder finder)
        {
            return !parameter.IsArray
                   && !parameter.IsParamArray
                   && parameter.IsByRef
                   && !parameter.References
                       .Any(reference => reference.IsAssignment)
                   && (parameter.AsTypeDeclaration == null
                       || parameter.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType)
                   && !IsPotentiallyUsedAsByRefParameter(parameter, finder);
        }

        private static bool IsPotentiallyUsedAsByRefParameter(ParameterDeclaration parameter, DeclarationFinder finder)
        {
            return IsPotentiallyUsedAsByRefMethodParameter(parameter, finder)
                   || IsPotentiallyUsedAsByRefEventParameter(parameter, finder);
        }

        private static bool IsPotentiallyUsedAsByRefMethodParameter(ParameterDeclaration parameter, DeclarationFinder finder)
        {
            var module = parameter.QualifiedModuleName;
            return parameter.References.Any(reference => IsPotentiallyAssignedByRefArgument(module, reference, finder));
        }

        private static bool IsPotentiallyAssignedByRefArgument(QualifiedModuleName module, IdentifierReference reference, DeclarationFinder finder)
        {
            var argExpression = ImmediateArgumentExpressionContext(reference);

            if (argExpression == null)
            {
                return false;
            }

            var argument = argExpression.GetAncestor<VBAParser.ArgumentContext>();
            var parameter = finder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argument, module);

            if (parameter == null)
            {
                //We have no idea what parameter it is passed to as argument. So, we have to err on the safe side and assume it is passed by reference.
                return true;
            }

            //We do not check whether the argument the parameter is actually assigned to costly recursions.
            return parameter.IsByRef;
        }

        private static VBAParser.ArgumentExpressionContext ImmediateArgumentExpressionContext(IdentifierReference reference)
        {
            var context = reference.Context;
            //The context is either already a simpleNameExprContext or an IdentifierValueContext used in a sub-rule of some other lExpression alternative. 
            var lExpressionNameContext = context is VBAParser.SimpleNameExprContext simpleName
                ? simpleName
                : context.GetAncestor<VBAParser.LExpressionContext>();

            //To be an immediate argument and, thus, assignable by ref, the structure must be argumentExpression -> expression -> lExpression.
            return lExpressionNameContext?
                .Parent?
                .Parent as VBAParser.ArgumentExpressionContext;
        }

        private static bool IsPotentiallyUsedAsByRefEventParameter(ParameterDeclaration parameter, DeclarationFinder finder)
        {
            var module = parameter.QualifiedModuleName;
            return parameter.References.Any(reference => IsPotentiallyAssignedByRefEventArgument(module, reference, finder));
        }

        private static bool IsPotentiallyAssignedByRefEventArgument(QualifiedModuleName module, IdentifierReference reference, DeclarationFinder finder)
        {
            var eventArgExpression = ImmediateEventArgumentExpressionContext(reference);

            if (eventArgExpression == null
                || eventArgExpression.BYVAL() != null)
            {
                return false;
            }

            var parameter = finder.FindParameterFromSimpleEventArgumentNotPassedByValExplicitly(eventArgExpression, module);

            if (parameter == null)
            {
                //We have no idea what parameter it is passed to as argument. So, we have to err on the safe side and assume it is passed by reference.
                return true;
            }

            return parameter.IsByRef;
        }

        private static VBAParser.EventArgumentContext ImmediateEventArgumentExpressionContext(IdentifierReference reference)
        {
            var context = reference.Context;
            //The context is either already a simpleNameExprContext or an IdentifierValueContext used in a sub-rule of some other lExpression alternative. 
            var lExpressionNameContext = context is VBAParser.SimpleNameExprContext simpleName
                ? simpleName
                : context.GetAncestor<VBAParser.LExpressionContext>();

            //To be an immediate argument and, thus, assignable by ref, the structure must be argumentExpression -> expression -> lExpression.
            return lExpressionNameContext?
                .Parent?
                .Parent as VBAParser.EventArgumentContext;
        }

        private static bool AllImplementationParametersCanBeChangedToByVal(ParameterDeclaration parameter, ModuleBodyElementDeclaration interfaceMember, DeclarationFinder finder)
        {
            if (!TryFindParameterIndex(parameter, interfaceMember, out var parameterIndex))
            {
                //This really should never happen.
                Debug.Fail($"Could not find index for parameter {parameter.IdentifierName} in interface member {interfaceMember.IdentifierName}.");
                return false;
            }

            var implementations = finder.FindInterfaceImplementationMembers(interfaceMember);
            return implementations.All(implementation => ParameterAtIndexCanBeChangedToBePassedByValIfRelatedParameterCan(implementation, parameterIndex, finder));
        }

        private static bool TryFindParameterIndex(ParameterDeclaration parameter, IParameterizedDeclaration enclosingMember, out int parameterIndex)
        {
            parameterIndex = enclosingMember.Parameters
                .ToList()
                .IndexOf(parameter);
            return parameterIndex != -1;
        }

        private static bool ParameterAtIndexCanBeChangedToBePassedByValIfRelatedParameterCan(IParameterizedDeclaration member, int parameterIndex, DeclarationFinder finder)
        {
            var parameter = member.Parameters.ElementAtOrDefault(parameterIndex);
            return parameter != null 
                   && CanBeChangedToBePassedByValIfRelatedParameterCan(parameter, finder);
        }

        private static bool CanBeChangedToBePassedByValIfRelatedParameterCan(ParameterDeclaration parameter, DeclarationFinder finder)
        {
            return !parameter.References
                        .Any(reference => reference.IsAssignment)
                    && !IsPotentiallyUsedAsByRefParameter(parameter, finder);
        }

        private static bool AllHandlerParametersCanBeChangedToByVal(ParameterDeclaration parameter, EventDeclaration eventDeclaration, DeclarationFinder finder)
        {
            if (!TryFindParameterIndex(parameter, eventDeclaration, out var parameterIndex))
            {
                //This really should never happen.
                Debug.Fail($"Could not find index for parameter {parameter.IdentifierName} in event {eventDeclaration.IdentifierName}.");
                return false;
            }

            if (!eventDeclaration.IsUserDefined)
            {
                return false;
            }

            var handlers = finder.FindEventHandlers(eventDeclaration);
            return handlers.All(handler => ParameterAtIndexCanBeChangedToBePassedByValIfRelatedParameterCan(handler, parameterIndex, finder));
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.ParameterCanBeByValInspection, declaration.IdentifierName);
        }
    }
}
